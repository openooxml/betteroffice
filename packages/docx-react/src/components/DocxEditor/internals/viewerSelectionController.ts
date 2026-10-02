import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type {
  DocxDisplayRange,
  DocxDisplaySelectionText,
  DocxSelectionUnit,
  ResidentEngineWorkerClient,
} from '@betteroffice/docx/yrs';
import type { WorkerFrameProvenance } from './layoutProvenance';
import { readAt, type ViewerReadOutcome } from './viewerReads';
import { displayListSelectionUnit } from './viewerSelectionUnits';

type Intent = 'caret' | 'range' | DocxSelectionUnit;
interface Selection extends DocxDisplayRange {
  gesture: number;
  revision: number;
  intent: Intent;
  frame: WorkerFrameProvenance;
  phase: 'live' | 'mapping';
  unitPending: boolean;
}

interface Token {
  gesture: number;
  revision?: number;
  version: string;
}

interface Capture {
  revision: number;
  version: string;
  sticky: DocxDisplaySelectionText['sticky'];
  text: string;
}

interface ReadTask<V> {
  token: Token;
  outcome: Promise<ViewerReadOutcome<V>>;
  resolve(outcome: ViewerReadOutcome<V>): void;
  read(): Promise<ViewerReadOutcome<V>>;
  apply(outcome: ViewerReadOutcome<V>): void;
}

class ReadChannel<V> {
  private active: ReadTask<V> | null = null;
  private queued: ReadTask<V> | null = null;

  enqueue(
    token: Token,
    read: ReadTask<V>['read'],
    apply: ReadTask<V>['apply']
  ): ReadTask<V> {
    let resolve!: ReadTask<V>['resolve'];
    const outcome = new Promise<ViewerReadOutcome<V>>((done) => { resolve = done; });
    const task = { token, outcome, resolve, read, apply };
    if (this.active) {
      this.clearQueued();
      this.queued = task;
    } else {
      this.start(task);
    }
    return task;
  }

  clearQueued(): void {
    this.queued?.resolve({ status: 'superseded', version: null });
    this.queued = null;
  }

  private start(task: ReadTask<V>): void {
    this.active = task;
    void task.read().then((outcome) => {
      task.resolve(outcome);
      task.apply(outcome);
      this.active = null;
      const next = this.queued;
      this.queued = null;
      if (next) this.start(next);
    });
  }
}

interface Waiter {
  gesture: number;
  resolve(text: string): void;
  reject(error: Error): void;
  timer: ReturnType<typeof setTimeout>;
}

export interface ViewerSelectionControllerOptions {
  read: ResidentEngineWorkerClient['documentRead'];
  story: string;
  queries(): DisplayListQueries | null;
}

export class ViewerSelectionController {
  private frame: WorkerFrameProvenance | null = null;
  private selection: Selection | null = null;
  private capture: Capture | null = null;
  private captureRead: ReadTask<DocxDisplaySelectionText | null> | null = null;
  private readonly unit = new ReadChannel<DocxDisplayRange | null>();
  private readonly captures = new ReadChannel<DocxDisplaySelectionText | null>();
  private readonly map = new ReadChannel<DocxDisplayRange | null>();
  private readonly listeners = new Set<() => void>();
  private readonly waiters = new Set<Waiter>();
  private gesture = 0;
  private revision = 0;
  private reservedGesture = false;
  private unitPosition = 0;
  private goalX: number | undefined;

  constructor(private readonly options: ViewerSelectionControllerOptions) {}

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  beginGesture(): number {
    this.gesture += 1;
    this.reservedGesture = true;
    this.goalX = undefined;
    this.unit.clearQueued();
    this.captures.clearQueued();
    this.map.clearQueued();
    this.rejectWaiters('Selection gesture changed');
    return this.gesture;
  }

  isCurrent(gesture: number): boolean {
    return gesture === this.gesture;
  }

  currentGesture(): number {
    return this.gesture;
  }

  select(anchor: number, head = anchor, gesture = this.takeGesture()): void {
    if (!this.isCurrent(gesture)) return;
    this.reservedGesture = false;
    if (!this.frame) return;
    this.setRange({ anchor, head }, this.frame, gesture, anchor === head ? 'caret' : 'range', false);
  }

  expand(position: number, intent: 'word' | 'paragraph'): void {
    const gesture = this.takeGesture();
    if (!this.frame) return;
    const queries = this.options.queries();
    const local = queries && this.options.story === 'body'
      ? displayListSelectionUnit(queries, position, intent)
      : null;
    this.unitPosition = position;
    this.setRange(local ?? { anchor: position, head: position }, this.frame, gesture, intent, true);
    this.readUnit(this.unit, position, intent);
  }

  selectAll(): void {
    this.beginGesture();
    const gesture = this.takeGesture();
    if (!this.frame) return;
    this.unitPosition = 0;
    this.setRange({ anchor: 0, head: this.positionLimit() }, this.frame, gesture, 'story', true);
    this.readUnit(this.unit, 0, 'story');
  }

  displaySelection(): DocxDisplayRange | null {
    const selection = this.selection;
    return selection?.phase === 'live' && selection.frame.version === this.frame?.version
      ? { anchor: selection.anchor, head: selection.head }
      : null;
  }

  settledText(gesture = this.gesture): string | null {
    const selection = this.selection;
    const capture = this.capture;
    const frame = this.frame;
    return selection && capture && frame && this.isCurrent(gesture) && selection.gesture === gesture &&
      selection.phase === 'live' && !selection.unitPending &&
      selection.frame.version === frame.version &&
      capture.revision === selection.revision && capture.version === frame.version
      ? capture.text
      : null;
  }

  readSelectedText(): Promise<string> | null {
    const selection = this.selection;
    return selection && (selection.anchor !== selection.head || selection.unitPending || selection.phase === 'mapping')
      ? this.whenSettled(this.gesture)
      : null;
  }

  whenSettled(gesture: number, timeoutMs = 10_000): Promise<string> {
    if (!this.isCurrent(gesture) || this.selection?.gesture !== gesture) {
      return Promise.reject(new Error('Selection gesture changed'));
    }
    const text = this.settledText(gesture);
    if (text !== null) return Promise.resolve(text);
    return new Promise((resolve, reject) => {
      const waiter: Waiter = {
        gesture, resolve, reject,
        timer: setTimeout(() => {
          this.waiters.delete(waiter);
          reject(new Error('Selection did not settle'));
        }, timeoutMs),
      };
      this.waiters.add(waiter);
    });
  }

  onFrame(frame: WorkerFrameProvenance | null): void {
    const previous = this.frame;
    this.frame = frame ? { ...frame } : null;
    const selection = this.selection;
    if (!selection) return;
    if (frame?.version === previous?.version) {
      this.emit();
      return;
    }
    selection.phase = 'mapping';
    this.emit();
    if (!frame || !this.isCurrent(selection.gesture)) return;
    if (selection.frame.version === frame.version) {
      selection.phase = 'live';
      if (!this.capture) this.captureRead = null;
      this.captureLatest();
      if (selection.unitPending && (selection.intent === 'word' || selection.intent === 'paragraph' || selection.intent === 'story')) {
        this.readUnit(this.unit, this.unitPosition, selection.intent);
      }
      this.emit();
    } else if (selection.intent === 'story') {
      this.readUnit(this.map, 0, 'story');
    } else if (selection.frame.preview && !frame.preview) {
      const pending = selection.unitPending;
      this.setRange(selection, frame, selection.gesture, selection.intent, pending);
      if (pending && (selection.intent === 'word' || selection.intent === 'paragraph')) {
        this.readUnit(this.unit, this.unitPosition, selection.intent);
      }
    } else {
      void this.mapCaptured(selection, frame);
    }
  }

  reset(): void {
    this.beginGesture();
    this.reservedGesture = false;
    this.selection = null;
    this.capture = null;
    this.captureRead = null;
    this.frame = null;
    this.emit();
  }

  move(key: string, extend: boolean): boolean {
    const selection = this.selection;
    const queries = this.options.queries();
    if (!selection || !queries?.isReady() || !this.displaySelection() || !this.frame) return false;
    const from = Math.min(selection.anchor, selection.head);
    const to = Math.max(selection.anchor, selection.head);
    let head = selection.head;
    let goalX: number | undefined;
    if ((key === 'ArrowLeft' || key === 'ArrowRight') && !extend && from !== to) {
      head = key === 'ArrowLeft' ? from : to;
    } else if (key === 'ArrowLeft' || key === 'ArrowRight') {
      const step = key === 'ArrowLeft' ? -1 : 1;
      const limit = this.positionLimit();
      for (let next = head + step; next >= 0 && next <= limit; next += step) {
        if (queries.visualLineAtPosition(next)) {
          head = next;
          break;
        }
      }
    } else if (key === 'ArrowUp' || key === 'ArrowDown') {
      const moved = queries.verticalMove(head, key === 'ArrowUp' ? 'up' : 'down', this.goalX);
      if (moved) {
        head = moved.position;
        goalX = moved.goalX;
      }
    } else {
      const line = queries.visualLineAtPosition(head);
      if (line) head = key === 'Home' ? line.from : line.to;
    }
    this.beginGesture();
    const gesture = this.takeGesture();
    this.setRange({ anchor: extend ? selection.anchor : head, head }, this.frame, gesture, extend ? 'range' : 'caret', false);
    this.goalX = goalX;
    return true;
  }

  private takeGesture(): number {
    if (!this.reservedGesture) this.beginGesture();
    this.reservedGesture = false;
    return this.gesture;
  }

  private positionLimit(): number {
    let limit = Math.max(this.selection?.anchor ?? 0, this.selection?.head ?? 0);
    for (const page of this.options.queries()?.displayList.pages ?? []) {
      for (const primitive of page.primitives) {
        limit = Math.max(limit, primitive.docStart ?? 0, primitive.docEnd ?? 0,
          primitive.fragmentDocStart ?? 0, primitive.fragmentDocEnd ?? 0);
      }
    }
    return limit;
  }

  private token(): Token {
    return { gesture: this.selection!.gesture, revision: this.selection!.revision, version: this.frame!.version };
  }

  private applies<V>(token: Token, outcome: ViewerReadOutcome<V>): outcome is Extract<ViewerReadOutcome<V>, { status: 'ok' }> {
    const selection = this.selection;
    return outcome.status === 'ok' && selection !== null && this.isCurrent(token.gesture) &&
      token.gesture === selection.gesture &&
      (token.revision === undefined || token.revision === selection.revision) &&
      outcome.version === token.version && token.version === this.frame?.version;
  }

  private setRange(range: DocxDisplayRange, frame: WorkerFrameProvenance, gesture: number, intent: Intent, unitPending: boolean): void {
    this.selection = { anchor: range.anchor, head: range.head, gesture, revision: ++this.revision,
      intent, frame: { ...frame }, phase: 'live', unitPending };
    this.capture = null;
    this.captureLatest();
    this.emit();
  }

  private readUnit(channel: ReadChannel<DocxDisplayRange | null>, position: number, unit: DocxSelectionUnit): void {
    const token = this.token();
    channel.enqueue(token, () => readAt(this.options.read, {
      kind: 'selectionUnit', story: this.options.story, position, unit, expectVersion: token.version,
    }), (outcome) => {
      if (!this.applies(token, outcome)) return;
      if (!outcome.value) { this.drop(); return; }
      const selection = this.selection!;
      const pending = unit === 'story' && this.frame!.preview;
      if (selection.frame.version === token.version && selection.phase === 'live' &&
        selection.anchor === outcome.value.anchor && selection.head === outcome.value.head) {
        selection.unitPending = pending;
        this.emit();
      } else {
        this.setRange(outcome.value, this.frame!, token.gesture, selection.intent, pending);
      }
    });
  }

  private captureLatest(): void {
    const selection = this.selection;
    if (!selection || selection.phase !== 'live' || selection.frame.version !== this.frame?.version) return;
    if (this.capture?.revision === selection.revision && this.capture.version === this.frame?.version) return;
    const token = this.token();
    if (this.captureRead?.token.revision === token.revision && this.captureRead.token.version === token.version) return;
    const { anchor, head } = selection;
    this.captureRead = this.captures.enqueue(token, () => readAt(this.options.read, {
      kind: 'selectionText', story: this.options.story, anchor, head, expectVersion: token.version,
    }), (outcome) => {
      if (!this.applies(token, outcome)) return;
      if (!outcome.value) { this.drop(); return; }
      this.capture = { revision: token.revision!, version: token.version,
        sticky: outcome.value.sticky, text: outcome.value.text };
      this.emit();
    });
  }

  private async mapCaptured(selection: Selection, frame: WorkerFrameProvenance): Promise<void> {
    let capture = this.capture?.revision === selection.revision ? this.capture : null;
    const pending = this.captureRead;
    if (!capture && pending?.token.revision === selection.revision) {
      const outcome = await pending.outcome;
      if (outcome.status === 'ok' && outcome.value) {
        capture = { revision: selection.revision, version: outcome.version,
          sticky: outcome.value.sticky, text: outcome.value.text };
      }
    }
    const token = { gesture: selection.gesture, revision: selection.revision, version: frame.version };
    if (!this.applies(token, { status: 'ok', version: frame.version, value: null })) return;
    if (!capture?.sticky) { this.drop(); return; }
    const sticky = capture.sticky;
    this.map.enqueue(token, () => readAt(this.options.read, {
      kind: 'stickyPosition', story: this.options.story, anchor: sticky.anchor,
      head: sticky.head, expectVersion: token.version,
    }), (outcome) => {
      if (!this.applies(token, outcome)) return;
      if (!outcome.value) { this.drop(); return; }
      this.setRange(outcome.value, this.frame!, token.gesture, selection.intent, selection.unitPending);
      if (selection.unitPending && (selection.intent === 'word' || selection.intent === 'paragraph')) {
        this.unitPosition = outcome.value.anchor;
        this.readUnit(this.unit, this.unitPosition, selection.intent);
      }
    });
  }

  private drop(): void {
    this.selection = null;
    this.capture = null;
    this.captureRead = null;
    this.rejectWaiters('Selection dropped');
    this.emit();
  }

  private rejectWaiters(message: string): void {
    for (const waiter of this.waiters) {
      clearTimeout(waiter.timer);
      waiter.reject(new Error(message));
    }
    this.waiters.clear();
  }

  private emit(): void {
    const text = this.settledText();
    if (text !== null) {
      for (const waiter of this.waiters) {
        if (!this.isCurrent(waiter.gesture)) continue;
        clearTimeout(waiter.timer);
        this.waiters.delete(waiter);
        waiter.resolve(text);
      }
    }
    for (const listener of this.listeners) listener();
  }
}
