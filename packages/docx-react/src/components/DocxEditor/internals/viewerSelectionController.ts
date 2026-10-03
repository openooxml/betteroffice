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
  unitPending: boolean;
}

interface Token {
  gesture: number;
  revision?: number;
  version: string;
}

interface Capture extends DocxDisplaySelectionText {
  revision: number;
  version: string;
}

export interface ViewerSelectionChange {
  displayRange: { story: string; from: number; to: number; layoutId: string } | null;
  isMultiParagraph: boolean;
}

interface ReadTask<V> {
  token: Token;
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
    const task = { token, read, apply };
    if (this.active) {
      this.queued = task;
    } else {
      this.start(task);
    }
    return task;
  }

  clearQueued(): void {
    this.queued = null;
  }

  private start(task: ReadTask<V>): void {
    this.active = task;
    void task.read().then((outcome) => {
      try {
        task.apply(outcome);
      } finally {
        this.active = null;
        const next = this.queued;
        this.queued = null;
        if (next) this.start(next);
      }
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
    this.rejectWaiters('Selection gesture changed');
    if (this.selection) this.emit();
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
    const selection = this.liveSelection();
    return selection && selection.frame.version === this.frame?.version
      ? { anchor: selection.anchor, head: selection.head }
      : null;
  }

  settledText(gesture = this.gesture): string | null {
    return this.settledCapture(gesture)?.text ?? null;
  }

  settledCapture(gesture = this.gesture): DocxDisplaySelectionText | null {
    const selection = this.selection;
    const capture = this.capture;
    const frame = this.frame;
    return selection && capture && frame && this.isCurrent(gesture) && selection.gesture === gesture &&
      !selection.unitPending && selection.frame.version === frame.version &&
      capture.revision === selection.revision && capture.version === frame.version
      ? { text: capture.text, range: capture.range }
      : null;
  }

  readSelectedText(): Promise<string> | null {
    const selection = this.liveSelection();
    return selection && (selection.anchor !== selection.head || selection.unitPending)
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

  /**
   * Follows the presented frame. A selection lives in one document version: a frame at another
   * version clears it, except the hand-over from a preview to the whole document as opened, whose
   * display positions are the preview's.
   */
  onFrame(frame: WorkerFrameProvenance | null): void {
    const previous = this.frame;
    this.frame = frame ? { ...frame } : null;
    const selection = this.selection;
    if (!selection || !frame || frame.version === previous?.version) {
      this.emit();
    } else if (selection.frame.version === frame.version) {
      if (!this.capture) this.captureRead = null;
      this.captureLatest();
      this.readPendingUnit(selection);
      this.emit();
    } else if (selection.frame.preview && !frame.preview && frame.asOpened) {
      this.setRange(selection, frame, selection.gesture, selection.intent,
        selection.unitPending || selection.intent === 'story');
      this.readPendingUnit(this.selection!);
    } else {
      this.clear();
    }
  }

  reset(): void {
    this.frame = null;
    this.clear();
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
      head = horizontalStep(queries, head, key === 'ArrowLeft' ? -1 : 1) ?? head;
    } else if (key === 'ArrowUp' || key === 'ArrowDown') {
      const line = queries.visualLineAtPosition(head);
      const moved = line && queries.verticalMove(head, key === 'ArrowUp' ? 'up' : 'down', this.goalX);
      const target = moved && queries.visualLineAtPosition(moved.position);
      if (line && moved && target && Math.abs(target.pageIndex - line.pageIndex) <= 1) {
        head = moved.position;
        goalX = moved.goalX;
      }
    } else {
      const line = queries.visualLineAtPosition(head);
      if (line) head = key === 'Home' ? line.from : line.to;
    }
    if (head === selection.head && (extend || from === to)) return true;
    this.beginGesture();
    const gesture = this.takeGesture();
    this.setRange({ anchor: extend ? selection.anchor : head, head }, this.frame, gesture, extend ? 'range' : 'caret', false);
    this.goalX = goalX;
    return true;
  }

  /** The selection, unless a newer gesture that has not selected yet replaced it. */
  private liveSelection(): Selection | null {
    const selection = this.selection;
    return selection && !(this.reservedGesture && selection.gesture !== this.gesture) ? selection : null;
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
      intent, frame: { ...frame }, unitPending };
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
      const selection = this.selection!;
      if (!outcome.value) {
        if (unit === 'story') {
          this.drop();
        } else {
          selection.unitPending = false;
          this.emit();
        }
        return;
      }
      const pending = unit === 'story' && this.frame!.preview;
      if (selection.anchor === outcome.value.anchor && selection.head === outcome.value.head) {
        selection.unitPending = pending;
        this.emit();
      } else {
        this.setRange(outcome.value, this.frame!, token.gesture, selection.intent, pending);
      }
    });
  }

  private captureLatest(): void {
    const selection = this.selection;
    if (!selection || selection.frame.version !== this.frame?.version) return;
    if (this.capture?.revision === selection.revision && this.capture.version === this.frame?.version) return;
    const token = this.token();
    const pending = this.captureRead?.token;
    if (pending?.revision === token.revision && pending?.version === token.version) return;
    const { anchor, head } = selection;
    this.captureRead = this.captures.enqueue(token, () => readAt(this.options.read, {
      kind: 'selectionText', story: this.options.story, anchor, head, expectVersion: token.version,
    }), (outcome) => {
      if (!this.applies(token, outcome)) return;
      this.capture = { revision: token.revision!, version: token.version,
        text: outcome.value?.text ?? '', range: outcome.value?.range ?? null };
      this.emit();
    });
  }

  private readPendingUnit(selection: Selection): void {
    if (selection.unitPending && selection.intent !== 'caret' && selection.intent !== 'range') {
      this.readUnit(this.unit, this.unitPosition, selection.intent);
    }
  }

  private clear(): void {
    this.gesture += 1;
    this.reservedGesture = false;
    this.goalX = undefined;
    this.unit.clearQueued();
    this.captures.clearQueued();
    this.selection = null;
    this.capture = null;
    this.captureRead = null;
    this.rejectWaiters('Selection cleared');
    this.emit();
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

/** The next caret position from `head` on its page or the next one; null past the built pages. */
function horizontalStep(queries: DisplayListQueries, head: number, step: -1 | 1): number | null {
  const line = queries.visualLineAtPosition(head);
  if (!line) return null;
  const pages = [line.pageIndex, line.pageIndex + step];
  let bound = step > 0 ? line.to : line.from;
  for (const pageIndex of pages) {
    for (const candidate of queries.visualLinesOnPage(pageIndex)) {
      bound = step > 0 ? Math.max(bound, candidate.to) : Math.min(bound, candidate.from);
    }
  }
  for (let next = head + step; step > 0 ? next <= bound : next >= bound; next += step) {
    const target = queries.visualLineAtPosition(next);
    if (target && pages.includes(target.pageIndex)) return next;
  }
  return null;
}
