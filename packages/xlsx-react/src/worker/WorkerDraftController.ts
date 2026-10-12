import type { CellAddr, Direction, Selection, WorkbookCommandSelection, XlsxRangeTarget } from '@betteroffice/xlsx';
import type { InputDraft } from '../commands/inputCoordinator';
import type { PeerlessWorkerCommandOptions, PeerlessWorkerCommandTicket, PeerlessWorkerInputCoordinator, PeerlessWorkerInputSeal } from '../commands/peerlessWorkerInputCoordinator';

export type WorkerNativeDraft = InputDraft & { revision: number; modified: boolean;
  after?: PeerlessWorkerCommandOptions['after']; target?: XlsxRangeTarget };

interface RefusedDraft {
  seq: number | null;
  draft: WorkerNativeDraft;
}

interface DraftSnapshot {
  cell: WorkerNativeDraft | null;
  formula: string | null;
  refused: readonly RefusedDraft[];
}

interface DraftControllerOptions {
  generation(): number;
  identity(): string;
  navigation(): number;
  selection(): Selection | null;
  sheet(): number;
  target(sheet: number, row: number, col: number): XlsxRangeTarget;
  resolveSheet(sheetId: string): number | undefined;
  editable(): boolean;
  prefill(sheet: number, cell: CellAddr): Promise<string>;
  select(sheet: number, cell: CellAddr, move?: Direction): void;
  selectionTarget(): WorkbookCommandSelection | undefined;
  dependency(): PeerlessWorkerCommandOptions['after'];
  navigate(): void;
  committed(ticket: PeerlessWorkerCommandTicket, draft: WorkerNativeDraft): void;
  fail(error: unknown): void;
  focus(): void;
}

export class WorkerDraftController {
  private coordinator: PeerlessWorkerInputCoordinator | null = null;
  private listeners = new Set<() => void>();
  private snapshot: DraftSnapshot = { cell: null, formula: null, refused: [] };
  private draft: WorkerNativeDraft | null = null;
  private formulaTarget: WorkerNativeDraft | null = null;
  private formulaThroughSeq: number | null = null;
  private restoredRefusal: InputDraft | null = null;
  private revision = 0;
  private compositionRevision = 0;
  private composing: { id: string; source: InputDraft['source']; target: WorkerNativeDraft;
    done: Promise<boolean>; settle(value: boolean): void; sealed: boolean; completed: WorkerNativeDraft | null } | null = null;
  private trailing: { source: InputDraft['source']; input: HTMLInputElement; value: string } | null = null;
  private conversionEnter = false;
  private disposed = false;
  cellInput: HTMLInputElement | null = null;
  formulaInput: HTMLInputElement | null = null;

  constructor(private readonly options: DraftControllerOptions) {}

  attach(coordinator: PeerlessWorkerInputCoordinator): void { this.coordinator = coordinator; }
  getSnapshot = (): DraftSnapshot => this.snapshot;
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  };
  get compositionId(): string | undefined { return this.composing?.id; }
  get prefillIdentity(): string { return `${this.revision}:${this.compositionRevision}`; }
  isComposing(source: InputDraft['source']): boolean { return this.composing?.source === source; }
  get ownsFocus(): boolean { return this.draft !== null || this.snapshot.cell !== null || this.composing !== null; }

  private publish(patch: Partial<DraftSnapshot>): void {
    this.snapshot = { ...this.snapshot, ...patch };
    for (const listener of this.listeners) listener();
  }

  private setDraft(draft: WorkerNativeDraft | null): void {
    this.draft = draft;
    if (!draft) this.restoredRefusal = null;
    this.coordinator?.setDraft(draft);
  }

  private create(source: InputDraft['source'], cell: CellAddr, value: string, modified = true): WorkerNativeDraft {
    return { generation: this.options.generation(), sheet: this.options.sheet(), ...cell, source, value,
      revision: ++this.revision, modified, after: this.options.dependency(),
      target: this.options.target(this.options.sheet(), cell.row, cell.col) };
  }

  open(seed?: string): void {
    const selected = this.options.selection();
    if (!this.options.editable() || !selected || this.composing || this.snapshot.cell) return;
    this.commit('formula');
    this.trailing = null;
    const draft = this.create('cell', selected.focus, seed ?? '', seed !== undefined);
    this.setDraft(draft);
    this.publish({ cell: draft, formula: null });
    if (seed !== undefined) return;
    const identity = this.options.identity();
    const navigation = this.options.navigation();
    const composition = this.compositionRevision;
    void this.options.prefill(draft.sheet, draft).then((value) => {
      if (this.disposed || this.options.identity() !== identity || this.options.navigation() !== navigation ||
        this.draft?.revision !== draft.revision || this.compositionRevision !== composition || this.composing) return;
      const next = { ...draft, value };
      this.setDraft(next);
      const input = this.cellInput;
      if (input) {
        const start = input.selectionStart;
        const end = input.selectionEnd;
        input.value = value;
        if (start !== null && end !== null) input.setSelectionRange(start, end);
      }
      this.publish({ cell: next });
    }).catch((error) => {
      if (!this.disposed && this.options.identity() === identity && this.options.navigation() === navigation &&
        this.draft?.revision === draft.revision && this.compositionRevision === composition) this.options.fail(error);
    });
  }

  change(source: InputDraft['source'], value: string, input: HTMLInputElement | null): void {
    if (!this.options.editable() || !this.coordinator) return;
    if (this.trailing?.source === source && this.trailing.input === input && this.trailing.value === value) return;
    this.trailing = null;
    const selected = this.options.selection();
    if (!selected) return;
    if (this.draft && this.draft.source !== source && !this.composing) this.commit(this.draft.source);
    const base = this.draft ?? (source === 'cell' ? this.snapshot.cell : null) ??
      this.create(source, selected.focus, value);
    const draft = { ...base, value, modified: true, revision: ++this.revision };
    this.setDraft(draft);
    if (this.composing) return;
    this.publish(source === 'cell' ? { cell: draft, formula: null } : { formula: value, cell: null });
  }

  sync = (): void => {
    const draft = this.draft;
    const input = draft?.source === 'cell' ? this.cellInput : this.formulaInput;
    if (!draft || !input || input.value === draft.value) return;
    const next = { ...draft, value: input.value, modified: true, revision: ++this.revision };
    this.setDraft(next);
    if (!this.composing) this.publish(next.source === 'cell' ? { cell: next } : { formula: next.value });
  };

  commit(source: InputDraft['source'], move?: Direction): void {
    if (!this.options.editable() || this.composing || !this.coordinator) return;
    if (this.trailing?.source === source && this.conversionEnter && (move === 'up' || move === 'down')) {
      this.trailing = null;
      this.conversionEnter = false;
      return;
    }
    if (move) { this.trailing = null; this.conversionEnter = false; }
    this.sync();
    const draft = this.draft?.source === source ? this.draft : null;
    const displayed = source === 'cell' ? this.snapshot.cell : this.formulaTarget;
    if (!draft && !displayed) return;
    if (!draft && source === 'formula' && !move) return;
    const target = draft ?? displayed!;
    this.setDraft(null);
    if (source === 'cell') this.publish({ cell: null });
    else if (draft) {
      this.formulaTarget = { ...draft, modified: false };
      this.formulaThroughSeq = null;
      this.publish({ formula: draft.value });
    }
    const sheet = target.target ? this.options.resolveSheet(target.target.sheetId) : target.sheet;
    if (sheet !== undefined) this.options.select(sheet, target, move);
    if (draft?.modified) {
      const ticket = this.coordinator.submit(draft, { selection: this.options.selectionTarget(), after: draft.after });
      this.options.committed(ticket, draft);
    } else this.options.navigate();
    if (move || source === 'cell') this.options.focus();
  }

  cancel = (): void => {
    this.composing?.settle(false);
    this.composing = null;
    this.compositionRevision += 1;
    this.trailing = null;
    this.formulaTarget = null;
    this.formulaThroughSeq = null;
    if (this.restoredRefusal) this.coordinator?.discardRejected(this.restoredRefusal);
    this.setDraft(null);
    this.publish({ cell: null, formula: null });
    this.options.focus();
  };

  start(source: InputDraft['source']): void {
    if (!this.options.editable() || this.composing) return;
    const selected = this.options.selection();
    if (!selected) return;
    if (this.draft && this.draft.source !== source) this.commit(this.draft.source);
    const input = source === 'cell' ? this.cellInput : this.formulaInput;
    const target = this.draft ?? this.create(source, selected.focus, input?.value ?? '', false);
    this.setDraft(target);
    this.trailing = null;
    this.conversionEnter = false;
    let settle!: (value: boolean) => void;
    const done = new Promise<boolean>((resolve) => { settle = resolve; });
    this.composing = { id: `${this.options.identity()}:composition:${++this.compositionRevision}`,
      source, target, done, settle, sealed: false, completed: null };
  }

  end(): void {
    const composition = this.composing;
    if (!composition) return;
    queueMicrotask(() => { if (this.composing === composition) this.completeComposition(); });
  }

  private completeComposition(): void {
    const composition = this.composing;
    if (!composition || !this.coordinator) return;
    this.sync();
    const input = composition.source === 'cell' ? this.cellInput : this.formulaInput;
    const draft: WorkerNativeDraft = { ...composition.target, value: input?.value ?? this.draft?.value ?? '',
      modified: true, revision: ++this.revision };
    composition.completed = draft.value !== '' ? Object.freeze({ ...draft }) : null;
    this.composing = null;
    this.setDraft(composition.sealed && draft.value !== '' ? draft : null);
    if (input) this.trailing = { source: draft.source, input, value: draft.value };
    if (draft.value !== '') {
      const sheet = draft.target ? this.options.resolveSheet(draft.target.sheetId) : draft.sheet;
      if (!composition.sealed && sheet !== undefined && sheet !== this.options.sheet()) this.options.select(sheet, draft);
      if (draft.source === 'formula') {
        this.formulaTarget = { ...draft, modified: false };
        this.formulaThroughSeq = null;
        this.publish({ formula: draft.value });
      }
      if (!composition.sealed) {
        const ticket = this.coordinator.submit(draft, { compositionId: composition.id,
          selection: this.options.selectionTarget(), after: draft.after });
        this.options.committed(ticket, draft);
      }
      this.publish(draft.source === 'cell' ? { cell: draft } : { formula: draft.value });
    } else this.publish({ cell: null, formula: null });
    composition.settle(draft.value !== '');
  }

  key(source: InputDraft['source'], event: { key: string; keyCode: number; nativeEvent: { isComposing?: boolean } }): boolean {
    if (this.isComposing(source) || event.nativeEvent.isComposing || event.keyCode === 229) {
      if (event.key === 'Enter') this.conversionEnter = true;
      return true;
    }
    if (event.key === 'Enter' && this.conversionEnter) {
      this.conversionEnter = false;
      this.trailing = null;
      return true;
    }
    return false;
  }

  keyUp(): void { if (!this.composing) { this.trailing = null; this.conversionEnter = false; } }

  seal = (): PeerlessWorkerInputSeal => {
    if (!this.composing) {
      if (!this.draft && this.snapshot.cell) this.publish({ cell: null });
      return {};
    }
    const composition = this.composing;
    composition.sealed = true;
    const input = composition.source === 'cell' ? this.cellInput : this.formulaInput;
    input?.blur();
    return { composition: composition.done, compositionDraft: () => composition.completed, restoreFocus: () => {
      if (this.disposed || this.composing || this.draft || !this.options.editable()) return;
      const active = document.activeElement;
      if (active && active !== document.body && active !== input) return;
      if (composition.source === 'formula') this.formulaInput?.focus({ preventScroll: true });
      else this.options.focus();
    } };
  };

  refused(seq: number | null, value: InputDraft): void {
    const draft = { ...value, modified: true, revision: ++this.revision, after: null };
    const record = { seq, draft };
    this.publish({ refused: [...this.snapshot.refused, record].slice(-64) });
    if (!this.ownsFocus && this.options.editable()) this.restore(record);
  }

  restore(record = this.snapshot.refused[0]): void {
    if (!record || this.ownsFocus || !this.options.editable()) return;
    const sheet = record.draft.target ? this.options.resolveSheet(record.draft.target.sheetId) : record.draft.sheet;
    if (sheet === undefined) return;
    const draft = { ...record.draft, sheet };
    this.options.select(sheet, draft);
    this.setDraft(draft);
    this.restoredRefusal = record.draft;
    this.publish({ cell: draft.source === 'cell' ? draft : null,
      formula: draft.source === 'formula' ? draft.value : null,
      refused: this.snapshot.refused.filter((entry) => entry !== record) });
    if (record.draft.source === 'formula') this.formulaInput?.focus({ preventScroll: true });
  }

  finish(): void {
    if (this.composing) { this.sync(); this.end(); return; }
    this.sync();
    const draft = this.draft;
    if (draft?.modified && this.coordinator) this.options.committed(this.coordinator.submit(draft, { after: draft.after }), draft);
    this.setDraft(null);
    this.publish({ cell: null, formula: null });
  }

  selectionChanged(): void {
    this.formulaTarget = null;
    this.formulaThroughSeq = null;
    if (!this.draft && !this.composing && this.snapshot.formula !== null) this.publish({ formula: null });
  }

  submitted(value: InputDraft, seq: number): void {
    if (this.composing) return;
    const submitted = value as WorkerNativeDraft;
    const matches = this.draft?.source === value.source && this.draft.sheet === value.sheet &&
      this.draft.row === value.row && this.draft.col === value.col && this.draft.value === value.value &&
      this.draft.revision === submitted.revision;
    if (matches) this.draft = null;
    if (value.source === 'cell') {
      if (matches) this.publish({ cell: null });
    } else if (matches || this.formulaTarget?.revision === submitted.revision) {
      this.formulaTarget = { ...submitted, modified: false };
      this.formulaThroughSeq = seq;
    }
  }

  presented(seq: number): void {
    if (this.formulaThroughSeq === null || seq < this.formulaThroughSeq || this.draft || this.composing) return;
    this.formulaThroughSeq = null;
    if (this.snapshot.formula !== null) this.publish({ formula: null });
  }

  dispose(): void {
    this.disposed = true;
    this.composing?.settle(false);
    this.composing = null;
    this.listeners.clear();
  }
}
