import type { CellAddr, Direction, Selection, WorkbookCommandSelection, XlsxRangeTarget } from '@betteroffice/xlsx';
import type { InputDraft } from '../commands/inputCoordinator';
import type { PeerlessWorkerCommandOptions, PeerlessWorkerCommandTicket, PeerlessWorkerInputCoordinator, PeerlessWorkerInputSeal, PeerlessWorkerPendingInput } from '../commands/peerlessWorkerInputCoordinator';
import { isMacPlatform } from '../commands/descriptors';
import { cellInputOperation, replayCellInput, type CellInputOperation, type CellInputState } from './cellInputReplay';
import { holdWorkerCellInputRead } from './WorkerCellInputReader';

export type WorkerNativeDraft = InputDraft & { revision: number; modified: boolean;
  after?: PeerlessWorkerCommandOptions['after']; target?: XlsxRangeTarget; reserved?: boolean };

interface RefusedDraft {
  seq: number | null;
  draft: WorkerNativeDraft;
}

interface DraftSnapshot {
  cell: WorkerNativeDraft | null;
  formula: string | null;
  refused: readonly RefusedDraft[];
}

interface CellPrefill {
  draft: WorkerNativeDraft;
  identity: string;
  navigation: number;
  operations: CellInputOperation[];
  nativeSelection?: { anchor: number; focus: number };
  cancelled: boolean;
  sealed: boolean;
  value?: string;
  completed: WorkerNativeDraft | null;
  input: PeerlessWorkerPendingInput;
  settle(value: boolean): void;
}

interface DraftControllerOptions {
  generation(): number;
  identity(): string;
  valueRevision?(): string;
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
  report(error: unknown): void;
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
    done: Promise<boolean>; settle(value: boolean): void; sealed: boolean; completed: WorkerNativeDraft | null;
    loading?: CellPrefill; inputState?: CellInputState; value?: string; result?: string } | null = null;
  private trailing: { source: InputDraft['source']; input: HTMLInputElement; value: string;
    prefill?: boolean; data?: string; replay?: CellInputState } | null = null;
  private conversionEnter = false;
  private disposed = false;
  private loading: CellPrefill | null = null;
  private cellSelection: { revision: number; anchor: number; focus: number } | null = null;
  cellInput: HTMLInputElement | null = null;
  formulaInput: HTMLInputElement | null = null;
  focusedCellInput: { identity: string; value: string } | null = null;

  constructor(private readonly options: DraftControllerOptions) {}

  attach(coordinator: PeerlessWorkerInputCoordinator): void { this.coordinator = coordinator; }
  getSnapshot = (): DraftSnapshot => this.snapshot;
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  };
  get compositionId(): string | undefined { return this.composing?.id; }
  get loadingCell(): boolean { return this.loading !== null; }
  get prefillIdentity(): string { return `${this.revision}:${this.compositionRevision}`; }
  get focusedCellIdentity(): string {
    const cell = this.options.selection()?.focus;
    return JSON.stringify([this.options.identity(), this.options.navigation(), this.options.sheet(),
      cell?.row, cell?.col, this.options.valueRevision?.()]);
  }
  isComposing(source: InputDraft['source']): boolean { return this.composing?.source === source; }
  get ownsFocus(): boolean { return this.draft !== null || this.snapshot.cell !== null || this.composing !== null; }

  private publish(patch: Partial<DraftSnapshot>): void {
    this.snapshot = { ...this.snapshot, ...patch };
    for (const listener of this.listeners) listener();
  }

  private setDraft(draft: WorkerNativeDraft | null): void {
    this.draft = draft;
    if (!draft) this.cellSelection = null;
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
    if (!this.coordinator || !this.options.editable() || !selected || this.composing || this.snapshot.cell) return;
    this.commit('formula');
    this.trailing = null;
    const cached = this.focusedCellInput?.identity === this.focusedCellIdentity ? this.focusedCellInput.value : undefined;
    const preview = this.coordinator.pending ? undefined : cached;
    const draft = this.create('cell', selected.focus, seed ?? preview ?? '', seed !== undefined);
    let settle!: (value: boolean) => void;
    const ready = new Promise<boolean>((done) => { settle = done; });
    const loading: CellPrefill | null = seed === undefined ? {
      draft, identity: this.options.identity(), navigation: this.options.navigation(),
      operations: [], cancelled: false, sealed: false, completed: null, settle,
      input: { ready, target: draft, draft: () => loading!.completed },
    } : null;
    const reading = loading ? holdWorkerCellInputRead(this.coordinator, () => this.options.prefill(draft.sheet, draft),
      (value) => { loading.value = value; this.completePrefill(loading); }, loading.input.ready, () => loading.cancelled,
      (error) => this.refusePrefill(loading, error)) : null;
    this.loading = loading;
    this.setDraft(draft);
    this.publish({ cell: draft, formula: null });
    if (!loading || !reading) return;
    void reading.catch((error) => {
      loading.settle(false);
      if (!loading.cancelled && !this.disposed && this.options.identity() === loading.identity) this.options.fail(error);
    });
  }

  private completePrefill(loading: CellPrefill): void {
    if (loading.cancelled || this.disposed || this.options.identity() !== loading.identity ||
      !loading.sealed && (this.loading !== loading || this.options.navigation() !== loading.navigation)) {
      loading.settle(false);
      return;
    }
    if (loading.completed || loading.value === undefined || this.composing?.loading === loading) return;
    if (this.loading === loading) this.captureCellSelection();
    const replay = replayCellInput(loading.value, loading.operations);
    const modified = replay.value !== loading.value || loading.operations.some((operation) => operation.kind === 'insert' || operation.kind === 'delete');
    const draft = Object.freeze({ ...loading.draft, value: replay.value, modified, reserved: loading.sealed,
      revision: modified ? ++this.revision : loading.draft.revision });
    loading.completed = draft;
    if (!loading.sealed) {
      this.loading = null;
      this.setDraft(draft);
      if (this.trailing?.prefill && this.trailing.input === this.cellInput) this.trailing.replay = replay;
      this.writeCellSelection(draft, replay);
      this.publish({ cell: draft });
    }
    loading.settle(true);
  }

  private refusePrefill(loading: CellPrefill, error: unknown): void {
    loading.settle(false);
    if (loading.cancelled || this.disposed || this.options.identity() !== loading.identity) return;
    if (this.loading === loading) {
      this.cancelPrefill();
      this.setDraft(null);
      this.publish({ cell: null });
    }
    if (!this.coordinator?.error) this.options.report(error);
  }

  private cancelPrefill(): void {
    if (this.loading) { this.loading.cancelled = true; this.loading.settle(false); }
    this.loading = null;
  }

  private sealPrefill(loading: CellPrefill): PeerlessWorkerInputSeal {
    const input = this.cellInput;
    const wasSealed = loading.sealed;
    loading.sealed = true;
    if (this.composing?.loading === loading) {
      if (!wasSealed && input && input.ownerDocument?.activeElement === input) input.blur();
    }
    else if (this.loading === loading) {
      this.loading = null;
      this.setDraft(null);
      this.publish({ cell: null });
    }
    return { pendingInput: loading.input, restoreFocus: () => {
      if (!input || this.disposed || this.composing || this.draft || !this.options.editable()) return;
      const selected = this.options.selection()?.focus;
      const sheet = loading.draft.target ? this.options.resolveSheet(loading.draft.target.sheetId) : loading.draft.sheet;
      if (this.options.sheet() !== sheet || selected?.row !== loading.draft.row || selected?.col !== loading.draft.col) return;
      const active = input.ownerDocument.activeElement;
      if (active && active !== input.ownerDocument.body && active !== input) return;
      this.options.focus();
    } };
  }

  private reservePrefill(loading: CellPrefill): void {
    this.sealPrefill(loading);
    if (this.coordinator && !loading.input.ticket) {
      loading.input.ticket = this.coordinator.reserveInput(loading.input);
      this.options.committed(loading.input.ticket, loading.draft);
    }
  }

  private writeCellSelection(draft: WorkerNativeDraft, replay: CellInputState): void {
    const input = this.cellInput;
    this.cellSelection = input && input.ownerDocument.activeElement === input ? null :
      { revision: draft.revision, anchor: replay.anchor, focus: replay.focus };
    if (input) {
      input.value = replay.value;
      input.setSelectionRange(Math.min(replay.anchor, replay.focus), Math.max(replay.anchor, replay.focus),
        replay.anchor > replay.focus ? 'backward' : 'forward');
    }
  }

  restoreCellSelection(): void {
    const selection = this.cellSelection;
    const input = this.cellInput;
    if (!selection || !input) return;
    this.cellSelection = null;
    if (this.draft?.revision === selection.revision) input.setSelectionRange(
      Math.min(selection.anchor, selection.focus), Math.max(selection.anchor, selection.focus),
      selection.anchor > selection.focus ? 'backward' : 'forward');
  }

  holdCellEdit(operation: CellInputOperation): boolean {
    if (!this.loading) return false;
    if (!this.loading.sealed) this.loading.operations.push(operation);
    return true;
  }

  captureCellSelection(): void {
    const loading = this.loading;
    const input = this.cellInput;
    if (!loading || !input || this.composing?.loading === loading || !input.value.length) return;
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? start;
    const selection = input.selectionDirection === 'backward' ? { anchor: end, focus: start } : { anchor: start, focus: end };
    const previous = loading.nativeSelection;
    if (previous?.anchor === selection.anchor && previous.focus === selection.focus) return;
    loading.nativeSelection = selection;
    if (start === 0 && end === input.value.length) this.holdCellEdit({ kind: 'selectAll' });
    else if (start === end && end === input.value.length) {
      if (previous) this.holdCellEdit({ kind: 'move', direction: 'end' });
    } else this.holdCellEdit({ kind: 'select', ...selection });
  }

  holdCellBeforeInput(event: Pick<InputEvent, 'inputType' | 'data' | 'dataTransfer' | 'isComposing'> &
    Partial<Pick<InputEvent, 'getTargetRanges'>>): boolean {
    if (this.trailing?.prefill && !this.composing &&
      ['insertText', 'insertFromComposition', 'insertCompositionText'].includes(event.inputType) && event.data === this.trailing.data) {
      return this.trailing.data !== '';
    }
    if (!this.loading || this.composing?.loading === this.loading) return false;
    this.captureCellSelection();
    const input = this.cellInput;
    if (event.inputType === 'insertReplacementText' && input?.value.length) {
      const range = event.getTargetRanges?.()[0];
      if (range) this.holdCellEdit(range.startOffset === 0 && range.endOffset === input.value.length
        ? { kind: 'selectAll' } : { kind: 'select', anchor: range.startOffset, focus: range.endOffset });
    }
    const operation = cellInputOperation(event, isMacPlatform());
    return operation ? this.holdCellEdit(operation) : false;
  }

  holdCellKey(event: { key: string; ctrlKey: boolean; metaKey: boolean; altKey: boolean; shiftKey: boolean }): boolean {
    const loading = this.loading;
    if (!loading) return false;
    this.captureCellSelection();
    const operation = cellInputOperation(event, isMacPlatform());
    if (!operation || operation.kind === 'insert') return false;
    this.holdCellEdit(operation);
    const input = this.cellInput;
    if (input) {
      const replay = replayCellInput(input.value, [operation], {
        anchor: input.selectionDirection === 'backward' ? input.selectionEnd ?? 0 : input.selectionStart ?? 0,
        focus: input.selectionDirection === 'backward' ? input.selectionStart ?? 0 : input.selectionEnd ?? 0,
      });
      input.setSelectionRange(Math.min(replay.anchor, replay.focus), Math.max(replay.anchor, replay.focus),
        replay.anchor > replay.focus ? 'backward' : 'forward');
      loading.nativeSelection = { anchor: replay.anchor, focus: replay.focus };
    }
    return true;
  }

  cellKey(event: { key: string; ctrlKey: boolean; metaKey: boolean; altKey: boolean; shiftKey: boolean }): boolean {
    const operation = cellInputOperation(event, isMacPlatform());
    const draft = this.snapshot.cell;
    if (!operation || !draft) return false;
    if (this.holdCellEdit(operation)) return true;
    const selection = this.cellSelection?.revision === draft.revision ? this.cellSelection : undefined;
    const replay = replayCellInput(draft.value, [operation], selection);
    if (operation.kind === 'insert' || operation.kind === 'delete') this.change('cell', replay.value, null);
    this.writeCellSelection(this.draft ?? draft, replay);
    return true;
  }

  change(source: InputDraft['source'], value: string, input: HTMLInputElement | null): void {
    if (!this.options.editable() || !this.coordinator) return;
    if (this.trailing?.source === source && this.trailing.input === input && this.trailing.value === value) {
      if (this.trailing.replay && this.draft?.value === this.trailing.replay.value) this.writeCellSelection(this.draft, this.trailing.replay);
      return;
    }
    if (source === 'cell' && this.composing?.loading) { this.composing.value = value; return; }
    if (source === 'cell' && this.loading) this.cancelPrefill();
    this.cellSelection = null;
    this.trailing = null;
    const selected = this.options.selection();
    if (!selected) return;
    if (this.draft && this.draft.source !== source && !this.composing) {
      this.commit(this.draft.source);
    }
    const base = this.draft ?? (source === 'cell' ? this.snapshot.cell : null) ??
      this.create(source, selected.focus, value);
    const draft = { ...base, value, modified: true, revision: ++this.revision };
    this.setDraft(draft);
    if (this.composing) return;
    this.publish(source === 'cell' ? { cell: draft, formula: null } : { formula: value, cell: null });
  }

  sync = (): void => {
    if (this.loading) return;
    const draft = this.draft;
    const input = draft?.source === 'cell' ? this.cellInput : this.formulaInput;
    if (!draft || !input || input.value === draft.value) return;
    const next = { ...draft, value: input.value, modified: true, revision: ++this.revision };
    this.setDraft(next);
    if (!this.composing) this.publish(next.source === 'cell' ? { cell: next } : { formula: next.value });
  };

  commit(source: InputDraft['source'], move?: Direction): void {
    if (!this.options.editable() || !this.coordinator) return;
    if (this.composing) {
      const loading = this.composing.loading;
      if (source === 'cell' && !move && loading && !loading.sealed) this.reservePrefill(loading);
      return;
    }
    if (source === 'cell' && this.loading) {
      const loading = this.loading;
      if (loading.operations.length) {
        this.reservePrefill(loading);
        if (move) {
          const sheet = loading.draft.target ? this.options.resolveSheet(loading.draft.target.sheetId) : loading.draft.sheet;
          if (sheet !== undefined) this.options.select(sheet, loading.draft, move);
          loading.input.options = { selection: this.options.selectionTarget(), selectionRevision: this.options.navigation() };
          if (!loading.operations.some((operation) => operation.kind === 'insert' || operation.kind === 'delete')) this.options.navigate();
          this.options.focus();
        }
        return;
      }
      this.cancelPrefill();
    }
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
    this.cancelPrefill();
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
    if (source === 'cell') this.captureCellSelection();
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
    const loading = source === 'cell' ? this.loading ?? undefined : undefined;
    this.composing = { id: `${this.options.identity()}:composition:${++this.compositionRevision}`,
      source, target, done, settle, sealed: false, completed: null, loading,
      inputState: loading ? { value: input?.value ?? '', anchor: input?.selectionStart ?? 0, focus: input?.selectionEnd ?? 0 } : undefined };
  }

  end(result?: string): void {
    const composition = this.composing;
    if (!composition) return;
    if (result !== undefined) composition.result = result;
    queueMicrotask(() => { if (this.composing === composition) this.completeComposition(); });
  }

  private completeComposition(): void {
    const composition = this.composing;
    if (!composition || !this.coordinator) return;
    if (composition.loading) {
      const loading = composition.loading;
      const input = this.cellInput;
      const before = composition.inputState!;
      const actual = input?.value ?? composition.value ?? '';
      const start = Math.min(before.anchor, before.focus);
      const tail = before.value.length - Math.max(before.anchor, before.focus);
      const value = composition.result ?? (actual === before.value ? '' : actual.slice(start, Math.max(start, actual.length - tail)));
      const native = value === '' ? before : replayCellInput(before.value, [{ kind: 'insert', text: value }], before);
      if (value !== '') loading.operations.push({ kind: 'insert', text: value });
      if (input) this.trailing = { source: 'cell', input, value: native.value, data: value, prefill: true };
      this.composing = null;
      if (input) { input.value = native.value; input.setSelectionRange(Math.min(native.anchor, native.focus), Math.max(native.anchor, native.focus)); }
      if (loading.sealed) {
        this.loading = null;
        this.setDraft(null);
        this.publish({ cell: null });
      }
      this.completePrefill(loading);
      composition.settle(value !== '');
      return;
    }
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
    if (this.trailing?.prefill) this.trailing = null;
    return false;
  }

  keyUp(): void { if (!this.composing) { this.trailing = null; this.conversionEnter = false; } }

  seal = (): PeerlessWorkerInputSeal => {
    if (this.loading && (this.loading.operations.length || this.composing?.loading === this.loading)) return this.sealPrefill(this.loading);
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
    const draft: WorkerNativeDraft = { ...value, modified: true, revision: ++this.revision, after: null };
    const record = { seq, draft };
    this.publish({ refused: [...this.snapshot.refused, record].slice(-64) });
    const selected = this.options.selection()?.focus;
    const sheet = draft.target ? this.options.resolveSheet(draft.target.sheetId) : draft.sheet;
    const current = this.options.sheet() === sheet && selected?.row === draft.row && selected?.col === draft.col;
    if (!this.ownsFocus && this.options.editable() && (!(value as WorkerNativeDraft).reserved || current)) this.restore(record);
    else this.coordinator?.setDraft(this.draft);
  }

  restore(record = this.snapshot.refused[0]): void {
    if (!record || this.ownsFocus || !this.options.editable()) return;
    const sheet = record.draft.target ? this.options.resolveSheet(record.draft.target.sheetId) : record.draft.sheet;
    if (sheet === undefined) return;
    const draft = { ...record.draft, sheet, reserved: false };
    this.options.select(sheet, draft);
    this.setDraft(draft);
    this.restoredRefusal = record.draft;
    this.publish({ cell: draft.source === 'cell' ? draft : null,
      formula: draft.source === 'formula' ? draft.value : null,
      refused: this.snapshot.refused.filter((entry) => entry !== record) });
    if (record.draft.source === 'formula') this.formulaInput?.focus({ preventScroll: true });
  }

  finish(): void {
    if (this.loading) {
      if (this.loading.operations.length || this.composing?.loading === this.loading) {
        this.reservePrefill(this.loading);
        if (this.composing) this.end();
      } else {
        this.cancelPrefill();
        this.setDraft(null);
        this.publish({ cell: null, formula: null });
      }
      return;
    }
    if (this.composing) { this.sync(); this.end(); return; }
    this.sync();
    const draft = this.draft;
    if (draft?.modified && this.coordinator) this.options.committed(this.coordinator.submit(draft, { after: draft.after }), draft);
    this.setDraft(null);
    this.publish({ cell: null, formula: null });
  }

  selectionChanged(): void {
    if (this.loading && this.options.navigation() !== this.loading.navigation) {
      if (this.loading.operations.length || this.composing?.loading === this.loading) this.reservePrefill(this.loading);
      else {
        this.cancelPrefill();
        this.setDraft(null);
        this.publish({ cell: null });
      }
    }
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
    this.cancelPrefill();
    this.composing?.settle(false);
    this.composing = null;
    this.listeners.clear();
  }
}
