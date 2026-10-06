import type {
  CellEdit, CellInputEdit, EditResult, Selection, WorkbookHandle, WorkbookEditPeer, WorkbookReplayOp, XlsxEditResult, XlsxReadRequest, XlsxReadResult,
} from '@betteroffice/xlsx';
import {
  WorkbookEditPeerFailedError, WorkbookPeerHydrationError, WorkbookRecoveryRefusal, workbookEditPeerOperations,
} from '@betteroffice/xlsx';
import type { XlsxWorkerViewerApi } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import {
  WorkerInputNotReadyError, WorkerInputRefusal, inputRefusal, type WorkerInputCoordinator, type WorkerInputLease,
} from '../commands/workerInputCoordinator';

/** @experimental */
export class XlsxPeerNotReadyError extends WorkerInputNotReadyError {
  constructor() {
    super();
    this.name = 'XlsxPeerNotReadyError';
  }
}

/** @experimental */
export class XlsxWorkerEditorCollaborationError extends Error {
  readonly code = 'collaboration-unavailable';
  constructor() {
    super('Collaboration is unavailable in the worker editor');
    this.name = 'XlsxWorkerEditorCollaborationError';
  }
}

export interface WorkerEditorSessionAccess {
  readonly current: boolean;
  readonly ready: boolean;
  readonly peer: WorkbookHandle | null;
  readonly editPeer: WorkbookEditPeer | null;
  readonly failure: Error | null;
  readonly recovering?: boolean;
  readonly retiring?: boolean;
  whenHydrated(): Promise<void>;
  requestHydration(reason: string): Promise<void>;
}

export interface WorkerEditorApiBridge {
  coordinator(): WorkerInputCoordinator | null;
  readOnly(): boolean;
  clearSelection(): void;
  focus(): void;
  refreshProposals(): void;
  recoverInput(): Promise<void>;
  selectCells(sheet: number, selection: Selection): boolean;
  selectCellsAsync(sheet: number, selection: Selection): Promise<boolean>;
  previewEdits?(sheet: number, edits: readonly CellInputEdit[], op?: WorkbookReplayOp): Promise<void | (() => Promise<void>)>;
  canNavigateSync?(): boolean;
  apply(result: EditResult | XlsxEditResult, op?: WorkbookReplayOp): void;
}

/** @experimental */
export interface XlsxWorkerEditorApi extends Omit<XlsxWorkerViewerApi, 'save' | 'selectCells'> {
  readonly hydrated: boolean;
  readonly failure: Error | null;
  whenHydrated(): Promise<void>;
  flush(): Promise<void>;
  save(): Promise<Uint8Array | null>;
  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }>;
  selectCells(sheet: number, selection: Selection): boolean;
  cell(sheet: number, row: number, col: number): CellEdit | null;
  cellAsync(sheet: number, row: number, col: number): Promise<CellEdit | null>;
  rangeCells(sheet: number, range: string): CellEdit[][];
  rangeCellsAsync(sheet: number, range: string): Promise<CellEdit[][] | null>;
  readCellsSync(request: XlsxReadRequest): XlsxReadResult | null;
  editCell(sheet: number, row: number, col: number, input: string): EditResult;
  editCellAsync(sheet: number, row: number, col: number, input: string): Promise<EditResult | null>;
}

function validSelection(peer: WorkbookHandle, sheet: number, selection: Selection): boolean {
  return Number.isInteger(sheet) && sheet >= 0 && sheet < peer.sheetCount() &&
    [selection?.anchor, selection?.focus].every((cell) => cell &&
      Number.isInteger(cell.row) && cell.row >= 0 && cell.row <= 1_048_575 &&
      Number.isInteger(cell.col) && cell.col >= 0 && cell.col <= 16_383);
}

export function createWorkerEditorApi(
  session: WorkerEditorSessionAccess,
  commands: XlsxCommandStore,
  bridge: () => WorkerEditorApiBridge
): XlsxWorkerEditorApi {
  const coordinator = () => {
    const value = bridge().coordinator();
    if (!value) throw new XlsxCommandAdmissionError('editor-unavailable');
    return value;
  };
  const assertCurrent = () => {
    if (!session.current) throw new XlsxCommandAdmissionError('document-replaced');
  };
  const requirePeer = () => {
    assertCurrent();
    if (session.failure && !session.recovering) throw session.failure;
    if (!session.ready || !session.peer || !session.editPeer) throw new XlsxPeerNotReadyError();
    return { peer: session.peer, edits: session.editPeer };
  };
  const ordered = <T,>(
    reason: string, operation: (markApplied: WorkerInputLease) => T | Promise<T>, mutation = false,
    prepare?: () => void | (() => Promise<void>) | Promise<void | (() => Promise<void>)>
  ): Promise<T | null> => {
    if (!session.current || session.retiring) return Promise.resolve(null);
    if (session.failure) return Promise.reject(session.failure);
    try {
      const hydration = session.requestHydration(reason);
      void hydration.catch(() => {});
      let discardPreview: (() => Promise<void>) | undefined;
      return coordinator().runAfterPendingInput(async (_, markApplied) => {
        await (session.recovering ? session.requestHydration('recovery') : hydration);
        markApplied.check();
        requirePeer();
        if (session.recovering && !mutation) { markApplied(); return null; }
        let value: T;
        try { value = await operation(markApplied); }
        catch (error) {
          if (inputRefusal(error)) await discardPreview?.();
          throw error;
        }
        if (value !== null && typeof value === 'object' &&
          ('ok' in value && value.ok === false || 'error' in value && value.error)) {
          await discardPreview?.();
          markApplied.check();
          const failure = value as { failure?: { message: string }; error?: Error };
          markApplied.refuse(new WorkerInputRefusal(failure.failure?.message ?? failure.error?.message ?? 'Cell operation was refused'), mutation);
        }
        return session.current ? value : null;
      }, { kind: 'host', recover: mutation, barrier: reason === 'save', prepare: prepare ? async () => {
        if (!session.retiring) {
          const cancel = await prepare();
          discardPreview = typeof cancel === 'function' ? cancel : undefined;
        }
      } : undefined });
    } catch (error) { return Promise.reject(error); }
  };
  const synchronous = <T,>(operation: (markApplied: WorkerInputLease) => T): T => {
    if (session.retiring) throw new XlsxCommandAdmissionError('document-replaced');
    if (session.failure) throw session.failure;
    requirePeer();
    try { return coordinator().runSync((_, markApplied) => operation(markApplied), { kind: 'host' }); }
    catch (error) {
      if (error instanceof WorkerInputNotReadyError) throw new XlsxPeerNotReadyError();
      throw error;
    }
  };
  const refusal = () => ({
    ok: false as const, version: requirePeer().peer.version(),
    failure: { code: 'read-only' as const, message: 'The editor is read-only' },
  });
  const editCell = (
    sheet: number, row: number, col: number, input: string, markApplied: WorkerInputLease
  ) => {
    markApplied.check();
    const result = requirePeer().edits.editCell(sheet, row, col, input);
    markApplied();
    if (!session.retiring) bridge().apply(result);
    return result;
  };
  const previewRequest = async (request: Parameters<WorkbookHandle['applyEdits']>[0], op: WorkbookReplayOp) => {
    if (request.steps.some((step) => step.op !== 'setCellInputs')) {
      return bridge().previewEdits?.(0, [], op);
    }
    const discards: (() => Promise<void>)[] = [];
    const discard = async () => { for (const cancel of discards.splice(0).reverse()) await cancel(); };
    try {
      for (const step of request.steps) {
        if (step.op !== 'setCellInputs') continue;
        const sheet = Number(/^sheet:(\d+)$/.exec(step.target.sheetId)?.[1] ?? -1);
        const range = step.target.range;
        let start: { row: number; col: number };
        if (range.kind === 'rowCol') start = range.start;
        else {
          const match = /^\$?([A-Z]+)\$?(\d+)/i.exec(range.a1);
          if (!match) throw new WorkerInputRefusal('Invalid cell range');
          start = { row: Number(match[2]) - 1, col: [...match[1].toUpperCase()].reduce((col, letter) => col * 26 + letter.charCodeAt(0) - 64, 0) - 1 };
        }
        const edits = step.inputs.flatMap((row, dr) => row.map((value, dc) => ({
          row: start.row + dr, col: start.col + dc, input: value,
        })));
        const cancel = await bridge().previewEdits?.(sheet, edits, op);
        if (cancel) discards.push(cancel);
      }
    } catch (error) { await discard(); throw error; }
    return discard;
  };
  const save = () => ordered('save', async () => new Uint8Array(await requirePeer().edits.save()));
  const readable = () => session.current && !session.retiring && session.ready && !session.failure &&
    !bridge().coordinator()?.unapplied.length && !bridge().coordinator()?.draft;

  return {
    handle: null, commands,
    get hydrated() { return session.ready && !session.retiring && !session.failure; },
    get failure() { return session.failure; },
    whenHydrated: () => session.whenHydrated(),
    async flush() {
      assertCurrent();
      await coordinator().flush();
      requirePeer();
    },
    save, saveAsync: save,
    async recoverySave() {
      assertCurrent();
      if (!session.failure && session.editPeer?.state !== 'failed') {
        throw new Error('Recovery requires a failed workbook edit peer');
      }
      try {
        await session.requestHydration('recovery');
        assertCurrent();
        await bridge().recoverInput();
        assertCurrent();
        if (!session.editPeer) throw new XlsxPeerNotReadyError();
        const result = session.editPeer.recoverySave();
        return { bytes: new Uint8Array(result.bytes), recovery: true };
      } catch (error) {
        if (error instanceof XlsxCommandAdmissionError || error instanceof XlsxPeerNotReadyError ||
          error instanceof WorkbookPeerHydrationError || error instanceof WorkbookEditPeerFailedError) throw error;
        throw new WorkbookEditPeerFailedError(error instanceof Error ? error : new Error(String(error)));
      }
    },
    clearSelection: () => { if (session.current && !session.failure) bridge().clearSelection(); },
    focus: () => { if (session.current && !session.failure) bridge().focus(); },
    refreshProposals: () => { if (session.ready) bridge().refreshProposals(); },
    selectCells(sheet, selection) {
      if (session.failure || !session.ready || !session.peer || !validSelection(session.peer, sheet, selection) ||
        bridge().canNavigateSync?.() === false) return false;
      return synchronous((markApplied) => {
        requirePeer().edits.setActiveSheet(sheet);
        markApplied();
        return bridge().selectCells(sheet, selection);
      });
    },
    selectCellsAsync(sheet, selection) {
      const target = structuredClone(selection);
      return ordered('select-cells', async (markApplied) => {
        const { peer, edits } = requirePeer();
        if (!validSelection(peer, sheet, target)) return false;
        markApplied.check();
        edits.setActiveSheet(sheet);
        markApplied();
        return bridge().selectCellsAsync(sheet, target);
      }, true).then((selected) => selected ?? false);
    },
    version: () => ordered('version', () => requirePeer().peer.version()),
    readCells: (request) => {
      const input = structuredClone(request);
      return ordered('read-cells', () => requirePeer().peer.readCells(input));
    },
    findText: (request) => {
      const input = structuredClone(request);
      return ordered('find-text', () => requirePeer().peer.findText(input));
    },
    validateEdits: (request) => {
      const input = structuredClone(request);
      return ordered('validate-edits', () => bridge().readOnly() ? refusal() : requirePeer().peer.validateEdits(input));
    },
    applyEdits: (request) => {
      const editable = !bridge().readOnly();
      const input = structuredClone(request);
      const op: WorkbookReplayOp = { method: 'applyEdits', args: [input] };
      return ordered('apply-edits', (markApplied) => {
        if (!editable) return refusal();
        let result: XlsxEditResult;
        try {
          result = session.recovering || !bridge().previewEdits ? requirePeer().edits.applyEdits(input) :
            workbookEditPeerOperations(requirePeer().edits).applyQueuedOp(op) as XlsxEditResult;
        }
        catch (error) {
          if (!(error instanceof WorkbookRecoveryRefusal)) throw error;
          result = error.result as XlsxEditResult;
        }
        if (result.ok) markApplied();
        if (!session.retiring) bridge().apply(result, op);
        return result;
      }, true, () => editable ? previewRequest(input, op) : undefined);
    },
    cell: (sheet, row, col) => readable() ? session.peer!.cell(sheet, row, col) : null,
    cellAsync: (sheet, row, col) => ordered('cell', () => requirePeer().peer.cell(sheet, row, col)),
    rangeCells: (sheet, range) => readable() ? session.peer!.rangeCells(sheet, range) : [],
    rangeCellsAsync: (sheet, range) => ordered('range-cells', () => requirePeer().peer.rangeCells(sheet, range)),
    readCellsSync: (request) => readable() ? session.peer!.readCells(request) : null,
    editCell(sheet, row, col, input) {
      requirePeer();
      if (bridge().readOnly()) throw new Error('The editor is read-only');
      return synchronous((markApplied) => editCell(sheet, row, col, input, markApplied));
    },
    editCellAsync: (sheet, row, col, input) => {
      const editable = !bridge().readOnly();
      const op: WorkbookReplayOp = { method: 'editCell', args: [sheet, row, col, input] };
      return ordered('edit-cell', (markApplied) => {
        if (!editable) return { error: new Error('The editor is read-only') };
        if (!bridge().previewEdits || session.recovering) return { result: editCell(sheet, row, col, input, markApplied) };
        markApplied.check();
        const result = workbookEditPeerOperations(requirePeer().edits).applyQueuedOp(op) as EditResult;
        markApplied();
        if (!session.retiring) bridge().apply(result, op);
        return { result };
      }, true, () => editable ? bridge().previewEdits?.(sheet, [{ row, col, input }], op) : undefined).then((outcome) => {
        if (outcome?.error) throw outcome.error;
        return outcome?.result ?? null;
      });
    },
  };
}
