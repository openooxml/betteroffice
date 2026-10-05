import type {
  CellEdit, EditResult, Selection, WorkbookHandle, XlsxEditResult, XlsxReadRequest, XlsxReadResult,
} from '@betteroffice/xlsx';
import type { WorkbookEditPeer } from '../../../xlsx/src/session/editPeer';
import type { XlsxWorkerViewerApi } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import {
  WorkerInputNotReadyError, type WorkerInputCoordinator,
} from '../commands/workerInputCoordinator';

export class XlsxPeerNotReadyError extends WorkerInputNotReadyError {
  constructor() {
    super();
    this.name = 'XlsxPeerNotReadyError';
  }
}

export interface WorkerEditorSessionAccess {
  readonly current: boolean;
  readonly ready: boolean;
  readonly peer: WorkbookHandle | null;
  readonly editPeer: WorkbookEditPeer | null;
  readonly failure: Error | null;
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
  apply(result: EditResult | XlsxEditResult): void;
}

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
    if (session.failure) throw session.failure;
    if (!session.ready || !session.peer || !session.editPeer) throw new XlsxPeerNotReadyError();
    return { peer: session.peer, edits: session.editPeer };
  };
  const ordered = <T,>(
    reason: string, operation: (markApplied: () => void) => T | Promise<T>
  ): Promise<T | null> => {
    if (!session.current) return Promise.resolve(null);
    try {
      const hydration = session.requestHydration(reason);
      void hydration.catch(() => {});
      return coordinator().runAfterPendingInput(async (_, markApplied) => {
        await hydration;
        requirePeer();
        const result = operation(markApplied);
        markApplied();
        const value = await result;
        return session.current ? value : null;
      }, { kind: 'host' });
    } catch (error) { return Promise.reject(error); }
  };
  const synchronous = <T,>(operation: (markApplied: () => void) => T): T => {
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
    sheet: number, row: number, col: number, input: string, markApplied: () => void
  ) => {
    const result = requirePeer().edits.editCell(sheet, row, col, input);
    markApplied();
    bridge().apply(result);
    return result;
  };
  const save = () => ordered('save', async () => new Uint8Array(await requirePeer().edits.save()));

  return {
    handle: null, commands,
    get hydrated() { return session.ready; },
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
      await session.requestHydration('recovery');
      assertCurrent();
      await bridge().recoverInput();
      assertCurrent();
      if (!session.editPeer) throw new XlsxPeerNotReadyError();
      const result = session.editPeer.recoverySave();
      return { bytes: new Uint8Array(result.bytes), recovery: true };
    },
    clearSelection: () => { if (session.current && !session.failure) bridge().clearSelection(); },
    focus: () => { if (session.current && !session.failure) bridge().focus(); },
    refreshProposals: () => { if (session.ready) bridge().refreshProposals(); },
    selectCells(sheet, selection) {
      if (!session.ready || !session.peer || !validSelection(session.peer, sheet, selection)) return false;
      return synchronous((markApplied) => {
        requirePeer().edits.setActiveSheet(sheet);
        markApplied();
        return bridge().selectCells(sheet, selection);
      });
    },
    selectCellsAsync(sheet, selection) {
      const target = structuredClone(selection);
      return ordered('select-cells', (markApplied) => {
        const { peer, edits } = requirePeer();
        if (!validSelection(peer, sheet, target)) return false;
        edits.setActiveSheet(sheet);
        markApplied();
        return bridge().selectCellsAsync(sheet, target);
      }).then((selected) => selected ?? false);
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
      const input = structuredClone(request);
      return ordered('apply-edits', (markApplied) => {
        if (bridge().readOnly()) return refusal();
        const result = requirePeer().edits.applyEdits(input);
        markApplied();
        bridge().apply(result);
        return result;
      });
    },
    cell: (sheet, row, col) => session.ready ? session.peer!.cell(sheet, row, col) : null,
    cellAsync: (sheet, row, col) => ordered('cell', () => requirePeer().peer.cell(sheet, row, col)),
    rangeCells: (sheet, range) => session.ready ? session.peer!.rangeCells(sheet, range) : [],
    rangeCellsAsync: (sheet, range) => ordered('range-cells', () => requirePeer().peer.rangeCells(sheet, range)),
    readCellsSync: (request) => session.ready ? session.peer!.readCells(request) : null,
    editCell(sheet, row, col, input) {
      requirePeer();
      if (bridge().readOnly()) throw new Error('The editor is read-only');
      return synchronous((markApplied) => editCell(sheet, row, col, input, markApplied));
    },
    editCellAsync: (sheet, row, col, input) => ordered('edit-cell', (markApplied) => {
      if (bridge().readOnly()) return { error: new Error('The editor is read-only') };
      return { result: editCell(sheet, row, col, input, markApplied) };
    }).then((outcome) => {
      if (outcome?.error) throw outcome.error;
      return outcome?.result ?? null;
    }),
  };
}
