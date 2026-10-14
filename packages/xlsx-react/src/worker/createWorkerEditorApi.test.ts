import { describe, expect, mock, test } from 'bun:test';
import { selectionAt } from '@betteroffice/xlsx';
import type {
  CellEdit, EditResult, WorkbookHandle, XlsxEditRequest, XlsxEditResult,
  XlsxFindRequest, XlsxReadRequest, XlsxReadResult, XlsxValidationResult,
} from '@betteroffice/xlsx';
import { WorkbookEditPeerFailedError, type WorkbookEditPeer } from '@betteroffice/xlsx';
import { createXlsxCommandController } from '../commands/createXlsxCommandStore';
import type { InputDraft } from '../commands/inputCoordinator';
import { createWorkerInputCoordinator } from '../commands/workerInputCoordinator';
import {
  createWorkerEditorApi, XlsxPeerNotReadyError,
  type WorkerEditorApiBridge, type WorkerEditorSessionAccess,
} from './createWorkerEditorApi';

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function harness(ready = true) {
  const log: string[] = [];
  const hydration = deferred<void>();
  const state = {
    current: true, ready, failure: null as Error | null, generation: 1,
    readOnly: false, preview: Promise.resolve(), value: 'initial',
  };
  if (ready) hydration.resolve();
  const result: EditResult = {
    applied: true,
    sheetInfo: {
      sheetIds: ['sheet:0'], sheetNames: ['Sheet'], activeSheet: 0,
      contentWidth: 100, contentHeight: 100, frozenRows: 0, frozenCols: 0,
      initialScrollX: 0, initialScrollY: 0,
    },
  };
  const read: XlsxReadResult = {
    ok: true, version: 'v1', sheets: [], ranges: [],
    calculation: { cycleCells: [], limitedCells: [], truncated: false },
  };
  const batch: XlsxEditResult = {
    ok: true, applied: true, baseVersion: 'v1', version: 'v2', receipts: [], source: 'host',
    changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
  };
  const peerMethods = {
    cell: mock((_sheet: number, _row: number, _col: number): CellEdit => ({ a1: 'A1', input: state.value, isFormula: false })),
    rangeCells: mock((_sheet: number, _range: string): CellEdit[][] => [[{ a1: 'A1', input: state.value, isFormula: false }]]),
    sheetCount: () => 1,
    version: mock(() => 'v1'),
    readCells: mock((_request: XlsxReadRequest) => read),
    findText: mock((_request: XlsxFindRequest) => ({ ok: true as const, version: 'v1', matches: [], truncated: false })),
    validateEdits: mock((_request: XlsxEditRequest): XlsxValidationResult => ({
      ok: true, baseVersion: 'v1', wouldApply: true, previews: [],
    })),
    applyEdits: mock((_request: XlsxEditRequest) => batch),
    editCell: mock(() => result),
    setActiveSheet: mock(() => {}),
    save: mock(() => new Uint8Array([0])),
  };
  const flush = mock(async () => { log.push('flush'); });
  const editMethods = {
    state: 'ready' as const,
    editCell: mock((_sheet: number, _row: number, _col: number, input: string) => {
      state.value = input;
      log.push(`edit:${input}`);
      return result;
    }),
    applyEdits: mock((_request: XlsxEditRequest) => { log.push('batch'); return batch; }),
    setActiveSheet: mock((_sheet: number) => { log.push('sheet'); }),
    flush,
    save: mock(async () => {
      await flush();
      log.push(`save:${state.value}`);
      return new Uint8Array([8, 9]).buffer;
    }),
    recoverySave: mock(() => { log.push('recovery-save'); return { bytes: new Uint8Array([7]).buffer, recovery: true as const }; }),
  };
  const peer = peerMethods as unknown as WorkbookHandle;
  const edits = editMethods as unknown as WorkbookEditPeer;
  const session: WorkerEditorSessionAccess = {
    get current() { return state.current; }, get ready() { return state.ready && !state.failure && state.current; },
    get peer() { return state.ready ? peer : null; }, get editPeer() { return state.ready ? edits : null; },
    get failure() { return state.failure; },
    whenHydrated: mock(() => hydration.promise),
    requestHydration: mock((_reason: string) => hydration.promise),
  };
  const coordinator = createWorkerInputCoordinator({
    generation: () => state.generation, capture: () => ({ sheet: 0, target: 'A1' }),
    isReady: () => state.ready && !state.failure,
    whenReady: () => hydration.promise, seal: () => ({}), sync: () => {},
    preview: async () => { log.push('preview'); await state.preview; },
    write: (draft) => { edits.editCell(draft.sheet, draft.row, draft.col, draft.value); return true; },
    requestHydration: (reason) => session.requestHydration(reason), flushEdits: () => edits.flush(),
  });
  const bridge: WorkerEditorApiBridge = {
    coordinator: () => coordinator, readOnly: () => state.readOnly,
    clearSelection: mock(() => {}), focus: mock(() => {}), refreshProposals: mock(() => {}),
    recoverInput: mock(async () => { log.push('recover-input'); }),
    selectCells: mock(() => true), selectCellsAsync: mock(async () => true),
    apply: mock((_result: EditResult | XlsxEditResult) => { log.push('apply'); }),
  };
  const api = createWorkerEditorApi(session, createXlsxCommandController().store, () => bridge);
  const hydrate = () => { state.ready = true; hydration.resolve(); };
  const draft = (value: string): InputDraft => ({ generation: state.generation, sheet: 0, row: 0, col: 0, source: 'cell', value });
  return { api, bridge, coordinator, session, state, hydrate, draft, log, peer: peerMethods, edits: editMethods, result, batch };
}

describe('worker editor API', () => {
  test('keeps the raw handle absent and synchronous reads nullable before hydration', () => {
    const { api, peer, session } = harness(false);
    expect(api.handle).toBeNull();
    expect(api.hydrated).toBe(false);
    expect(api.cell(0, 0, 0)).toBeNull();
    expect(api.rangeCells(0, 'A1')).toEqual([]);
    expect(api.readCellsSync({ ranges: [] })).toBeNull();
    expect(api.selectCells(0, selectionAt({ row: 0, col: 0 }))).toBe(false);
    expect(() => api.editCell(0, 0, 0, 'text')).toThrow(XlsxPeerNotReadyError);
    expect(peer.cell).not.toHaveBeenCalled();
    expect(session.requestHydration).not.toHaveBeenCalled();
    for (const method of ['exportStructured', 'exportMarkdown', 'exportPng', 'renderPng', 'renderRangePng', 'printDisplayList']) {
      expect(method in api).toBe(false);
    }
  });

  test('reads cells, ranges, versions, search and validation from the peer', async () => {
    const { api, peer, edits } = harness();
    expect(api.cell(0, 0, 0)?.input).toBe('initial');
    expect((await api.cellAsync(0, 0, 0))?.input).toBe('initial');
    expect(api.rangeCells(0, 'A1')).toHaveLength(1);
    expect(await api.rangeCellsAsync(0, 'A1')).toHaveLength(1);
    expect(await api.version()).toBe('v1');
    expect(api.readCellsSync({ ranges: [] })?.version).toBe('v1');
    await api.readCells({ ranges: [] });
    await api.findText({ text: 'initial' });
    await api.validateEdits({ expectVersion: 'v1', steps: [] });
    expect(peer.readCells).toHaveBeenCalledTimes(2);
    expect(peer.findText).toHaveBeenCalledWith({ text: 'initial' });
    expect(peer.validateEdits).toHaveBeenCalledWith({ expectVersion: 'v1', steps: [] });
    expect(edits.applyEdits).not.toHaveBeenCalled();
    expect(edits.save).not.toHaveBeenCalled();
  });

  test('routes synchronous and async cell edits, batches and sheet changes through the facade', async () => {
    const { api, edits, peer, bridge, result, batch } = harness();
    expect(api.editCell(0, 0, 0, 'sync')).toBe(result);
    expect(await api.editCellAsync(0, 1, 0, 'async')).toBe(result);
    expect(await api.applyEdits({ expectVersion: 'v1', steps: [] })).toBe(batch);
    expect(api.selectCells(0, selectionAt({ row: 0, col: 0 }))).toBe(true);
    expect(await api.selectCellsAsync(0, selectionAt({ row: 1, col: 0 }))).toBe(true);
    expect(edits.editCell).toHaveBeenCalledTimes(2);
    expect(edits.applyEdits).toHaveBeenCalledTimes(1);
    expect(edits.setActiveSheet).toHaveBeenCalledTimes(2);
    expect(bridge.apply).toHaveBeenCalledTimes(3);
    expect(peer.editCell).not.toHaveBeenCalled();
    expect(peer.applyEdits).not.toHaveBeenCalled();
    expect(peer.setActiveSheet).not.toHaveBeenCalled();
    expect(peer.save).not.toHaveBeenCalled();
    expect(edits.save).not.toHaveBeenCalled();
  });

  test('demands async readiness without overtaking earlier accepted input', async () => {
    const { api, coordinator, draft, hydrate, log, session } = harness(false);
    const first = coordinator.submitAsync(draft('first'));
    const second = api.editCellAsync(0, 1, 0, 'second');
    const read = api.cellAsync(0, 1, 0);
    expect(session.requestHydration).toHaveBeenCalledWith('edit-cell');
    expect(log).toEqual([]);
    hydrate();
    await Promise.all([first, second]);
    expect((await read)?.input).toBe('second');
    expect(log).toEqual(['preview', 'edit:first', 'edit:second', 'apply']);
  });

  test('rejects synchronous mutations while accepted input still waits for preview paint', async () => {
    const { api, coordinator, draft, state, edits } = harness();
    const paint = deferred<void>();
    state.preview = paint.promise;
    const pending = coordinator.submitAsync(draft('pending'));
    expect(() => api.editCell(0, 1, 0, 'later')).toThrow(XlsxPeerNotReadyError);
    expect(edits.editCell).not.toHaveBeenCalled();
    paint.resolve();
    await pending;
    expect(api.editCell(0, 1, 0, 'later').applied).toBe(true);
  });

  test('withholds synchronous reads until accepted input crosses preview paint', async () => {
    const { api, coordinator, draft, state, peer } = harness();
    const paint = deferred<void>();
    state.preview = paint.promise;
    const pending = coordinator.submitAsync(draft('pending'));
    expect(api.cell(0, 0, 0)).toBeNull();
    expect(api.rangeCells(0, 'A1')).toEqual([]);
    expect(api.readCellsSync({ ranges: [] })).toBeNull();
    expect(peer.cell).not.toHaveBeenCalled();
    expect(peer.rangeCells).not.toHaveBeenCalled();
    expect(peer.readCells).not.toHaveBeenCalled();
    paint.resolve();
    await pending;
    expect(api.cell(0, 0, 0)?.input).toBe('pending');
    expect(api.rangeCells(0, 'A1')[0][0].input).toBe('pending');
    expect(api.readCellsSync({ ranges: [] })?.version).toBe('v1');
    coordinator.setDraft(draft('live'));
    expect(api.cell(0, 0, 0)).toBeNull();
    expect(api.rangeCells(0, 'A1')).toEqual([]);
    expect(api.readCellsSync({ ranges: [] })).toBeNull();
    await api.flush();
    expect(api.cell(0, 0, 0)?.input).toBe('live');
  });

  test('saves through the facade after pending input and returns Uint8Array bytes', async () => {
    const { api, coordinator, draft, hydrate, log, edits, peer } = harness(false);
    coordinator.setDraft(draft('saved'));
    const saving = api.save();
    hydrate();
    expect(await saving).toEqual(new Uint8Array([8, 9]));
    expect(log).toEqual(['preview', 'edit:saved', 'flush', 'save:saved']);
    expect(await api.saveAsync()).toEqual(new Uint8Array([8, 9]));
    expect(edits.save).toHaveBeenCalledTimes(2);
    expect(peer.save).not.toHaveBeenCalled();
  });

  test('flush demands hydration and drains accepted input before edit acknowledgements', async () => {
    const { api, coordinator, draft, hydrate, session, log } = harness(false);
    coordinator.setDraft(draft('flushed'));
    const flushing = api.flush();
    expect(session.requestHydration).toHaveBeenCalledWith('flush');
    hydrate();
    await flushing;
    expect(log.slice(0, 2)).toEqual(['preview', 'edit:flushed']);
    expect(log.slice(2).every((entry) => entry === 'flush')).toBe(true);
    expect(coordinator.pending).toBe(false);
  });

  test('keeps validation and batch read-only refusals while blocking cell mutations', async () => {
    const { api, state, peer, edits } = harness();
    state.readOnly = true;
    const request = { expectVersion: 'v1', steps: [] };
    const refusal = { ok: false, version: 'v1', failure: { code: 'read-only', message: 'The editor is read-only' } } as const;
    expect(await api.validateEdits(request)).toEqual(refusal);
    expect(await api.applyEdits(request)).toEqual(refusal);
    expect(() => api.editCell(0, 0, 0, 'blocked')).toThrow('The editor is read-only');
    await expect(api.editCellAsync(0, 0, 0, 'blocked')).rejects.toThrow('The editor is read-only');
    expect(await api.version()).toBe('v1');
    expect(peer.validateEdits).not.toHaveBeenCalled();
    expect(edits.applyEdits).not.toHaveBeenCalled();
    expect(edits.editCell).not.toHaveBeenCalled();
  });

  test('snapshots queued requests and drains accepted save input before generation replacement', async () => {
    const { api, state, coordinator, hydrate, peer, edits, log } = harness(false);
    const request: XlsxReadRequest = { ranges: [] };
    const reading = api.readCells(request);
    request.ranges = [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2' } }];
    hydrate();
    await reading;
    expect(peer.readCells).toHaveBeenCalledWith({ ranges: [] });
    state.ready = false;
    const gate = deferred<void>();
    state.preview = gate.promise;
    coordinator.setDraft({ generation: 1, sheet: 0, row: 0, col: 0, source: 'cell', value: 'old' });
    const saving = api.saveAsync();
    const draining = coordinator.drain();
    expect(edits.editCell).not.toHaveBeenCalled();
    hydrate();
    gate.resolve();
    expect(await saving).toEqual(new Uint8Array([8, 9]));
    await draining;
    expect(edits.editCell.mock.calls).toEqual([[0, 0, 0, 'old']]);
    expect(log.indexOf('edit:old')).toBeLessThan(log.indexOf('save:old'));
    state.current = false;
    state.generation += 1;
    coordinator.reset();
    expect(edits.editCell).toHaveBeenCalledTimes(1);
    expect(await api.readCells({ ranges: [] })).toBeNull();
    expect(await api.save()).toBeNull();
  });

  test('preserves typed facade failures and keeps recovery saving on the peer side', async () => {
    const { api, state, coordinator, edits, peer, bridge, log } = harness();
    await expect(api.recoverySave()).rejects.toThrow('Recovery requires a failed workbook edit peer');
    expect(bridge.recoverInput).not.toHaveBeenCalled();
    const failure = new WorkbookEditPeerFailedError(new Error('Replay failed'));
    state.failure = failure;
    coordinator.fail(failure);
    expect(api.failure).toBe(failure);
    await expect(api.saveAsync()).rejects.toBe(failure);
    expect(await api.recoverySave()).toEqual({ bytes: new Uint8Array([7]), recovery: true });
    expect(bridge.recoverInput).toHaveBeenCalledTimes(1);
    expect(log).toEqual(['recover-input', 'recovery-save']);
    expect(edits.recoverySave).toHaveBeenCalledTimes(1);
    expect(edits.save).not.toHaveBeenCalled();
    expect(peer.save).not.toHaveBeenCalled();
  });
});
