import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, mock, test } from 'bun:test';
import { selectionAt, StaleProposalError } from '@betteroffice/xlsx';
import type { EditResult, WorkbookHandle } from '@betteroffice/xlsx';
import { WorkbookEditPeerFailedError, type WorkbookEditPeer } from '@betteroffice/xlsx';
import { createXlsxCommandController } from './createXlsxCommandStore';
import type { XlsxCommandEnvironment } from './evaluate';
import type { InputDraft } from './inputCoordinator';
import { PLAIN_FORMATTING } from './testing';
import type { XlsxEditorView } from './useXlsxCommands';
import {
  createWorkerXlsxCommandBinding, useWorkerXlsxCommands, type WorkerXlsxEditorBridge,
} from './useWorkerXlsxCommands';
import { createWorkerInputCoordinator } from './workerInputCoordinator';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const restorers: (() => void)[] = [];

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => { resolve = done; reject = failed; });
  return { promise, resolve, reject };
}

function harness(ready = true, attach = true) {
  const log: string[] = [];
  const errors: unknown[] = [];
  const hydration = deferred<void>();
  const previewEntered = deferred<void>();
  const paintEntered = deferred<void>();
  const state = {
    ready, generation: 1, mutation: 0, status: 'ready' as XlsxCommandEnvironment['status'],
    readOnly: false, collaborative: false, canUndo: true, canRedo: true,
    preview: Promise.resolve(), paint: Promise.resolve(true),
  };
  if (ready) hydration.resolve();
  const view: XlsxEditorView = {
    sheet: 0, selection: { anchor: { row: 0, col: 0 }, focus: { row: 2, col: 1 } },
    chartSelected: false, zoom: 1, capturedFormat: null, borderStyle: undefined, borderColor: undefined,
    proposals: [{ id: 'p1', agentId: 'author', note: null, cells: [] }],
    proposalsAvailable: true, proposalsPanelOpen: false, pngExport: true,
  };
  const result: EditResult = {
    applied: true,
    sheetInfo: {
      sheetIds: ['sheet:0'], sheetNames: ['Sheet'], activeSheet: 0,
      contentWidth: 100, contentHeight: 100, frozenRows: 0, frozenCols: 0,
      initialScrollX: 0, initialScrollY: 0,
    },
  };
  const rawMutation = mock((..._args: unknown[]) => { throw new Error('Unexpected peer mutation'); });
  const peerMethods = {
    cell: mock((_sheet: number, row: number, col: number) => ({ a1: `${String.fromCharCode(65 + col)}${row + 1}`, input: '', isFormula: false })),
    selectionFormatting: mock((_sheet: number, _range: string) => PLAIN_FORMATTING),
    mergedRanges: mock((_sheet: number, _range: string) => [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }]),
    historyState: mock(() => ({ canUndo: state.canUndo, canRedo: state.canRedo, undoDepth: 1, redoDepth: 1 })),
    captureFormat: mock((_sheet: number, _range: string) => ({ rows: 1, columns: 1, formats: [] })),
    patchRangeStyle: rawMutation, setNumberFormat: rawMutation, applyOps: rawMutation,
    undo: rawMutation, redo: rawMutation, acceptProposal: rawMutation, rejectProposal: rawMutation,
    editCell: rawMutation, save: rawMutation,
  };
  const flush = mock(async () => { log.push('flush'); });
  const edits = {
    state: 'ready' as const,
    editCell: mock((_sheet: number, _row: number, _col: number, value: string) => { log.push(`edit:${value}`); return result; }),
    patchRangeStyle: mock((..._args: unknown[]) => { log.push('style'); return result; }),
    setNumberFormat: mock((..._args: unknown[]) => result),
    applyOps: mock((..._args: unknown[]) => result),
    undo: mock(() => result), redo: mock(() => result),
    acceptProposal: mock((..._args: unknown[]) => ({ ...result, proposalId: 'p1' })),
    rejectProposal: mock((_id: string) => true),
    flush,
    save: mock(async () => { await flush(); log.push('save'); return new Uint8Array([8, 9]).buffer; }),
  };
  const facade = edits as unknown as WorkbookEditPeer;
  const coordinator = createWorkerInputCoordinator({
    generation: () => state.generation, capture: () => ({ sheet: view.sheet, target: 'selection' }),
    isReady: () => state.ready, whenReady: () => hydration.promise,
    seal: () => ({}), sync: () => {},
    preview: async () => { log.push('input-preview'); },
    write: (draft) => { facade.editCell(draft.sheet, draft.row, draft.col, draft.value); return true; },
    requestHydration: () => hydration.promise, flushEdits: () => facade.flush(),
    onError: (error) => errors.push(error),
  });
  const bridge: WorkerXlsxEditorBridge = {
    peer: () => state.ready ? peerMethods as unknown as WorkbookHandle : null,
    editPeer: () => state.ready ? facade : null,
    status: () => state.status, readOnly: () => state.readOnly, collaborative: () => state.collaborative,
    mutation: () => state.mutation, generation: () => state.generation, view: () => view,
    coordinator, i18n: () => undefined, translate: (key) => key,
    apply: mock((_result: EditResult) => { state.mutation += 1; }), fail: (error) => errors.push(error),
    preview: mock(async () => { log.push('command-preview'); previewEntered.resolve(); await state.preview; }),
    setZoom: (scale) => { view.zoom = scale; },
    setProposalsPanelOpen: (open) => { view.proposalsPanelOpen = open; },
    setCapturedFormat: (format) => { view.capturedFormat = format; },
    setBorderStyle: (style) => { view.borderStyle = style; },
    setBorderColor: (color) => { view.borderColor = color; },
    markStale: mock((_id: string, _cells: string[] | null) => {}), refreshProposals: mock(() => {}),
    deliver: mock((_bytes: Uint8Array) => { log.push('deliver'); }),
    afterPaint: mock(() => { paintEntered.resolve(); return state.paint; }), focusGrid: mock(() => {}),
  };
  const controller = createXlsxCommandController();
  const binding = createWorkerXlsxCommandBinding(() => bridge);
  if (attach) controller.attach(binding);
  const hydrate = () => { state.ready = true; hydration.resolve(); };
  const draft = (value: string): InputDraft => ({ generation: state.generation, sheet: view.sheet, row: 0, col: 0, source: 'cell', value });
  return { controller, store: controller.store, binding, bridge, state, view, coordinator, hydrate, draft, peer: peerMethods, edits, rawMutation, log, errors, previewEntered, paintEntered };
}

afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('worker XLSX commands', () => {
  test('reads formatting, merged ranges and history from the peer and writes styles through the facade', async () => {
    const { store, peer, edits, rawMutation } = harness();
    expect(store.getState('bold').active).toBe(false);
    expect(store.getState('undo').enabled).toBe(true);
    expect(await store.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
    expect(peer.selectionFormatting).toHaveBeenCalledWith(0, 'A1:B3');
    expect(peer.mergedRanges).toHaveBeenCalledWith(0, 'A1:B3');
    expect(peer.historyState).toHaveBeenCalled();
    expect(edits.patchRangeStyle).toHaveBeenCalledWith(0, 'A1:B3', { bold: true });
    expect(rawMutation).not.toHaveBeenCalled();
    expect(edits.save).not.toHaveBeenCalled();
  });

  test('routes number formats, merges, history and proposals through the facade', async () => {
    const { store, edits, rawMutation, bridge } = harness();
    await store.execute('numberFormat', { value: 'custom' });
    await store.execute('decimalPlaces', { direction: 'increase' });
    await store.execute('merge', { value: 'horizontal' });
    await store.execute('undo', null);
    await store.execute('redo', null);
    await store.execute('proposalAccept', { proposalId: 'p1', force: true });
    await store.execute('proposalReject', { proposalId: 'p1' });
    expect(edits.setNumberFormat).toHaveBeenCalledWith(0, 'A1:B3', { type: 'custom', pattern: '0.00' });
    expect(edits.setNumberFormat).toHaveBeenCalledWith(0, 'A1:B3', 'increaseDecimal');
    expect(edits.applyOps).toHaveBeenCalledWith([
      { type: 'mergeCells', sheet: 0, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } } },
      { type: 'mergeCells', sheet: 0, range: { start: { row: 1, col: 0 }, end: { row: 1, col: 1 } } },
      { type: 'mergeCells', sheet: 0, range: { start: { row: 2, col: 0 }, end: { row: 2, col: 1 } } },
    ]);
    expect(edits.undo).toHaveBeenCalledTimes(1);
    expect(edits.redo).toHaveBeenCalledTimes(1);
    expect(edits.acceptProposal).toHaveBeenCalledWith('p1', { force: true });
    expect(edits.rejectProposal).toHaveBeenCalledWith('p1');
    expect(bridge.refreshProposals).toHaveBeenCalledTimes(2);
    expect(rawMutation).not.toHaveBeenCalled();
    expect(edits.save).not.toHaveBeenCalled();
  });

  test('captures paint formatting from the peer without mutating it', async () => {
    const { store, view, peer, rawMutation } = harness();
    await store.execute('paintFormat', null);
    expect(peer.captureFormat).toHaveBeenCalledWith(0, 'A1:B3');
    expect(view.capturedFormat).toEqual({ rows: 1, columns: 1, formats: [] });
    await store.execute('paintFormat', null);
    expect(view.capturedFormat).toBeNull();
    expect(rawMutation).not.toHaveBeenCalled();
  });

  test('captures targets before a peer exists and preserves FIFO with accepted input', async () => {
    const { store, coordinator, binding, draft, hydrate, log, edits } = harness(false);
    expect(binding.capture('bold')).toEqual({ generation: 1, target: '0:{"top":0,"bottom":2,"left":0,"right":1}' });
    const input = coordinator.submitAsync(draft('typed'));
    const bold = store.execute('bold', null);
    const italic = store.execute('italic', null);
    expect(log).toEqual([]);
    hydrate();
    await Promise.all([input, bold, italic]);
    expect(log).toEqual(['input-preview', 'edit:typed', 'command-preview', 'style', 'command-preview', 'style']);
    expect(edits.patchRangeStyle.mock.calls).toEqual([
      [0, 'A1:B3', { bold: true }], [0, 'A1:B3', { italic: true }],
    ]);
  });

  test('refuses queued commands when generation or target changes', async () => {
    const first = harness(false);
    const replaced = first.store.execute('bold', null);
    first.state.generation += 1;
    first.hydrate();
    expect(await replaced).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    expect(first.edits.patchRangeStyle).not.toHaveBeenCalled();
    const second = harness(false);
    const moved = second.store.execute('bold', null);
    second.view.selection = selectionAt({ row: 8, col: 2 });
    second.hydrate();
    expect(await moved).toMatchObject({ ok: false, failure: { code: 'target-changed' } });
    expect(second.edits.patchRangeStyle).not.toHaveBeenCalled();
  });

  for (const readOnly of [false, true]) {
    test(`uses submission-time permissions for queued formatting (${readOnly ? 'read-only' : 'editable'})`, async () => {
      const { store, state, previewEntered, edits } = harness();
      const paint = deferred<void>();
      state.preview = paint.promise;
      state.readOnly = readOnly;
      const command = store.execute('bold', null);
      state.readOnly = !readOnly;
      await previewEntered.promise;
      expect(edits.patchRangeStyle).not.toHaveBeenCalled();
      paint.resolve();
      expect(await command).toMatchObject(readOnly ? { ok: false, failure: { code: 'read-only' } } : { ok: true });
      expect(edits.patchRangeStyle).toHaveBeenCalledTimes(readOnly ? 0 : 1);
    });
  }

  test('rechecks scoped grants after readiness and preview paint', async () => {
    const { controller, state, hydrate, previewEntered, edits } = harness(false);
    const paint = deferred<void>();
    state.preview = paint.promise;
    let denied = false;
    const store = controller.scoped({ deny: () => denied ? 'permission-denied' : null, subscribe: () => () => {} });
    const command = store.execute('bold', null);
    hydrate();
    await previewEntered.promise;
    denied = true;
    paint.resolve();
    expect(await command).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });
    expect(edits.patchRangeStyle).not.toHaveBeenCalled();
  });

  test('keeps selection, collaboration, argument and history gates unchanged', async () => {
    const { store, view, state, edits, controller } = harness();
    view.chartSelected = true;
    expect(await store.execute('bold', null)).toMatchObject({ ok: false, failure: { code: 'unsupported-selection' } });
    view.chartSelected = false;
    view.selection = selectionAt({ row: 0, col: 0 });
    expect(await store.execute('merge', { value: 'all' })).toMatchObject({ ok: false, failure: { code: 'multiple-cells-required' } });
    view.selection = { anchor: { row: 0, col: 0 }, focus: { row: 2, col: 1 } };
    state.collaborative = true;
    expect(await store.execute('merge', { value: 'all' })).toMatchObject({ ok: false, failure: { code: 'collaboration-unsupported' } });
    state.collaborative = false;
    expect(await store.execute('fontSize', { points: 0 })).toMatchObject({ ok: false, failure: { code: 'invalid-arguments' } });
    state.canUndo = false;
    state.canRedo = false;
    state.mutation += 1;
    controller.refresh();
    expect(await store.execute('undo', null)).toMatchObject({ ok: false, failure: { code: 'nothing-to-undo' } });
    expect(await store.execute('redo', null)).toMatchObject({ ok: false, failure: { code: 'nothing-to-redo' } });
    expect(edits.patchRangeStyle).not.toHaveBeenCalled();
    expect(edits.undo).not.toHaveBeenCalled();
    expect(edits.redo).not.toHaveBeenCalled();
  });

  test('saves through the facade after drafts and delivers worker bytes', async () => {
    const { store, coordinator, draft, hydrate, bridge, edits, rawMutation, log } = harness(false);
    coordinator.setDraft(draft('saved'));
    const saving = store.execute('save', null);
    hydrate();
    expect(await saving).toEqual({ ok: true, status: 'executed' });
    expect(log).toEqual(['input-preview', 'edit:saved', 'command-preview', 'flush', 'save', 'deliver']);
    expect(bridge.deliver).toHaveBeenCalledWith(new Uint8Array([8, 9]));
    expect(edits.save).toHaveBeenCalledTimes(1);
    expect(rawMutation).not.toHaveBeenCalled();
  });

  test('keeps live read-only gates during an admitted save and preserves editable queued commands', async () => {
    const { store, controller, binding, state, edits, bridge } = harness();
    const saveEntered = deferred<void>();
    const saved = deferred<ArrayBuffer>();
    edits.save.mockImplementation(async () => { saveEntered.resolve(); return saved.promise; });
    const plugin = 'plugin:acme/stamp' as const;
    const execute = mock(async () => {
      edits.applyOps([{ type: 'stamp' }]);
      return { ok: true as const, status: 'executed' as const };
    });
    controller.setPluginCommands([{
      descriptor: { id: plugin, label: 'Stamp', mutatesDocument: true, shortcuts: [] },
      state: () => ({ enabled: true }),
      execute,
    }], []);
    expect(store.getState('bold').enabled).toBe(true);
    expect(store.getState(plugin).enabled).toBe(true);
    const saving = store.execute('save', null);
    await saveEntered.promise;
    const editable = store.execute('bold', null);
    state.readOnly = true;
    controller.refresh();
    expect(binding.environment(false).readOnly).toBe(true);
    expect(binding.environment(true).readOnly).toBe(true);
    expect(store.getState('bold')).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
    expect(store.getState(plugin)).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
    expect(await store.execute(plugin, null)).toMatchObject({ ok: false, failure: { code: 'read-only' } });
    expect(execute).not.toHaveBeenCalled();
    expect(edits.applyOps).not.toHaveBeenCalled();
    expect(edits.patchRangeStyle).not.toHaveBeenCalled();
    expect(bridge.deliver).not.toHaveBeenCalled();
    saved.resolve(new Uint8Array([8, 9]).buffer);
    expect(await saving).toEqual({ ok: true, status: 'executed' });
    expect(await editable).toEqual({ ok: true, status: 'executed' });
    expect(bridge.deliver).toHaveBeenCalledWith(new Uint8Array([8, 9]));
    expect(edits.patchRangeStyle.mock.calls).toEqual([[0, 'A1:B3', { bold: true }]]);
    expect(execute).not.toHaveBeenCalled();
    expect(edits.applyOps).not.toHaveBeenCalled();
  });

  test('disables PNG exports with the viewer png-unavailable reason', async () => {
    const { store, binding } = harness();
    expect(binding.environment(false).pngExport).toBe(false);
    expect(store.getState('exportPng')).toMatchObject({ enabled: false, disabledReason: { code: 'png-unavailable' } });
    expect(await store.execute('exportPng', null)).toMatchObject({ ok: false, failure: { code: 'png-unavailable' } });
  });

  test('prints only after the matching accepted paint and refuses replacement', async () => {
    const original = Object.getOwnPropertyDescriptor(window, 'print');
    const print = mock(() => {});
    Object.defineProperty(window, 'print', { configurable: true, value: print });
    restorers.push(() => {
      if (original) Object.defineProperty(window, 'print', original);
      else Reflect.deleteProperty(window, 'print');
    });
    const first = harness();
    const paint = deferred<boolean>();
    first.state.paint = paint.promise;
    const printing = first.store.execute('print', null);
    await first.paintEntered.promise;
    expect(print).not.toHaveBeenCalled();
    paint.resolve(true);
    expect(await printing).toEqual({ ok: true, status: 'executed' });
    expect(print).toHaveBeenCalledTimes(1);
    const second = harness();
    const nextPaint = deferred<boolean>();
    second.state.paint = nextPaint.promise;
    const replaced = second.store.execute('print', null);
    await second.paintEntered.promise;
    second.state.generation += 1;
    nextPaint.resolve(true);
    expect(await replaced).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    expect(print).toHaveBeenCalledTimes(1);
  });

  test('retains stale proposal reporting and typed facade errors', async () => {
    const { store, edits, bridge, errors } = harness();
    edits.acceptProposal.mockImplementation(() => { throw new StaleProposalError(['A1']); });
    expect(await store.execute('proposalAccept', { proposalId: 'p1' })).toMatchObject({ ok: false, failure: { code: 'proposal-stale' } });
    expect(bridge.markStale).toHaveBeenCalledWith('p1', ['A1']);
    const failure = new WorkbookEditPeerFailedError(new Error('Replay stopped'));
    edits.patchRangeStyle.mockImplementation(() => { throw failure; });
    expect(await store.execute('bold', null)).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
    expect(errors).toEqual([failure]);
  });

  test('retains unapplied commands when preview fails before facade mutation', async () => {
    const { store, state, coordinator, edits, errors, previewEntered } = harness();
    const paint = deferred<void>();
    state.preview = paint.promise;
    const command = store.execute('bold', null);
    await previewEntered.promise;
    const failure = new Error('Preview failed');
    paint.reject(failure);
    expect(await command).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
    expect(coordinator.unapplied).toHaveLength(1);
    expect(edits.patchRangeStyle).not.toHaveBeenCalled();
    expect(errors).toEqual([failure]);
  });

  test('attaches the hook to the existing controller and detaches on unmount', async () => {
    const first = harness(true, false);
    const second = harness(true, false);
    const { rerender, unmount } = renderHook(
      (bridge: WorkerXlsxEditorBridge) => useWorkerXlsxCommands(first.controller, bridge),
      { initialProps: first.bridge }
    );
    await act(async () => { await first.store.execute('bold', null); });
    expect(first.edits.patchRangeStyle).toHaveBeenCalledTimes(1);
    rerender(second.bridge);
    await act(async () => { await first.store.execute('bold', null); });
    expect(second.edits.patchRangeStyle).toHaveBeenCalledTimes(1);
    unmount();
    expect(first.store.getState('bold')).toMatchObject({ enabled: false, disabledReason: { code: 'editor-unavailable' } });
  });
});
