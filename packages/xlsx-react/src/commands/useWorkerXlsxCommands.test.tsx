import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { selectionAt, StaleProposalError } from '@betteroffice/xlsx';
import type { CommandStatus, EditResult, WorkbookAck, WorkbookCommandSelection, WorkbookEditState, WorkbookHandle } from '@betteroffice/xlsx';
import { WorkbookEditPeerFailedError, type WorkbookEditPeer } from '@betteroffice/xlsx';
import { createXlsxCommandController } from './createXlsxCommandStore';
import type { XlsxCommandEnvironment } from './evaluate';
import type { WorkbookFormatTokenResult } from '../../../xlsx/src/session/commandProtocol';
import type { InputDraft } from './inputCoordinator';
import { PLAIN_FORMATTING } from './testing';
import type { XlsxEditorView } from './useXlsxCommands';
import { createWorkerXlsxCommandBinding, useWorkerXlsxCommands, type WorkerXlsxEditorBridge } from './useWorkerXlsxCommands';
import { createPeerlessWorkerXlsxCommandBinding, usePeerlessWorkerXlsxCommands, type PeerlessWorkerXlsxEditorBridge } from './usePeerlessWorkerXlsxCommands';
import { createWorkerInputCoordinator } from './workerInputCoordinator';
import { createPeerlessWorkerInputCoordinator, type PeerlessWorkerCommand } from './peerlessWorkerInputCoordinator';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const restorers: (() => void)[] = [];

function peerlessCommands(attach = true) {
  const legacy = harness(true, false);
  const listeners = new Set<(ack: WorkbookAck) => void>();
  const sent: PeerlessWorkerCommand[] = [];
  const replies = new Map<number, ReturnType<typeof deferred<CommandStatus>>>();
  let throughSeq = 0;
  let version = 'v0';
  let workerBold = false;
  let editState: Partial<WorkbookEditState> = {
    formatting: { ...PLAIN_FORMATTING }, formattingAvailable: true, selectionEditable: true,
    mergedRanges: [], canUndo: true, canRedo: true, readOnly: false,
    activeSheet: 0, sheets: [{ index: 0, id: 'sheet:0', name: 'Sheet' }],
  };
  let selection: WorkbookCommandSelection | undefined;
  let token: string | null = null;
  let formatRevision = 0;
  const session = {
    commandIdentity: { session: 'book', incarnation: 'worker' },
    get confirmedBasis() { return { epoch: 1, version }; },
    get throughSeq() { return throughSeq; },
    get editState() { return editState; },
    get selection() { return selection; },
    submit(command: PeerlessWorkerCommand) {
      sent.push(command);
      const result = deferred<CommandStatus>();
      replies.set(command.seq, result);
      return result.promise;
    },
    onAck(listener: (ack: WorkbookAck) => void) {
      listeners.add(listener); return () => { listeners.delete(listener); };
    },
    save: async () => new Uint8Array([8, 9]), whenPresented: mock(async (_seq: number) => {}),
  };
  const coordinator = createPeerlessWorkerInputCoordinator({
    session: () => session, generation: () => legacy.state.generation,
    isReady: () => legacy.state.ready,
    seal: () => ({}), sync: () => {},
    target: (_sheet, row, col) => ({ sheetId: 'sheet:0', range: {
      kind: 'rowCol', start: { row, col }, end: { row, col },
    } }),
    onError: legacy.bridge.fail,
  });
  const bridge: PeerlessWorkerXlsxEditorBridge = {
    ...legacy.bridge, session: () => session, coordinator,
    frame: () => undefined, selectionRevision: () => 0,
    formatToken: () => token, formatRevision: () => formatRevision,
    setFormatToken: (value) => { token = value; formatRevision++; },
  };
  const binding = createPeerlessWorkerXlsxCommandBinding(() => bridge);
  if (attach) legacy.controller.attach(binding);
  const push = (ack: WorkbookAck) => {
    if (ack.throughSeq >= throughSeq) version = ack.version;
    throughSeq = Math.max(throughSeq, ack.throughSeq);
    editState = { ...editState, ...ack.editState };
    if (ack.selection) selection = ack.selection;
    for (const listener of listeners) listener(ack);
    for (const status of ack.statuses) replies.get(status.seq)?.resolve(status);
  };
  const ack = (seq: number, edits?: Partial<WorkbookEditState>, restored?: WorkbookCommandSelection, status: CommandStatus['status'] = 'applied') => {
    const command = sent.find((entry) => entry.seq === seq);
    if (status === 'applied' && command?.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'bold') {
      workerBold = !workerBold;
      edits = { ...edits, formatting: { ...PLAIN_FORMATTING, bold: workerBold } };
    }
    push({ throughSeq: seq, version: `v${seq}`, editState: edits, selection: restored,
      statuses: [{ seq, status, ...(status === 'refused' ? { reason: 'stale-proposal', result: { cells: ['A1'] } } : {}) }],
    });
  };
  return { ...legacy, bridge, coordinator, binding, session, sent, push, ack };
}

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

describe('peerless worker XLSX commands', () => {
  test('save consults the host request and honors its veto', async () => {
    const { store, bridge, session } = peerlessCommands();
    const save = spyOn(session, 'save');
    restorers.push(() => save.mockRestore());
    bridge.requestSave = mock(async () => ({ ok: true as const, status: 'noop' as const }));
    expect(await store.execute('save', null)).toEqual({ ok: true, status: 'noop' });
    expect(bridge.requestSave).toHaveBeenCalledTimes(1);
    expect(save).not.toHaveBeenCalled();
    expect(bridge.deliver).not.toHaveBeenCalled();
  });

  for (const readOnly of [false, true]) {
    test(`queued peerless formatting keeps submission-time permissions (${readOnly ? 'read-only' : 'editable'})`, async () => {
      const { store, state, coordinator, sent, ack } = peerlessCommands();
      const capture = deferred<string>();
      const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
        kind: 'editCell', payload: { args: [0, 0, 0, value] },
      }));
      state.readOnly = readOnly;
      const formatting = store.execute('bold', null);
      state.readOnly = !readOnly;
      const posted = deferred<void>();
      const off = coordinator.subscribe((event) => {
        if (event.type === 'dispatched' && event.seq === (readOnly ? 1 : 2)) posted.resolve();
      });
      capture.resolve('paste');
      await posted.promise;
      off();
      ack(1);
      if (!readOnly) ack(2);
      expect(await formatting).toMatchObject(readOnly
        ? { ok: false, failure: { code: 'read-only' } } : { ok: true, status: 'executed' });
      await pasted.result;
      expect(sent.map((command) => command.kind)).toEqual(readOnly ? ['editCell'] : ['editCell', 'applyOps']);
    });
  }

  test('peerless scoped grants are rechecked after clipboard preparation', async () => {
    const { controller, coordinator, sent, ack, errors } = peerlessCommands();
    const capture = deferred<string>();
    let denied = false;
    const scoped = controller.scoped({ deny: () => denied ? 'permission-denied' : null, subscribe: () => () => {} });
    const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 0, value] },
    }));
    const undo = scoped.execute('undo', null);
    denied = true;
    capture.resolve('paste');
    expect(await undo).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });
    expect(sent.map((command) => command.kind)).toEqual(['editCell']);
    ack(1);
    await pasted.result;
    expect(errors).toEqual([]);
  });

  test('non-document capture stays executed and arms the exact token at an unchanged opaque version', async () => {
    const { store, bridge, push, session, errors } = peerlessCommands();
    push({ throughSeq: 0, version: 'engine:opaque/α', statuses: [] });
    const capturing = store.execute('paintFormat', null);
    push({ throughSeq: 1, version: 'engine:opaque/α', statuses: [{
      seq: 1, status: 'applied', committed: false, result: { formatToken: 'format:opaque/β' },
    }] });
    expect(await capturing).toEqual({ ok: true, status: 'executed' });
    expect(bridge.formatToken()).toBe('format:opaque/β');
    expect(store.getState('paintFormat').active).toBe(true);
    expect(session.throughSeq).toBe(1);
    expect(session.confirmedBasis.version).toBe('engine:opaque/α');
    expect(errors).toEqual([]);
  });

  test('successful non-document proposal rejection differs from a false noop result', async () => {
    const { store, push, bridge, errors } = peerlessCommands();
    const accepted = store.execute('proposalReject', { proposalId: 'p1' });
    push({ throughSeq: 1, version: 'engine:opaque/α', statuses: [{
      seq: 1, status: 'applied', committed: false, result: true,
    }] });
    expect(await accepted).toEqual({ ok: true, status: 'executed' });
    expect(bridge.markStale).toHaveBeenCalledWith('p1', null);
    expect(bridge.refreshProposals).toHaveBeenCalledTimes(1);
    const missing = store.execute('proposalReject', { proposalId: 'p1' });
    push({ throughSeq: 2, version: 'engine:opaque/α', statuses: [{
      seq: 2, status: 'noop', committed: false, result: false,
    }] });
    expect(await missing).toMatchObject({ ok: false, failure: { code: 'proposal-not-found' } });
    expect(bridge.refreshProposals).toHaveBeenCalledTimes(2);
    expect(errors).toEqual([]);
  });

  for (const id of ['bold', 'fillColor', 'numberFormat'] as const) {
    test(`${id} preserves the admitted reverse selection focus`, async () => {
      const { store, bridge, sent, push } = peerlessCommands();
      const selection: WorkbookCommandSelection = { sheetId: 'sheet:0', activeCell: { row: 0, col: 0 },
        range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 0 } } };
      bridge.selection = () => selection;
      const formatting = id === 'bold' ? store.execute(id, null) : id === 'fillColor'
        ? store.execute(id, { color: '#123456' }) : store.execute(id, { value: 'number' });
      expect(sent[0].payload).toMatchObject({ selection });
      expect('target' in sent[0].payload ? sent[0].payload.target :
        'toolbar' in sent[0].payload ? sent[0].payload.toolbar.target : undefined).toEqual({
        sheetId: selection.sheetId, range: selection.range,
      });
      push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'applied' }], selection });
      expect(await formatting).toMatchObject({ ok: true });
    });
  }

  test('delayed formatting retains its admitted frame and selection revision', async () => {
    const host = peerlessCommands();
    const capture = deferred<string>();
    const pasted = host.coordinator.clipboard(() => capture.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 0, value] },
    }));
    const frame = { sheet: 0, viewport: { x: 0, y: 0, width: 800, height: 600 }, zoom: 1, dpr: 1,
      generation: 1, requestGeneration: 2 };
    let revision = 3;
    host.bridge.frame = () => frame;
    host.bridge.selectionRevision = () => revision;
    const events: import('./peerlessWorkerInputCoordinator').PeerlessWorkerInputEvent[] = [];
    host.coordinator.subscribe((event) => events.push(event));
    const formatting = host.store.execute('bold', null);
    revision = 4;
    frame.requestGeneration = 5;
    capture.resolve('paste');
    await act(async () => {});
    expect(host.sent[1].frame).toMatchObject({ requestGeneration: 2 });
    expect(events.find((event) => event.type === 'dispatched' && event.seq === 2))
      .toMatchObject({ selectionRevision: 3 });
    host.ack(1); host.ack(2);
    await Promise.all([pasted.result, formatting]);
  });

  test('refuses a malformed format capture token at runtime', async () => {
    const required: WorkbookFormatTokenResult extends { formatToken: string } ? true : false = true;
    expect(required).toBe(true);
    const { store, bridge, push, errors } = peerlessCommands();
    const capturing = store.execute('paintFormat', null);
    push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'applied', result: { formatToken: 7 } }] });
    expect(await capturing).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
    expect(bridge.formatToken()).toBeNull();
    expect(errors).toHaveLength(1);
  });

  test('uses pushed toolbar state', async () => {
    const { controller, store, bridge, push, peer } = peerlessCommands(false);
    const renders = mock(() => {});
    renderHook(() => { renders(); usePeerlessWorkerXlsxCommands(controller, bridge); });
    const renderCount = renders.mock.calls.length;
    const release = controller.hold('bold');
    const changed = mock(() => {});
    const unsubscribe = store.subscribe(changed);
    await act(async () => {
      push({ throughSeq: 0, version: 'v0', statuses: [], editState: { formatting: { ...PLAIN_FORMATTING, bold: true } } });
    });
    expect(store.getState('bold').active).toBe(true);
    expect(changed).toHaveBeenCalledTimes(1);
    await act(async () => {
      push({ throughSeq: 0, version: 'v0', statuses: [] });
      push({ throughSeq: 0, version: 'v0', statuses: [], editState: { formatting: { ...PLAIN_FORMATTING, bold: true } } });
    });
    expect(changed).toHaveBeenCalledTimes(1);
    expect(peer.selectionFormatting).not.toHaveBeenCalled();
    expect(renders).toHaveBeenCalledTimes(renderCount);
    unsubscribe(); release();
  });

  test('rapid relative toggles resolve in worker', async () => {
    const { store, sent, ack, peer } = peerlessCommands();
    const first = store.execute('bold', null);
    const second = store.execute('bold', null);
    expect(sent).toHaveLength(2);
    expect(sent.map((command) => command.payload)).toEqual([
      expect.objectContaining({ toolbar: expect.objectContaining({ id: 'bold', args: null }) }),
      expect.objectContaining({ toolbar: expect.objectContaining({ id: 'bold', args: null }) }),
    ]);
    ack(1);
    ack(2);
    expect(await first).toMatchObject({ ok: true });
    expect(await second).toMatchObject({ ok: true });
    expect(store.getState('bold').active).toBe(false);
    expect(peer.selectionFormatting).not.toHaveBeenCalled();
  });

  test('undo redo restore worker selection', async () => {
    const { store, binding, ack } = peerlessCommands();
    const undo = store.execute('undo', null);
    ack(1, { canUndo: false, canRedo: true }, { sheetId: 'sheet:0', range: {
      kind: 'rowCol', start: { row: 4, col: 2 }, end: { row: 4, col: 2 },
    } });
    expect(await undo).toMatchObject({ ok: true });
    expect(binding.environment(false).selection).toMatchObject({ top: 4, left: 2, bottom: 4, right: 2 });
    const redo = store.execute('redo', null);
    ack(2, { canUndo: true, canRedo: false }, { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B9:C10' } });
    expect(await redo).toMatchObject({ ok: true });
    expect(binding.environment(false).selection).toMatchObject({ top: 8, left: 1, bottom: 9, right: 2 });
  });

  test('stale proposal refresh remains usable', async () => {
    const { store, bridge, errors, ack, sent } = peerlessCommands();
    const stale = store.execute('proposalAccept', { proposalId: 'p1' });
    ack(1, undefined, undefined, 'refused');
    expect(await stale).toMatchObject({ ok: false, failure: { code: 'proposal-stale' } });
    expect(bridge.markStale).toHaveBeenCalledWith('p1', ['A1']);
    expect(bridge.refreshProposals).toHaveBeenCalledTimes(1);
    expect(errors).toEqual([]);
    const forced = store.execute('proposalAccept', { proposalId: 'p1', force: true });
    expect(sent[1]).toMatchObject({ kind: 'acceptProposal', payload: { args: ['p1', { force: true }] } });
    ack(2);
    expect(await forced).toMatchObject({ ok: true });
  });

  test('cancels armed formatting locally and captures a fresh token', async () => {
    const { store, bridge, sent, push, peer } = peerlessCommands();
    const capturing = store.execute('paintFormat', null);
    expect(sent[0]).toMatchObject({ kind: 'applyOps', payload: {
      toolbar: { id: 'paintFormat', args: null, target: { sheetId: 'sheet:0' } },
    } });
    push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'applied', result: { formatToken: 'format:1' } }] });
    expect(await capturing).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBe('format:1');
    expect(store.getState('paintFormat').active).toBe(true);
    expect(await store.execute('paintFormat', null)).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBeNull();
    expect(store.getState('paintFormat').active).toBe(false);
    expect(sent).toHaveLength(1);
    const recapturing = store.execute('paintFormat', null);
    expect(sent).toHaveLength(2);
    push({ throughSeq: 2, version: 'v2', statuses: [{ seq: 2, status: 'applied', result: { formatToken: 'format:2' } }] });
    expect(await recapturing).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBe('format:2');
    expect(store.getState('paintFormat').active).toBe(true);
    expect(peer.captureFormat).not.toHaveBeenCalled();
  });

  test('ignores a capture reply after local cancellation', async () => {
    const { store, bridge, sent, push } = peerlessCommands();
    const capturing = store.execute('paintFormat', null);
    bridge.setFormatToken(null);
    push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'applied', result: { formatToken: 'format:1' } }] });
    expect(await capturing).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBeNull();
    expect(store.getState('paintFormat').active).toBe(false);
    expect(sent).toHaveLength(1);
  });

  test('ignores an older capture reply while a newer capture is pending', async () => {
    const { store, bridge, sent, push } = peerlessCommands();
    const first = store.execute('paintFormat', null);
    const second = store.execute('paintFormat', null);
    expect(sent).toHaveLength(2);
    push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'applied', result: { formatToken: 'format:1' } }] });
    expect(await first).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBeNull();
    push({ throughSeq: 2, version: 'v2', statuses: [{ seq: 2, status: 'applied', result: { formatToken: 'format:2' } }] });
    expect(await second).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBe('format:2');
  });

  test('ignores an older capture reply after a newer token is applied', async () => {
    const { store, bridge, coordinator, sent, push } = peerlessCommands();
    const first = store.execute('paintFormat', null);
    const second = store.execute('paintFormat', null);
    push({ throughSeq: 2, version: 'v2', statuses: [{ seq: 2, status: 'applied', result: { formatToken: 'format:2' } }] });
    expect(await second).toMatchObject({ ok: true });
    const applied = coordinator.command({ kind: 'applyFormat', payload: {
      target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'D4' } }, formatToken: bridge.formatToken()!,
    } });
    expect(sent[2]).toMatchObject({ kind: 'applyFormat', payload: { formatToken: 'format:2' } });
    push({ throughSeq: 3, version: 'v3', statuses: [{ seq: 3, status: 'applied' }] });
    await applied.result;
    bridge.setFormatToken(null);
    push({ throughSeq: 3, version: 'v3', statuses: [{ seq: 1, status: 'applied', result: { formatToken: 'format:1' } }] });
    expect(await first).toMatchObject({ ok: true });
    expect(bridge.formatToken()).toBeNull();
    expect(store.getState('paintFormat').active).toBe(false);
  });

  test('rejecting a stale proposal listing reports proposal-not-found', async () => {
    const { store, sent, push, bridge, errors } = peerlessCommands();
    const rejecting = store.execute('proposalReject', { proposalId: 'p1' });
    expect(sent[0]).toMatchObject({ kind: 'rejectProposal', payload: { args: ['p1'] } });
    push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'noop', result: false }] });
    expect(await rejecting).toMatchObject({ ok: false, failure: { code: 'proposal-not-found' } });
    expect(bridge.refreshProposals).toHaveBeenCalledTimes(1);
    expect(errors).toEqual([]);
  });
});

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
