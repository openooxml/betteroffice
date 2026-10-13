import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { initWasm, type CommandStatus, type WorkbookSession, type WorkbookSharedCommand } from '@betteroffice/xlsx';
import { isClientMessage, isHostMessage, SessionFailure, type SessionTransport } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createWorkbookSession } from '../../../xlsx/src/session/client';
import { WORKBOOK_SESSION_METHODS, type WorkbookSessionMethods } from '../../../xlsx/src/session/methods';
import { capturedWorkbookCommand, createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { createWorkbookSessionHost } from '../../../xlsx/src/session/host';
import {
  createWorkbookRecoverySession, workbookRecoveryBackend, workbookSessionRecovery,
} from '../../../xlsx/src/session/clientRecovery';
import { createPeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import { createXlsxCommandController } from '../commands/createXlsxCommandStore';
import { createAsyncEditorApi } from './createAsyncEditorApi';
import { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import { WorkerDraftController } from './WorkerDraftController';

const calculation = { nowSerial: 46_000.5, randSeed: 123456789 };

beforeAll(async () => {
  await initWasm(await readFile(new URL('../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

function crashableHost(): {
  wrap(transport: SessionTransport): SessionTransport;
  crash(): void;
} {
  let crash: ((error: unknown) => void) | undefined;
  return {
    wrap: (transport) => ({
      ...transport,
      onError(listener) { crash = listener; return transport.onError(listener); },
    }),
    crash() {
      if (!crash) throw new Error('Missing worker crash callback');
      crash(new SessionFailure('crash', 'Worker stopped'));
    },
  };
}

function recoveryWorker(change?: (status: CommandStatus) => CommandStatus,
  transports: () => { client: SessionTransport; host: SessionTransport } = createInProcessPair) {
  const open = spyOn(workbookRecoveryBackend, 'open').mockImplementation(async (...args) => {
    const pair = transports();
    createWorkbookSessionHost(pair.host);
    const recovery = await (Reflect.apply(createWorkbookRecoverySession, undefined,
      [args[0], args[1], pair.client, ...args.slice(2)]) as ReturnType<typeof createWorkbookRecoverySession>);
    return change ? { ...recovery, replay: async (command: WorkbookSharedCommand, pending?: boolean) => change(await recovery.replay(command, pending)) } : recovery;
  });
  return open;
}

function recoveryCommand(session: WorkbookSession, seq: number, row: number, input: string) {
  return { ...session.commandIdentity, seq, basis: session.confirmedBasis, kind: 'editCell' as const,
    payload: { args: [0, row, 0, input] as [number, number, number, string],
      target: { sheetId: session.state.sheets[0].id, range: { kind: 'a1' as const, a1: `A${row + 1}` } } } };
}

function recoveryEditor(session: WorkbookSession) {
  const run = new PeerlessWorkbookSession(session, 1, {
    changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
  });
  const api = createAsyncEditorApi(run, createXlsxCommandController().store,
    () => { throw new Error('Unexpected editor bridge'); });
  run.firstPaint();
  return { run, api };
}

function crashAndWait(session: WorkbookSession, host: ReturnType<typeof crashableHost>): Promise<void> {
  if (session.failure) { host.crash(); return Promise.resolve(); }
  const failed = new Promise<void>((resolve) => {
    const off = session.onFailure(() => { off(); resolve(); });
  });
  host.crash();
  return failed;
}

async function recoveredInputs(bytes: Uint8Array, range: string) {
  const reopened = await createTestWorkbookSession(bytes, undefined, { calculation });
  try { return (await reopened.call.cellInputs(0, range)).cells.map((row) => row[0]?.input ?? ''); }
  finally { await reopened.dispose(); }
}

test('public recovery includes a completed held prefill before the read hold releases', async () => {
  const zip = await JSZip.loadAsync(await initialPaintBook());
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', xml.replace('<row r="20">',
    '<row r="11"><c r="A11" t="inlineStr"><is><t>Q3 Sales Dashboard</t></is></c></row><row r="20">'));
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await zip.generateAsync({ type: 'uint8array' }), host.wrap,
    { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  let resolve!: (value: string) => void;
  const prefill = new Promise<string>((done) => { resolve = done; });
  let readingStarted!: () => void;
  const entered = new Promise<void>((done) => { readingStarted = done; });
  const controller = new WorkerDraftController({
    generation: () => 1, identity: () => 'book', navigation: () => 0,
    selection: () => ({ anchor: { row: 10, col: 0 }, focus: { row: 10, col: 0 } }), sheet: () => 0,
    target: () => ({ sheetId: session.state.sheets[0].id, range: { kind: 'a1', a1: 'A11' } }),
    resolveSheet: () => 0, editable: () => true, prefill: () => { readingStarted(); return prefill; },
    select: () => {}, selectionTarget: () => undefined, dependency: () => null,
    navigate: () => {}, committed: () => {}, fail: () => {}, report: () => {}, focus: () => {},
  });
  const coordinator = createPeerlessWorkerInputCoordinator({
    generation: () => 1, session: () => run.session, isReady: () => run.ready,
    target: () => ({ sheetId: session.state.sheets[0].id, range: { kind: 'a1', a1: 'A11' } }),
    seal: controller.seal, sync: controller.sync,
  });
  controller.attach(coordinator);
  run.connectRecoveryInput(coordinator, controller);
  try {
    controller.open();
    controller.holdCellEdit({ kind: 'insert', text: 'X' });
    const pending = controller.seal().pendingInput!;
    const failed = pending.ready.then(() => crashAndWait(session, host));
    coordinator.reserveInput(pending);
    const saving = coordinator.save();
    void saving.catch(() => {});
    await entered;
    resolve('Q3 Sales Dashboard');
    await failed;
    await saving.catch(() => {});
    expect(coordinator.recoveryCommands.filter((command) => command.kind === 'editCell').length +
      workbookSessionRecovery.get(session)!.retainedCommands).toBe(1);
    expect(await recoveredInputs((await api.recoverySave()).bytes, 'A11')).toEqual(['Q3 Sales DashboardX']);
  } finally { controller.dispose(); coordinator.retire(); opener.mockRestore(); await run.dispose(); }
});

function recoveryPair() {
  const listeners = [new Set<(message: unknown) => void>(), new Set<(message: unknown) => void>()];
  const calls: string[] = [];
  const endpoint = (side: number): SessionTransport => ({
    post(message, transfer = []) {
      if (isClientMessage(message) && message.kind === 'call') calls.push(message.method);
      const cloned = structuredClone(message, { transfer });
      queueMicrotask(() => { for (const listener of listeners[1 - side]) listener(cloned); });
    },
    listen(listener) { listeners[side].add(listener); return () => { listeners[side].delete(listener); }; },
    onError: () => () => {},
    close() { for (const group of listeners) group.clear(); },
  });
  return { client: endpoint(0), host: endpoint(1), calls };
}

test('collaborative editors edit and undo past the recovery checkpoint threshold', async () => {
  const pair = recoveryPair();
  createWorkbookSessionHost(pair.host);
  const session = await createWorkbookSession(await initialPaintBook(),
    { collaborative: true, clientId: 101, includeEditState: true }, pair.client);
  const scheduled: (() => void)[] = [];
  const run = new PeerlessWorkbookSession(session, 1, {
    changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
    checkpointScheduler: { now: () => 0,
      schedule(callback) { scheduled.push(callback); return () => {}; } },
  });
  run.connectRecoveryInput({ draft: null, pending: false }, { compositionId: undefined, subscribe: () => () => {} });
  run.firstPaint();
  try {
    for (let seq = 1; seq <= 64; seq += 1) {
      expect(await run.session.submit(recoveryCommand(session, seq, 10, String(seq))))
        .toMatchObject({ status: 'applied', committed: true });
    }
    for (let seq = 65; seq <= 128; seq += 1) {
      expect(await run.session.submit({ ...session.commandIdentity, seq, basis: session.confirmedBasis,
        kind: 'undo', payload: { args: [] } })).toMatchObject({ status: 'applied', committed: true });
    }
    expect(await run.session.submit({ ...session.commandIdentity, seq: 129, basis: session.confirmedBasis,
      kind: 'redo', payload: { args: [] } })).toMatchObject({ status: 'applied', committed: true });
    expect((await run.session.call.cellInputs(0, 'A11')).cells[0][0].input).toBe('1');
    expect(workbookSessionRecovery.has(session)).toBe(false);
    expect(scheduled).toHaveLength(0);
    expect(pair.calls).not.toContain('checkpoint');
    expect(session.failure).toBeUndefined();
    expect(run.failure).toBeNull();
    expect(run.readOnly).toBe(false);
  } finally { await run.dispose(); }
});

test('collaborative editors do not schedule idle recovery checkpoints', async () => {
  const pair = recoveryPair();
  createWorkbookSessionHost(pair.host);
  const session = await createWorkbookSession(await initialPaintBook(),
    { collaborative: true, clientId: 101, includeEditState: true }, pair.client);
  const scheduled: (() => void)[] = [];
  const run = new PeerlessWorkbookSession(session, 1, {
    changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
    checkpointScheduler: { now: () => 0,
      schedule(callback) { scheduled.push(callback); return () => {}; } },
  });
  run.connectRecoveryInput({ draft: null, pending: false }, { compositionId: undefined, subscribe: () => () => {} });
  run.firstPaint();
  try {
    for (let seq = 1; seq <= 64; seq += 1) {
      expect(await run.session.submit({ ...session.commandIdentity, seq, basis: session.confirmedBasis,
        kind: 'setReadOnly', payload: { readOnly: false } })).toMatchObject({ status: 'noop' });
    }
    expect(scheduled).toHaveLength(0);
    expect(pair.calls).not.toContain('checkpoint');
    expect(workbookSessionRecovery.has(session)).toBe(false);
    expect(run.ready).toBe(true);
  } finally { await run.dispose(); }
});

test('collaborative recovery save refuses explicitly and keeps initial remote edits usable', async () => {
  const bytes = await initialPaintBook();
  const sourcePair = recoveryPair();
  createWorkbookSessionHost(sourcePair.host);
  const source = await createWorkbookSession(bytes, { collaborative: true, clientId: 101 }, sourcePair.client);
  let run: PeerlessWorkbookSession | undefined;
  const opener = recoveryWorker();
  try {
    await source.submit(recoveryCommand(source, 1, 10, 'remote'));
    const initialUpdate = await source.replica!.async.encodeStateAsUpdate();
    const pair = recoveryPair();
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(bytes,
      { collaborative: true, clientId: 202, includeEditState: true, initialUpdate }, pair.client);
    const editor = recoveryEditor(session);
    run = editor.run;
    expect((await run.session.call.cellInputs(0, 'A11')).cells[0][0].input).toBe('remote');
    await expect(editor.api.recoverySave()).rejects.toThrow('Recovery is unavailable for collaborative workbooks');
    expect(opener).not.toHaveBeenCalled();
    expect(editor.api.failure).toBeNull();
    expect(editor.api.hydrated).toBe(true);
    expect(run.readOnly).toBe(false);
    expect(session.failure).toBeUndefined();
    expect(session.state.stage).toBe('ready');
    expect(await run.session.submit(recoveryCommand(session, 1, 11, 'local')))
      .toMatchObject({ status: 'applied', committed: true });
    expect(await run.session.submit({ ...session.commandIdentity, seq: 2, basis: session.confirmedBasis,
      kind: 'undo', payload: { args: [] } })).toMatchObject({ status: 'applied', committed: true });
    expect((await run.session.call.cellInputs(0, 'A11')).cells[0][0].input).toBe('remote');
    expect((await run.session.call.cellInputs(0, 'A12')).cells[0][0].input).toBe('');
  } finally { opener.mockRestore(); await run?.dispose(); await source.dispose(); }
});

test('recovery reopen refuses collaborative source workbooks', async () => {
  const pair = recoveryPair();
  createWorkbookSessionHost(pair.host);
  await expect(createWorkbookRecoverySession(await initialPaintBook(), { collaborative: true, clientId: 101 }, pair.client))
    .rejects.toThrow('Recovery is unavailable for collaborative workbooks');
  expect(pair.calls).not.toContain('open');
});

test('recovery reopen refuses collaborative checkpoints without poisoning the replacement', async () => {
  const originalPair = recoveryPair();
  const replacementPair = recoveryPair();
  createWorkbookSessionHost(originalPair.host);
  createWorkbookSessionHost(replacementPair.host);
  const original = createSessionClient<WorkbookSessionMethods, {}>(originalPair.client, { methods: WORKBOOK_SESSION_METHODS });
  const replacement = createSessionClient<WorkbookSessionMethods, {}>(replacementPair.client, { methods: WORKBOOK_SESSION_METHODS });
  try {
    const bytes = await initialPaintBook();
    await original.call.open(new Uint8Array(bytes).buffer, { collaborative: true, clientId: 101, includeEditState: true });
    const checkpoint = await original.call.checkpoint();
    await expect(replacement.call.openRecovery(checkpoint))
      .rejects.toThrow('Recovery is unavailable for collaborative workbooks');
    expect(replacement.failure).toBeUndefined();
    expect(await replacement.call.open(new Uint8Array(bytes).buffer)).toMatchObject({ stage: 'ready' });
  } finally { await original.dispose(); await replacement.dispose(); }
});

test.each(['undo', 'redo'] as const)('public recovery replays a withheld no-op %s after a checkpoint', async (kind) => {
  const host = crashableHost();
  const opener = recoveryWorker();
  let hold = false;
  let held!: () => void;
  const withheld = new Promise<void>((resolve) => { held = resolve; });
  const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => host.wrap({
    ...transport, post(message, transfer) {
      if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok &&
        message.value && typeof message.value === 'object' && 'shared' in message.value) {
        expect(message.value).toMatchObject({ statuses: [{ seq: 2, status: 'noop', committed: false }] });
        held();
        return;
      }
      transport.post(message, transfer);
    },
  }), { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    await run.session.submit({ ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
      kind: 'setReadOnly', payload: { readOnly: false } });
    const recovery = workbookSessionRecovery.get(session)!;
    expect(await recovery.capture(() => true)).toBe(true);
    expect(recovery.throughSeq).toBe(1);
    const expected = await session.save();
    hold = true;
    const pending = run.session.submit({ ...session.commandIdentity, seq: 2,
      basis: session.confirmedBasis, kind, payload: { args: [] } });
    void pending.catch(() => {});
    await withheld;
    await crashAndWait(session, host);
    await expect(pending).rejects.toMatchObject({ name: 'XlsxWorkerSessionFailedError' });
    const saved = await api.recoverySave();
    expect(saved).toEqual({ bytes: expected, recovery: true });
    expect(opener.mock.calls[0][2]?.throughSeq).toBe(1);
    expect(run.readOnly).toBe(true);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test.each(['source', 'journal', 'checkpoint'] as const)(
  'public recovery retains edits queued behind an unresolved clipboard read after %s', async (previous) => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  const input = createPeerlessWorkerInputCoordinator({
    session: () => run.session, generation: () => 1, isReady: () => run.ready,
    seal: () => ({}), sync: () => {},
    target: (_sheet, row) => ({ sheetId: session.state.sheets[0].id, range: { kind: 'a1', a1: `A${row + 1}` } }),
  });
  run.connectRecoveryInput(input, { compositionId: undefined, subscribe: () => () => {} });
  try {
    if (previous !== 'source') {
      await run.session.submit(recoveryCommand(session, 1, 10, 'earlier'));
      if (previous === 'checkpoint') await workbookSessionRecovery.get(session)!.capture(() => true);
    }
    const paste = input.clipboard(() => new Promise<string>(() => {}), (value) => ({
      kind: 'editCell', payload: { args: [0, 10, 0, value] },
    }));
    const edit = input.submit({ generation: 1, sheet: 0, row: 11, col: 0, source: 'cell', value: 'accepted' });
    expect(edit.seq).toBeNull();
    await crashAndWait(session, host);
    await expect(edit.result).rejects.toThrow();
    await expect(paste.result).rejects.toThrow();
    const saved = await api.recoverySave();
    expect(await recoveredInputs(saved.bytes, 'A12')).toEqual(['accepted']);
    const again = await api.recoverySave();
    expect(await recoveredInputs(again.bytes, 'A12')).toEqual(['accepted']);
    if (previous !== 'source') expect(await recoveredInputs(saved.bytes, 'A11')).toEqual(['earlier']);
  } finally { input.retire(); opener.mockRestore(); await run.dispose(); }
});

test('public recovery restores 300 awaited edits without input idle', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const options = { changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
    checkpointScheduler: { now: () => 0, schedule: (_callback: () => void, _delay: number) => () => {} } };
  const run = new PeerlessWorkbookSession(session, 1, options);
  const api = createAsyncEditorApi(run, createXlsxCommandController().store,
    () => { throw new Error('Unexpected editor bridge'); });
  run.connectRecoveryInput({ draft: null, pending: true }, { compositionId: undefined, subscribe: () => () => {} });
  run.firstPaint();
  try {
    for (let seq = 1; seq <= 300; seq += 1) {
      expect(await run.session.submit(recoveryCommand(session, seq, seq + 9, String(seq))))
        .toMatchObject({ status: 'applied', committed: true });
    }
    expect(workbookSessionRecovery.get(session)!.throughSeq).toBeGreaterThanOrEqual(192);
    await crashAndWait(session, host);
    const saved = await api.recoverySave();
    expect(await recoveredInputs(saved.bytes, 'A11:A310')).toEqual(Array.from({ length: 300 }, (_, index) => String(index + 1)));
    expect(run.readOnly).toBe(true);
    expect(api.hydrated).toBe(false);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test('public recovery saves edits before the first checkpoint', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    await run.session.submit(recoveryCommand(session, 1, 10, 'first'));
    await run.session.submit(recoveryCommand(session, 2, 11, 'second'));
    expect(workbookSessionRecovery.get(session)!.throughSeq).toBe(0);
    await crashAndWait(session, host);
    const saved = await api.recoverySave();
    expect(saved.recovery).toBe(true);
    expect(await recoveredInputs(saved.bytes, 'A11:A12')).toEqual(['first', 'second']);
    expect(run.readOnly).toBe(true);
    expect(api.hydrated).toBe(false);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test.each(['undo', 'redo'] as const)('public recovery preserves %s after checkpoint 64', async (kind) => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    for (let seq = 1; seq <= (kind === 'undo' ? 64 : 63); seq += 1) {
      await run.session.submit(recoveryCommand(session, seq, 10, String(seq)));
    }
    if (kind === 'redo') await run.session.submit({ ...session.commandIdentity, seq: 64,
      basis: session.confirmedBasis, kind: 'undo', payload: { args: [] } });
    const recovery = workbookSessionRecovery.get(session)!;
    expect(await recovery.capture(() => true)).toBe(true);
    expect(recovery.throughSeq).toBe(64);
    expect(recovery.retainedCommands).toBe(0);
    expect(await run.session.submit({ ...session.commandIdentity, seq: 65,
      basis: session.confirmedBasis, kind, payload: { args: [] } })).toMatchObject({ status: 'applied', committed: true });
    await crashAndWait(session, host);
    expect(await recoveredInputs((await api.recoverySave()).bytes, 'A11')).toEqual(['63']);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test('public recovery restores sheet identities when undo crosses a checkpoint', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    for (let seq = 1; seq <= 63; seq += 1) await run.session.submit(recoveryCommand(session, seq, 10, String(seq)));
    const sheet = session.state.sheets[0].id;
    await run.session.submit({ ...session.commandIdentity, seq: 64, basis: session.confirmedBasis,
      kind: 'applyOps', payload: { args: [[{ type: 'addSheet', index: 0, name: 'New' }]] } });
    await workbookSessionRecovery.get(session)!.capture(() => true);
    await run.session.submit({ ...session.commandIdentity, seq: 65, basis: session.confirmedBasis,
      kind: 'undo', payload: { args: [] } });
    expect(session.selection?.sheetId).toBe(sheet);
    await run.session.submit(recoveryCommand(session, 66, 11, 'after undo'));
    await crashAndWait(session, host);
    expect(await recoveredInputs((await api.recoverySave()).bytes, 'A11:A12')).toEqual(['63', 'after undo']);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test('public recovery accepts matching refusals and dependency cancellations', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const zip = await JSZip.loadAsync(await initialPaintBook());
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', xml.replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells></worksheet>'));
  const session = await createTestWorkbookSession(await zip.generateAsync({ type: 'uint8array' }), host.wrap,
    { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    await run.session.submit(recoveryCommand(session, 1, 10, 'first'));
    expect(await run.session.submit({ ...recoveryCommand(session, 2, 0, 'locked'), payload: { args: [0, 0, 4, 'locked'],
      target: { sheetId: session.state.sheets[0].id, range: { kind: 'a1', a1: 'E1' } } } }))
      .toMatchObject({ status: 'refused', reason: 'locked-target', committed: false });
    expect(await run.session.submit({ ...recoveryCommand(session, 3, 11, 'cancelled'), afterSeq: 2 }))
      .toMatchObject({ status: 'cancelled', reason: 'dependency-refused', committed: false });
    await run.session.submit(recoveryCommand(session, 4, 11, 'second'));
    await crashAndWait(session, host);
    expect(await recoveredInputs((await api.recoverySave()).bytes, 'A11:A12')).toEqual(['first', 'second']);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test('public recovery rejects changed refusal outcomes', async () => {
  const host = crashableHost();
  const opener = recoveryWorker((status) => status.seq === 2 ? { ...status, reason: 'stale-basis' } : status);
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    await run.session.submit(recoveryCommand(session, 1, 10, 'first'));
    expect(await run.session.submit({ ...recoveryCommand(session, 2, 11, 'refused'),
      payload: { args: [0, 11, 0, 'refused'], target: { sheetId: 'missing', range: { kind: 'a1', a1: 'A12' } } } }))
      .toMatchObject({ status: 'refused', reason: 'target-sheet-missing', committed: false });
    await run.session.submit(recoveryCommand(session, 3, 11, 'second'));
    await crashAndWait(session, host);
    await expect(api.recoverySave()).rejects.toMatchObject({ cause: {
      code: 'workbook-recovery-refused', seq: 2, reason: 'stale-basis',
    } });
    expect(run.readOnly).toBe(true);
    expect(api.failure).toBe(run.failure);
  } finally { opener.mockRestore(); await run.dispose(); }
});

test.each([false, true])('public recovery survives a withheld checkpoint reply (previous=%s)', async (previous) => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const requests = new Set<number>();
  let hold = false;
  let held!: () => void;
  const withheld = new Promise<void>((resolve) => { held = resolve; });
  let editRequest: number | undefined;
  let dispatch!: () => void;
  const dispatched = new Promise<void>((resolve) => { dispatch = resolve; });
  let acknowledge!: () => void;
  const acknowledged = new Promise<void>((resolve) => { acknowledge = resolve; });
  let deliver!: () => void;
  const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => host.wrap({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'checkpoint') requests.add(message.id);
      if (isClientMessage(message) && message.kind === 'call' && capturedWorkbookCommand(message)?.seq === 3) {
        editRequest = message.id;
        dispatch();
      }
      listener(message);
    }),
    post(message, transfer) {
      if (hold && isHostMessage(message) && message.kind === 'reply' && requests.has(message.id)) { held(); return; }
      if (isHostMessage(message) && message.kind === 'reply' && message.id === editRequest) {
        deliver = () => transport.post(message, transfer);
        acknowledge();
        return;
      }
      transport.post(message, transfer);
    },
  }), { includeEditState: true, calculation });
  const { run, api } = recoveryEditor(session);
  try {
    await run.session.submit(recoveryCommand(session, 1, 10, 'kept'));
    const recovery = workbookSessionRecovery.get(session)!;
    if (previous) await recovery.capture(() => true);
    await run.session.submit({ ...session.commandIdentity, seq: 2, basis: session.confirmedBasis,
      kind: 'applyOps', payload: { args: [[{ type: 'insertRows', sheet: 0, at: 10, count: 1 }]] } });
    hold = true;
    const capture = recovery.capture(() => true);
    void capture.catch(() => {});
    await withheld;
    const edit = run.session.submit(recoveryCommand(session, 3, 10, 'later'));
    void edit.catch(() => {});
    await dispatched;
    await acknowledged;
    expect(session.throughSeq).toBe(2);
    deliver();
    await edit;
    expect(session.throughSeq).toBe(3);
    await crashAndWait(session, host);
    await Promise.allSettled([capture, edit]);
    expect(await recoveredInputs((await api.recoverySave()).bytes, 'A11:A13')).toEqual(['later', 'kept', '']);
    expect(opener.mock.calls[0][2]?.throughSeq ?? 0).toBe(previous ? 1 : 0);
    expect(run.readOnly).toBe(true);
  } finally { opener.mockRestore(); host.crash(); await run.dispose(); }
});

test.each(['draft', 'composition'] as const)('public recovery keeps retained edits at the cap with an open %s', async (blocked) => {
  const host = crashableHost();
  // 768 sequential round trips: setTimeout(0) delivery would spend seconds in timer clamps.
  const opener = recoveryWorker(undefined, recoveryPair);
  const pair = recoveryPair();
  createWorkbookSessionHost(host.wrap(pair.host));
  const session = await createWorkbookSession(await initialPaintBook(), { includeEditState: true, calculation }, pair.client);
  const { run, api } = recoveryEditor(session);
  run.connectRecoveryInput({ draft: blocked === 'draft'
    ? { generation: 1, sheet: 0, row: 10, col: 0, value: 'draft', source: 'cell' } : null, pending: false },
  { compositionId: blocked === 'composition' ? 'composing' : undefined, subscribe: () => () => {} });
  try {
    await run.session.submit(recoveryCommand(session, 1, 10, 'kept'));
    for (let seq = 2; seq <= 256; seq += 1) await run.session.submit({ ...session.commandIdentity, seq,
      basis: session.confirmedBasis, kind: 'applyOps',
      payload: { args: [[{ type: 'insertRows', sheet: 0, at: 10, count: 1 }]] } });
    await expect(run.session.submit(recoveryCommand(session, 257, 10, 'overflow')))
      .rejects.toMatchObject({ cause: { seq: 257, reason: 'log-full' } });
    const failure = api.failure;
    expect(failure).toBeInstanceOf(Error);
    expect(run.readOnly).toBe(true);
    await crashAndWait(session, host);
    const saved = await api.recoverySave();
    expect(await recoveredInputs(saved.bytes, 'A265:A267')).toEqual(['', 'kept', '']);
    expect((await api.recoverySave()).bytes).toEqual(saved.bytes);
    expect(api.failure).toBe(failure);
    expect(api.hydrated).toBe(false);
    expect(run.readOnly).toBe(true);
    await expect(run.session.submit(recoveryCommand(session, 257, 10, 'blocked'))).rejects.toBe(failure);
  } finally { opener.mockRestore(); await run.dispose(); }
});

async function initialPaintBook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="D6" state="frozen"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><f>NOW()</f><v>0</v></c><c r="B1"><f>RAND()</f><v>0</v></c></row><row r="20"><c r="Z20"><v>7</v></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}
