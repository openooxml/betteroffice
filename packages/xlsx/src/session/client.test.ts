import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import {
  isClientMessage, isHostMessage, SessionFailure, type SessionTransport,
} from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { XlsxCellRead } from '../edits';
import * as workbookWasm from '../wasm/loader';
import { wasmAssetUrl } from '../wasm/asset';
import { XlsxDocument } from '../wasm/generated/xlsx_wasm.js';
import { openWorkbook, type WorkbookCalculationContext, type WorkbookHandle } from '../wasm/loader';
import { createWorkbookSession, hydratePeer, type WorkbookSession } from './client';
import { workbookPeerSources } from './clientInternals';
import { createWorkbookEditPeer, WorkbookEditPeerFailedError, type WorkbookEditPeer } from './editPeer';
import { workbookEditPeerOperations } from './editPeerInternals';
import { createWorkbookSessionHost } from './host';
import { WORKBOOK_SESSION_POLICIES } from './methods';
import type { WorkbookSharedCommand } from './commandProtocol';
import { WorkbookPeerHydrationError } from './peerHydrationError';
import { workbookSessionInternals, type WorkbookReplayEnvelope, type WorkbookReplayOp } from './replay';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';
import {
  createWorkbookRecoverySession, workbookRecoveryBackend, workbookSessionRecovery, WorkbookRecoveryReplayError,
} from './clientRecovery';

let fixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };

beforeAll(async () => {
  ({ fixture, wasmBytes } = await loadWorkbookSessionFixtures());
});

async function volatileWorkbookBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Volatile" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>NOW()</f><v>0</v></c><c r="B1"><f>TODAY()</f><v>0</v></c><c r="C1"><f>RAND()</f><v>0</v></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

for (const retainPeerHydration of [false, true]) {
  test(`${retainPeerHydration ? 'hybrid' : 'viewer'} startup works without crypto.randomUUID`, async () => {
    const descriptor = Object.getOwnPropertyDescriptor(globalThis.crypto, 'randomUUID');
    Object.defineProperty(globalThis.crypto, 'randomUUID', { configurable: true, value: undefined });
    const random = spyOn(globalThis.crypto, 'getRandomValues');
    let session: WorkbookSession | undefined;
    try {
      session = await createTestWorkbookSession(fixture, undefined, {
        calculation, wasm: wasmBytes.buffer, retainPeerHydration,
      });
      expect(session.state.stage).toBe('ready');
      expect(session.commandIdentity.session).toMatch(/^[0-9a-f]{32}$/);
      expect(session.commandIdentity.incarnation).toMatch(/^[0-9a-f]{32}$/);
      expect(session.commandIdentity.session).not.toBe(session.commandIdentity.incarnation);
      expect(random.mock.calls.filter(([value]) => value instanceof Uint8Array && value.length === 16).length).toBeGreaterThanOrEqual(2);
      expect((await session.call.sheets()).length).toBeGreaterThan(0);
      expect((await session.save()).length).toBeGreaterThan(0);
      expect(session.failure).toBeUndefined();
    } finally {
      random.mockRestore();
      if (descriptor) Object.defineProperty(globalThis.crypto, 'randomUUID', descriptor);
      else Reflect.deleteProperty(globalThis.crypto, 'randomUUID');
      await session?.dispose();
    }
  });
}

test('ack delivery skips removed listeners and defers replacements until the next ack', async () => {
  const session = await createTestWorkbookSession(fixture);
  const delivered: string[] = [];
  let offCurrent = () => {};
  let offRemoved = () => {};
  offCurrent = session.onAck((ack) => {
    delivered.push(`current:${ack.throughSeq}`);
    offCurrent();
    offRemoved();
    offCurrent = session.onAck((next) => { delivered.push(`replacement:${next.throughSeq}`); });
  });
  offRemoved = session.onAck((ack) => { delivered.push(`removed:${ack.throughSeq}`); });
  try {
    for (const seq of [1, 2]) {
      expect(await session.submit({ ...session.commandIdentity, seq, basis: session.confirmedBasis,
        kind: 'editCell', payload: { args: [0, 0, 0, String(seq + 10)] } }))
        .toMatchObject({ status: 'applied' });
      expect(delivered).toEqual(seq === 1 ? ['current:1'] : ['current:1', 'replacement:2']);
    }
    expect(session.failure).toBeUndefined();
  } finally { offCurrent(); offRemoved(); await session.dispose(); }
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

function recoveryWorker() {
  const open = spyOn(workbookRecoveryBackend, 'open').mockImplementation((...args) => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    return Reflect.apply(createWorkbookRecoverySession, undefined,
      [args[0], args[1], pair.client, ...args.slice(2)]) as ReturnType<typeof createWorkbookRecoverySession>;
  });
  return open;
}

function recoveryCommand(session: WorkbookSession, seq: number, row: number, input: string) {
  return { ...session.commandIdentity, seq, basis: session.confirmedBasis, kind: 'editCell' as const,
    payload: { args: [0, row, 0, input] as [number, number, number, string],
      target: { sheetId: session.state.sheets[0].id, range: { kind: 'a1' as const, a1: `A${row + 1}` } } } };
}

test('recovery saves acknowledged and unacknowledged commands once without a peer', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  let hold = false;
  let held!: () => void;
  const acknowledged = new Promise<void>((resolve) => { held = resolve; });
  const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => host.wrap({
    ...transport, post(message, transfer) {
      if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok &&
        message.value && typeof message.value === 'object' && 'shared' in message.value) { held(); return; }
      transport.post(message, transfer);
    },
  }), { includeEditState: true, calculation });
  try {
    const first = recoveryCommand(session, 1, 10, 'kept');
    expect(await session.submit(first)).toMatchObject({ status: 'applied' });
    await session.submit(first);
    hold = true;
    const pending = session.submit({ ...session.commandIdentity, seq: 2, basis: session.confirmedBasis,
      kind: 'applyOps', payload: { args: [[{ type: 'insertRows', sheet: 0, at: 10, count: 1 }]] } });
    void pending.catch(() => {});
    await acknowledged;
    host.crash();
    await expect(pending).rejects.toBeInstanceOf(SessionFailure);
    const recovery = workbookSessionRecovery.get(session)!;
    expect(recovery.retainedCommands).toBe(2);
    const bytes = await recovery.save(() => {});
    const reopened = await createTestWorkbookSession(bytes, undefined, { calculation });
    try {
      const cells = (await reopened.call.cellInputs(0, 'A11:A13')).cells;
      expect(cells.map((row) => row[0]?.input ?? '')).toEqual(['', 'kept', '']);
    } finally { await reopened.dispose(); }
    expect(await recovery.save(() => {})).toEqual(bytes);
    expect(opener).toHaveBeenCalledTimes(2);
    expect(session.state.stage).toBe('failed');
    await expect(session.submit(recoveryCommand(session, 3, 0, 'blocked'))).rejects.toBeInstanceOf(SessionFailure);
  } finally { opener.mockRestore(); await session.dispose(); }
});

test('recovery trims a checkpoint prefix and replays later commands', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  let hold = false;
  let held!: () => void;
  const acknowledged = new Promise<void>((resolve) => { held = resolve; });
  const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => host.wrap({
    ...transport, post(message, transfer) {
      if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok &&
        message.value && typeof message.value === 'object' && 'shared' in message.value) { held(); return; }
      transport.post(message, transfer);
    },
  }), { includeEditState: true, calculation });
  try {
    await session.submit(recoveryCommand(session, 1, 0, 'first'));
    await session.submit(recoveryCommand(session, 2, 1, 'second'));
    const recovery = workbookSessionRecovery.get(session)!;
    expect(await recovery.capture(() => true)).toBe(true);
    expect(recovery.throughSeq).toBe(2);
    expect(recovery.retainedCommands).toBe(0);
    await session.submit(recoveryCommand(session, 3, 2, 'third'));
    hold = true;
    const pending = session.submit(recoveryCommand(session, 4, 3, 'fourth'));
    void pending.catch(() => {});
    await acknowledged;
    expect(recovery.retainedCommands).toBe(2);
    host.crash();
    await expect(pending).rejects.toBeInstanceOf(SessionFailure);
    const bytes = await recovery.save(() => {});
    expect(opener.mock.calls[0][2]?.throughSeq).toBe(2);
    const reopened = await createTestWorkbookSession(bytes);
    try {
      expect((await reopened.call.cellInputs(0, 'A1:A4')).cells.map((row) => row[0].input))
        .toEqual(['first', 'second', 'third', 'fourth']);
    } finally { await reopened.dispose(); }
  } finally { opener.mockRestore(); await session.dispose(); }
});

test.each([false, true])('recovery preserves volatile calculation context (checkpoint=%s)', async (checkpoint) => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const bytes = await volatileWorkbookBytes();
  const session = await createTestWorkbookSession(bytes, host.wrap, { includeEditState: true, calculation });
  const control = await createTestWorkbookSession(bytes, undefined, { includeEditState: true, calculation });
  try {
    const changed = { nowSerial: 47_000.25, randSeed: 731 };
    for (const active of [session, control]) {
      await active.submit(recoveryCommand(active, 1, 1, '=RAND()+NOW()'));
      if (active === session && checkpoint) await workbookSessionRecovery.get(session)!.capture(() => true);
      await active.submit({ ...recoveryCommand(active, 2, 2, '=RAND()+TODAY()'),
        payload: { ...recoveryCommand(active, 2, 2, '').payload, args: [0, 2, 0, '=RAND()+TODAY()'], calculation: changed } });
    }
    const expected = await control.save();
    host.crash();
    expect(await workbookSessionRecovery.get(session)!.save(() => {})).toEqual(expected);
  } finally { opener.mockRestore(); await session.dispose(); await control.dispose(); }
});

test('recovery refuses an unreplayable admitted command with a typed sequence', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true });
  try {
    const pending = session.submit({ ...recoveryCommand(session, 1, 0, 'kept'),
      payload: { args: [0, 0, 0, 'kept'], target: { sheetId: 'missing', range: { kind: 'a1', a1: 'A1' } } } });
    void pending.catch(() => {});
    host.crash();
    await expect(pending).rejects.toBeInstanceOf(SessionFailure);
    const saved = workbookSessionRecovery.get(session)!.save(() => {});
    await expect(saved).rejects.toBeInstanceOf(WorkbookRecoveryReplayError);
    await expect(saved).rejects.toMatchObject({ code: 'workbook-recovery-refused', seq: 1, reason: 'target-sheet-missing' });
  } finally { opener.mockRestore(); await session.dispose(); }
});

test('recovery retains commands admitted behind a read barrier', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const reads = new Set<number>();
  let held!: () => void;
  const withheld = new Promise<void>((resolve) => { held = resolve; });
  const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => host.wrap({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'cellInputs') reads.add(message.id);
      listener(message);
    }),
    post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && reads.has(message.id)) { held(); return; }
      transport.post(message, transfer);
    },
  }), { includeEditState: true });
  try {
    const read = session.call.cellInputs(0, 'A1');
    const edit = session.submit(recoveryCommand(session, 1, 0, 'queued'));
    void read.catch(() => {}); void edit.catch(() => {});
    await withheld;
    host.crash();
    await Promise.allSettled([read, edit]);
    const bytes = await workbookSessionRecovery.get(session)!.save(() => {});
    const reopened = await createTestWorkbookSession(bytes);
    try { expect((await reopened.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('queued'); }
    finally { await reopened.dispose(); }
  } finally { opener.mockRestore(); await session.dispose(); }
});

test('recovery retains opening bytes and the default calculation context', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const bytes = await volatileWorkbookBytes();
  const session = await createTestWorkbookSession(bytes, host.wrap, { includeEditState: true });
  try {
    await session.submit(recoveryCommand(session, 1, 1, '=RAND()+NOW()'));
    const expected = await session.save();
    bytes.fill(0);
    host.crash();
    expect(await workbookSessionRecovery.get(session)!.save(() => {})).toEqual(expected);
    expect(opener.mock.calls[0][1].calculation).toMatchObject({ nowSerial: expect.any(Number), randSeed: expect.any(Number) });
  } finally { opener.mockRestore(); await session.dispose(); }
});

test('recovery bounds admission without dropping retained commands', async () => {
  const host = crashableHost();
  const opener = recoveryWorker();
  const session = await createTestWorkbookSession(await initialPaintBook(), host.wrap, { includeEditState: true });
  try {
    for (let seq = 1; seq <= 256; seq += 1) await session.submit(recoveryCommand(session, seq, 1, String(seq)));
    const recovery = workbookSessionRecovery.get(session)!;
    expect(recovery.retainedCommands).toBe(256);
    const rejected = session.submit(recoveryCommand(session, 257, 1, 'blocked'));
    await expect(rejected).rejects.toBeInstanceOf(WorkbookRecoveryReplayError);
    await expect(rejected).rejects.toMatchObject({ seq: 257, reason: 'log-full' });
    const expected = await session.save();
    host.crash();
    expect(await recovery.save(() => {})).toEqual(expected);
  } finally { opener.mockRestore(); await session.dispose(); }
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

describe('peerless workbook open', () => {
  test('peerless open retains no source or hydration', async () => {
    const messages: unknown[] = [];
    const session = await createTestWorkbookSession(await initialPaintBook(), (transport) => ({
      ...transport, post(message, transfer) { messages.push(message); transport.post(message, transfer); },
    }), { calculation, wasm: wasmBytes.buffer });
    try {
      expect(workbookPeerSources.has(session)).toBe(false);
      expect(messages.some((message) => isHostMessage(message) && message.kind === 'event' && message.name === 'peerOpened')).toBe(false);
      expect(messages.some((message) => isHostMessage(message) && message.kind === 'wasm-module' && message.hydration !== undefined)).toBe(false);
      const internal = workbookSessionInternals.get(session);
      expect(internal?.attachPeer).toBeUndefined();
      expect(internal?.detachPeer).toBeUndefined();
      expect(internal?.preview).toBeUndefined();
      await expect(hydratePeer(session)).rejects.toBeInstanceOf(TypeError);
    } finally { await session.dispose(); }
  });

  test('known viewport open carries final geometry and frame', async () => {
    const bytes = await initialPaintBook();
    const requests: string[] = [];
    const session = await createTestWorkbookSession(bytes, (transport) => ({
      ...transport, listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call') requests.push(message.method);
        listener(message);
      }),
    }), { calculation, wasm: wasmBytes.buffer,
      initialViewport: { sheet: 0, width: 800, height: 600, zoom: 2, dpr: 2, requestGeneration: 7 } });
    const main = openWorkbook(bytes, { calculation });
    try {
      expect(requests).toEqual(['open']);
      const initial = session.initialPaint!;
      expect(initial).toBeDefined();
      expect(initial.sheetView).toEqual(await session.call.sheetView(0));
      expect(initial.sheetView).toMatchObject({ frozenRows: 1, frozenCols: 1 });
      expect(initial.frame).toMatchObject({ ...session.commandIdentity, coversThroughSeq: 0,
        sequence: 0, requestGeneration: 7, zoom: 2, dpr: 2, sheet: 0,
        viewport: { x: initial.sheetView.initialScrollX, y: initial.sheetView.initialScrollY, width: 400, height: 300 } });
      expect(initial.frame.displayList).toEqual(main.displayList(initial.frame.viewport));
      expect(initial.editState).toMatchObject({ canUndo: false, canRedo: false, readOnly: false, activeSheet: 0 });
      expect(session.editState).toEqual(initial.editState);
      expect(session.markPresented(initial.frame)).toBe(true);
      expect(session.markPresented({ ...session.commandIdentity, coversThroughSeq: 0, displayList: initial.frame.displayList })).toBe(true);
    } finally { main.dispose(); await session.dispose(); }
  });

  test('first toolbar refusal preserves the initial selection without edit projection', async () => {
    const session = await createTestWorkbookSession(await initialPaintBook());
    try {
      const selection = { sheetId: session.state.sheets[session.state.activeSheet].id,
        range: { kind: 'a1' as const, a1: 'A1' } };
      expect(session.selection).toBeUndefined();
      const version = session.confirmedBasis.version;
      const destination = { ...selection, range: { kind: 'a1' as const, a1: 'B1' } };
      expect(await session.submit({ ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
        kind: 'applyOps', payload: {
          toolbar: { id: 'bold', args: { value: true }, target: destination }, selection: destination,
        } })).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      expect(session.selection).toBeUndefined();
      expect(session.confirmedBasis.version).toBe(version);
      expect(session.editState).toMatchObject({ canUndo: false, canRedo: false });
      expect(await session.submit({ ...session.commandIdentity, seq: 2, basis: session.confirmedBasis,
        kind: 'editCell', payload: { args: [session.state.activeSheet, 0, 0, '9'] },
      })).toMatchObject({ status: 'applied', committed: true });
      expect(session.selection).toEqual(selection);
    } finally { await session.dispose(); }
  });

  test('default never initializes main edit wasm', async () => {
    const bytes = await initialPaintBook();
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host, { initWasm: async () => {} });
    const initializing = spyOn(workbookWasm, 'initWasm');
    const snapshot = spyOn(workbookWasm, 'createWorkbookSnapshotBuilder');
    const peer = spyOn(workbookWasm, 'openWorkbookPeer');
    let session;
    try {
      session = await createWorkbookSession(bytes, { includeEditState: true }, pair.client);
      expect(session.editState).toMatchObject({ readOnly: false, activeSheet: 0 });
      expect(initializing).not.toHaveBeenCalled();
      expect(snapshot).not.toHaveBeenCalled();
      expect(peer).not.toHaveBeenCalled();
      expect(workbookPeerSources.has(session)).toBe(false);
    } finally {
      initializing.mockRestore(); snapshot.mockRestore(); peer.mockRestore();
      await session?.dispose();
    }
  });

  test.each([false, true])('open skips unrequested edit projection (legacy=%s)', async (retainPeerHydration) => {
    const opening = workbookWasm.openWorkbook;
    let projections = 0;
    let restoreHistory = () => {};
    const opened = spyOn(workbookWasm, 'openWorkbook').mockImplementation((...args) => {
      const worker = opening(...args);
      const history = worker.historyState;
      const spy = spyOn(worker, 'historyState').mockImplementation(() => {
        projections += 1;
        return history.call(worker);
      });
      restoreHistory = () => spy.mockRestore();
      return worker;
    });
    const replies: unknown[] = [];
    let session;
    try {
      session = await createTestWorkbookSession(await initialPaintBook(), (transport) => ({
        ...transport, post(message, transfer) {
          if (isHostMessage(message) && message.kind === 'reply' && message.ok) replies.push(message.value);
          transport.post(message, transfer);
        },
      }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration, includeEditState: retainPeerHydration });
      expect(projections).toBe(0);
      expect(replies[0]).not.toHaveProperty('editState');
      expect(replies[0]).not.toHaveProperty('selection');
      expect(replies[0]).not.toHaveProperty('initialPaint');
      expect(session.editState).toEqual({});
      if (!retainPeerHydration) {
        const ack = await session.call.command({ ...session.commandIdentity, seq: 1,
          basis: session.confirmedBasis, kind: 'select', payload: {
            selection: { sheetId: session.state.sheets[0].id, range: { kind: 'a1', a1: 'A1' } },
          } });
        expect(Object.keys(ack.editState!).sort()).toEqual([
          'activeSheet', 'calculation', 'canRedo', 'canUndo', 'formatting', 'formattingAvailable',
          'mergedRanges', 'readOnly', 'selectionEditable', 'sheets',
        ]);
        expect(ack.editState).toEqual(session.editState);
        expect(projections).toBe(1);
      }
    } finally { opened.mockRestore(); restoreHistory(); await session?.dispose(); }
  });
});

describe('workbook peer hydration', () => {
  test('refuses retained public mutations before attachment and after peer disposal without changing document state', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(await initialPaintBook(), {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    }, pair.client);
    let peer: WorkbookHandle | undefined;
    let edits: WorkbookEditPeer | undefined;
    const applyEdits = session.call.applyEdits;
    const events: string[] = [];
    try {
      for (const phase of ['before attachment', 'after peer disposal']) {
        if (phase === 'after peer disposal') {
          peer = await hydratePeer(session);
          edits = createWorkbookEditPeer({ session, peer });
          await edits.flush();
          edits.dispose();
        }
        const version = await session.call.version();
        const saved = await session.save();
        const state = structuredClone(session.state);
        const input = {
          expectVersion: version,
          steps: [{
            op: 'setCellInputs' as const,
            target: { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } },
            inputs: [['refused before attachment']],
          }],
        };
        const mutators = Object.entries(WORKBOOK_SESSION_POLICIES)
          .filter(([, policy]) => typeof policy !== 'function' && policy.mutates)
          .map(([method]) => method);
        expect(mutators).toEqual(['applyEdits', 'edit']);
        await expect(session.call.applyEdits(input)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
        await expect(session.call.applyEdits(input)).rejects.toMatchObject({ code: 'mutation-outside-replay' });
        await expect(applyEdits(input)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
        events.push(`${phase}:public-refused`);
        const edit: WorkbookSharedCommand = { session: session.commandIdentity.session,
          incarnation: 1, seq: 1, basis: { epoch: session.confirmedBasis.epoch, version: 0 },
          kind: 'editCell', payload: { payload: { args: [0, 0, 0, 'refused shared edit'] } } };
        for (const [index, method] of mutators.entries()) {
          const id = 10_000 + index;
          const reply = new Promise<unknown>((resolve) => {
            const off = pair.client.listen((message) => {
              if (isHostMessage(message) && message.kind === 'reply' && message.id === id) {
                off();
                resolve(message);
              }
            });
          });
          pair.client.post({ protocol: 1, kind: 'call', id, method, args: [method === 'edit' ? edit : input] });
          expect(await reply).toMatchObject({
            ok: false, error: { name: 'WorkbookPeerHydrationError', refusal: { code: 'mutation-outside-replay' } },
          });
          events.push(`${phase}:${method}:refused`);
        }
        const internal = workbookSessionInternals.get(session);
        if (!internal) throw new Error('Missing internal workbook replay helper');
        const envelope: WorkbookReplayEnvelope = {
          sequence: 1, calculation, op: { method: 'editCell', args: [0, 0, 0, 'refused replay'] },
        };
        await expect(internal.replay(envelope)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
        await expect(internal.replay(envelope)).rejects.toMatchObject({ code: 'mutation-outside-replay' });
        events.push(`${phase}:replay-refused`);
        expect(await session.call.version()).toBe(version);
        expect(await session.save()).toEqual(saved);
        expect(session.state).toMatchObject({ version: 0, dirty: false, stage: 'ready' });
        expect(session.state).toEqual(state);
        expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(0);
        expect(session.failure).toBeUndefined();
        events.push(`${phase}:unchanged`);
      }
      expect(events).toEqual(['before attachment:public-refused', 'before attachment:applyEdits:refused',
        'before attachment:edit:refused', 'before attachment:replay-refused', 'before attachment:unchanged',
        'after peer disposal:public-refused', 'after peer disposal:applyEdits:refused',
        'after peer disposal:edit:refused', 'after peer disposal:replay-refused', 'after peer disposal:unchanged']);
    } finally {
      edits?.dispose();
      peer?.dispose();
      await session.dispose();
    }
  });

  test('refuses retained peer attachment after a worker-side state change with a typed error', async () => {
    const opening = workbookWasm.openWorkbook;
    let worker: WorkbookHandle | undefined;
    const opened = spyOn(workbookWasm, 'openWorkbook').mockImplementation((...args) => {
      worker = opening(...args);
      return worker;
    });
    let session;
    try {
      session = await createTestWorkbookSession(fixture, undefined, {
        calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
      });
    } finally { opened.mockRestore(); }
    const peer = await hydratePeer(session);
    if (!worker) throw new Error('Missing worker workbook');
    const version = peer.version();
    expect(worker.editCell(0, 0, 0, 'worker-side change').applied).toBe(true);
    expect(worker.version()).not.toBe(version);
    const edits = createWorkbookEditPeer({ session, peer });
    try {
      await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(edits.error).toBeInstanceOf(WorkbookPeerHydrationError);
      expect(edits.error).toMatchObject({ code: 'version-mismatch' });
      expect(edits.state).toBe('failed');
      expect(edits.sentSequence).toBe(0);
      expect(peer.version()).toBe(version);
      expect(peer.cell(0, 0, 0).input).not.toBe('worker-side change');
      expect(session.failure).toBeUndefined();
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('refuses retained peer attachment when the hydration sequence differs from the worker', async () => {
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'wasm-module' && message.hydration !== undefined) {
          message = { ...message, sequence: 1 };
        }
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    const peer = await hydratePeer(session);
    const edits = createWorkbookEditPeer({ session, peer });
    try {
      expect(peer.version()).toBe(await session.call.version());
      await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(edits.error).toBeInstanceOf(WorkbookPeerHydrationError);
      expect(edits.error).toMatchObject({ code: 'version-mismatch' });
      expect(edits.sentSequence).toBe(0);
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(0);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('refuses queued retained replays after failed attachment without worker mutation', async () => {
    const requests: string[] = [];
    const refusals: unknown[] = [];
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call') requests.push(message.method);
        listener(message);
      }),
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'wasm-module') {
          message = { ...message, sequence: 1 };
        }
        if (isHostMessage(message) && message.kind === 'reply' && !message.ok) {
          refusals.push(message.error);
        }
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    const peer = await hydratePeer(session);
    const version = await session.call.version();
    const saved = await session.save();
    const state = structuredClone(session.state);
    requests.length = 0;
    const edits = createWorkbookEditPeer({ session, peer });
    try {
      expect(edits.editCell(0, 0, 0, 'queued after attachment').applied).toBe(true);
      await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(await session.call.version()).toBe(version);
      expect(requests.slice(0, 2)).toEqual(['attachPeer', 'replay']);
      expect(refusals).toContainEqual(expect.objectContaining({
        name: 'WorkbookPeerHydrationError', refusal: { code: 'mutation-outside-replay' },
      }));
      expect(edits.error).toBeInstanceOf(WorkbookPeerHydrationError);
      expect(edits.error).toMatchObject({ code: 'version-mismatch' });
      expect(edits.state).toBe('failed');
      expect(edits.sentSequence).toBe(1);
      expect(edits.acknowledgedSequence).toBe(0);
      expect(peer.cell(0, 0, 0).input).toBe('queued after attachment');
      expect(await session.save()).toEqual(saved);
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(0);
      expect(session.state).toEqual(state);
      expect(session.failure).toBeUndefined();
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('refuses two queued retained replays after failed reattachment without worker mutation', async () => {
    const requests: string[] = [];
    const refusals: unknown[] = [];
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call') requests.push(message.method);
        listener(message);
      }),
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'reply' && !message.ok) {
          refusals.push(message.error);
        }
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    const peer = await hydratePeer(session);
    const edits = createWorkbookEditPeer({ session, peer });
    let reattached: WorkbookEditPeer | undefined;
    try {
      const version = await session.call.version();
      await edits.flush();
      edits.setActiveSheet(0);
      await edits.flush();
      expect(edits.acknowledgedSequence).toBe(1);
      expect(peer.version()).toBe(version);
      expect(await session.call.version()).toBe(version);
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(1);
      edits.dispose();
      const saved = await session.save();
      const cells = await session.call.cellInputs(0, 'A1:B1');
      const state = structuredClone(session.state);
      requests.length = 0;
      reattached = createWorkbookEditPeer({ session, peer });
      expect(reattached.editCell(0, 0, 0, 'first queued reattachment edit').applied).toBe(true);
      expect(reattached.editCell(0, 0, 1, 'second queued reattachment edit').applied).toBe(true);
      await expect(reattached.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(await session.call.version()).toBe(version);
      expect(requests.slice(0, 3)).toEqual(['attachPeer', 'replay', 'replay']);
      expect(refusals).toHaveLength(3);
      expect(refusals[0]).toMatchObject({
        name: 'WorkbookPeerHydrationError', refusal: { code: 'version-mismatch' },
      });
      for (const refusal of refusals.slice(1)) {
        expect(refusal).toMatchObject({
          name: 'WorkbookPeerHydrationError', refusal: { code: 'mutation-outside-replay' },
        });
      }
      expect(reattached.error).toBeInstanceOf(WorkbookPeerHydrationError);
      expect(reattached.error).toMatchObject({ code: 'version-mismatch' });
      expect(reattached.state).toBe('failed');
      expect(reattached.sentSequence).toBe(2);
      expect(reattached.acknowledgedSequence).toBe(0);
      expect(peer.cell(0, 0, 0).input).toBe('first queued reattachment edit');
      expect(peer.cell(0, 0, 1).input).toBe('second queued reattachment edit');
      expect(await session.save()).toEqual(saved);
      expect(await session.call.cellInputs(0, 'A1:B1')).toEqual(cells);
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(1);
      expect(session.state).toEqual(state);
      expect(session.failure).toBeUndefined();
    } finally {
      reattached?.dispose();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('clears retained attachment when a reattachment is refused while the edit peer stays live', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(fixture, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    }, pair.client);
    const peer = await hydratePeer(session);
    const edits = createWorkbookEditPeer({ session, peer });
    try {
      await edits.flush();
      const internal = workbookSessionInternals.get(session);
      if (!internal?.attachPeer) throw new Error('Missing internal workbook attachment helper');
      const version = await session.call.version();
      const saved = await session.save();
      await expect(internal.attachPeer('stale peer version')).rejects.toMatchObject({ code: 'version-mismatch' });
      const envelope: WorkbookReplayEnvelope = {
        sequence: edits.sentSequence + 1, calculation, op: { method: 'editCell', args: [0, 0, 0, 'refused replay'] },
      };
      await expect(internal.replay(envelope)).rejects.toMatchObject({ code: 'mutation-outside-replay' });
      expect(await session.call.version()).toBe(version);
      expect(await session.save()).toEqual(saved);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('keeps retained hydration opaque during open and consumes only wire version and sequence', async () => {
    const attachments: unknown[][] = [];
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'attachPeer') {
          attachments.push(message.args);
        }
        listener(message);
      }),
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'wasm-module') {
          message = { ...message, hydration: 'opaque invalid JSON', version: 'wire-version', sequence: 7 };
        }
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    try {
      expect(session.state.stage).toBe('ready');
      const internal = workbookSessionInternals.get(session);
      expect(internal?.initialVersion).toBe('wire-version');
      if (!internal?.attachPeer) throw new Error('Missing retained peer attachment');
      await expect(internal.attachPeer('wire-version')).rejects.toMatchObject({ code: 'version-mismatch' });
      expect(attachments).toEqual([['wire-version', 7]]);
      await expect(hydratePeer(session)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
      await expect(hydratePeer(session)).rejects.toMatchObject({ code: 'missing-hydration' });
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('fails retained open with a typed error when hydration wire version or sequence is missing', async () => {
    for (const missing of ['version', 'sequence'] as const) {
      const opening = createTestWorkbookSession(fixture, (transport) => ({
        ...transport,
        post(message, transfer) {
          if (isHostMessage(message) && message.kind === 'wasm-module') {
            message = { ...message, [missing]: undefined };
          }
          transport.post(message, transfer);
        },
      }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
      await expect(opening).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
      await expect(opening).rejects.toMatchObject({ code: 'missing-hydration' });
    }
  });

  test('accepts legacy module advertisements without hydration headers for ordinary sessions', async () => {
    const module = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'event' && message.name === 'peerOpened') {
          transport.post({ protocol: 1, kind: 'wasm-module', url: wasmAssetUrl().href, module });
        }
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer });
    try {
      expect(session.state.stage).toBe('ready');
      expect(await session.call.applyEdits({
        expectVersion: await session.call.version(),
        steps: [{
          op: 'setCellInputs',
          target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } },
          inputs: [['ordinary edit']],
        }],
      })).toMatchObject({ ok: true, applied: true });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('ordinary edit');
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('fails retained open with a typed error when the worker module has no hydration payload', async () => {
    const independentOpen = spyOn(XlsxDocument, 'open');
    const calculatedOpen = spyOn(XlsxDocument as typeof XlsxDocument & {
      openWithCalculationJson(bytes: Uint8Array, context: string): XlsxDocument;
    }, 'openWithCalculationJson');
    const opening = createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'wasm-module') {
          const { hydration, ...legacy } = message;
          transport.post(legacy, transfer);
          return;
        }
        if (isHostMessage(message) && message.kind === 'reply') return;
        transport.post(message, transfer);
      },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    try {
      await expect(opening).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
      await expect(opening).rejects.toMatchObject({ code: 'missing-hydration' });
      expect(independentOpen).not.toHaveBeenCalled();
      expect(calculatedOpen).toHaveBeenCalledTimes(1);
    } finally {
      independentOpen.mockRestore();
      calculatedOpen.mockRestore();
    }
  });

  test('initializes hydration with the exact module received from the worker', async () => {
    const pair = createInProcessPair();
    let module: WebAssembly.Module | undefined;
    const transport: SessionTransport = {
      ...pair.client,
      listen: (listener) => pair.client.listen((message) => {
        if (isHostMessage(message) && message.kind === 'wasm-module') module = message.module;
        listener(message);
      }),
    };
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(fixture, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    }, transport);
    const initializing = spyOn(workbookWasm, 'initWasm');
    let peer: WorkbookHandle | undefined;
    try {
      if (!module) throw new Error('Missing worker module');
      peer = await hydratePeer(session);
      expect(initializing).toHaveBeenCalledTimes(1);
      expect(initializing).toHaveBeenCalledWith(module);
      expect(peer.version()).toBe(await session.call.version());
      expect(peer.save()).toEqual(await session.save());
    } finally {
      initializing.mockRestore();
      peer?.dispose();
      await session.dispose();
    }
  });

  test.each([
    ['east of UTC', Date.UTC(2026, 0, 1, 23, 30), -120, Date.UTC(2026, 0, 2, 1, 30)],
    ['west of UTC', Date.UTC(2026, 0, 2, 2), 300, Date.UTC(2026, 0, 1, 21)],
    ['half-hour offset', Date.UTC(2026, 0, 1, 20), -330, Date.UTC(2026, 0, 2, 1, 30)],
  ] as const)('shares local NOW and TODAY across retained open and edit replay (%s)', async (_, ms, offset, localMs) => {
    const clock = spyOn(Date, 'now').mockReturnValue(ms);
    const timezone = spyOn(Date.prototype, 'getTimezoneOffset').mockReturnValue(offset);
    let session: WorkbookSession | undefined;
    let peer: WorkbookHandle | undefined;
    let edits: WorkbookEditPeer | undefined;
    const nowSerial = localMs / 86_400_000 + 25_569;
    const read = { ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:C1' } }] } as const;
    try {
      session = await createTestWorkbookSession(await volatileWorkbookBytes(), undefined, {
        wasm: wasmBytes.buffer, retainPeerHydration: true,
      });
      const opening = await session.call.readCells(read);
      if (!opening.ok) throw new Error(opening.failure.message);
      expect(opening.ranges[0].cells[0].slice(0, 2).map((cell: XlsxCellRead) => cell.value)).toEqual([
        { kind: 'number', value: nowSerial },
        { kind: 'number', value: Math.floor(nowSerial) },
      ]);
      peer = await hydratePeer(session);
      expect(peer.readCells(read)).toEqual(opening);
      edits = createWorkbookEditPeer({ session, peer, randomSeed: () => 42 });
      clock.mockReturnValue(ms + 86_400_000);
      expect(edits.editCell(0, 1, 0, '1').applied).toBe(true);
      await edits.flush();
      const edited = peer.readCells(read);
      if (!edited.ok) throw new Error(edited.failure.message);
      expect(edited.ranges[0].cells[0].slice(0, 2).map((cell: XlsxCellRead) => cell.value)).toEqual([
        { kind: 'number', value: nowSerial + 1 },
        { kind: 'number', value: Math.floor(nowSerial) + 1 },
      ]);
      const workerEdited = await session.call.readCells(read);
      if (!workerEdited.ok) throw new Error(workerEdited.failure.message);
      expect(edited).toEqual(workerEdited);
      clock.mockReturnValue(ms + 2 * 86_400_000);
      expect(new Uint8Array(await edits.save())).toEqual(new Uint8Array(peer.save()));
      expect(peer.readCells(read)).toEqual(edited);
      expect(await session.call.readCells(read)).toEqual(edited);
      expect(session.failure).toBeUndefined();
    } finally {
      try {
        edits?.dispose();
        peer?.dispose();
        await session?.dispose();
      } finally {
        timezone.mockRestore();
        clock.mockRestore();
      }
    }
  });

  test('adopts worker volatile values and version without recalculating during hydration', async () => {
    const bytes = await volatileWorkbookBytes();
    const session = await createTestWorkbookSession(bytes, undefined, {
      wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    const read = { ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:C1' } }] } as const;
    const opening = await session.call.readCells(read);
    if (!opening.ok) throw new Error(opening.failure.message);
    expect(opening.ranges[0]?.cells[0]?.map((cell) => cell.value.kind)).toEqual(['number', 'number', 'number']);
    const saved = await session.save();
    const clock = spyOn(Date, 'now').mockReturnValue(0);
    const ordinaryOpen = spyOn(XlsxDocument, 'open');
    const calculatedOpen = spyOn(XlsxDocument as typeof XlsxDocument & {
      openWithCalculationJson(bytes: Uint8Array, context: string): XlsxDocument;
    }, 'openWithCalculationJson');
    let peer: WorkbookHandle | undefined;
    try {
      peer = await hydratePeer(session);
      expect(ordinaryOpen).not.toHaveBeenCalled();
      expect(calculatedOpen).not.toHaveBeenCalled();
    } finally {
      clock.mockRestore();
      ordinaryOpen.mockRestore();
      calculatedOpen.mockRestore();
    }
    if (!peer) throw new Error('Missing hydrated workbook peer');
    expect(peer.version()).toBe(await session.call.version());
    const edits = createWorkbookEditPeer({
      session, peer, now: () => (46_001.75 - 25_569) * 86_400_000, randomSeed: () => 42,
    });
    try {
      expect(peer.readCells(read)).toEqual(opening);
      expect(peer.version()).toBe(await session.call.version());
      expect(peer.save()).toEqual(saved);
      const cells = peer.readCells(read);
      if (!cells.ok) throw new Error(cells.failure.message);
      const guarded = cells.ranges[0]?.cells[0]?.[2];
      if (!guarded) throw new Error('Missing RAND cell');
      expect(edits.applyEdits({
        expectVersion: peer.version(),
        steps: [{
          op: 'setCellInputs', target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'C1' } },
          expect: { cells: [[{ value: guarded.value }]] }, inputs: [['=RAND()+1']],
        }],
      })).toMatchObject({ ok: true, applied: true });
      await edits.flush();
      expect(peer.readCells(read)).toEqual(await session.call.readCells(read));
      expect(peer.version()).toBe(await session.call.version());
      expect(peer.save()).toEqual(await session.save());
      expect(session.failure).toBeUndefined();
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('keeps hydration opt-in and normal sessions usable without retained state', async () => {
    const session = await createTestWorkbookSession(fixture, undefined, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    try {
      const version = await session.call.version();
      await expect(hydratePeer(session)).rejects.toThrow('does not retain peer hydration state');
      expect(await session.call.version()).toBe(version);
      expect(await session.call.readCells({ ranges: [] })).toEqual({
        ...peer.readCells({ ranges: [] }), version,
      });
      expect(await session.save()).toEqual(peer.save());
      expect(session.state.stage).toBe('ready');
      expect(session.failure).toBeUndefined();
    } finally {
      peer.dispose();
      await session.dispose();
    }
  });

  test('refuses retained hydration when the worker supplies no compiled module', async () => {
    await expect(createTestWorkbookSession(fixture, undefined, {
      calculation, retainPeerHydration: true,
    })).rejects.toThrow('did not retain a compiled module');
  });

  test('hydrates retained bytes after pre-peer worker failure and saves queued recovery edits once', async () => {
    const host = crashableHost();
    const bytes = fixture.slice();
    const openingCalculation = { ...calculation };
    const session = await createTestWorkbookSession(bytes, host.wrap, {
      calculation: openingCalculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    const errors: Error[] = [];
    let peer: WorkbookHandle | undefined;
    const original = openWorkbook(fixture, { calculation });
    try {
      bytes.fill(0);
      openingCalculation.nowSerial = 0;
      const waiting = Promise.allSettled([session.call.readCells({ ranges: [] }), session.save()]);
      host.crash();
      for (const result of await waiting) {
        if (result.status !== 'rejected') throw new Error('Normal request survived worker failure');
        expect(result.reason).toBeInstanceOf(SessionFailure);
      }
      expect(session.state.stage).toBe('failed');
      const compiling = spyOn(WebAssembly, 'compile');
      const streaming = spyOn(WebAssembly, 'compileStreaming');
      try {
        const hydration = hydratePeer(session);
        expect(hydratePeer(session)).toBe(hydration);
        peer = await hydration;
        expect(compiling).not.toHaveBeenCalled();
        expect(streaming).not.toHaveBeenCalled();
      } finally {
        compiling.mockRestore();
        streaming.mockRestore();
      }
      if (!peer) throw new Error('Missing hydrated workbook peer');
      expect(peer.save()).toEqual(original.save());
      const edits = createWorkbookEditPeer({
        session, peer, now: () => 0, randomSeed: () => 1, onError: (error) => { errors.push(error); },
      });
      const input: WorkbookReplayOp = { method: 'editCell', args: [0, 2, 1, 'queued before hydration'] };
      const rows: WorkbookReplayOp = {
        method: 'applyOps', args: [[{ type: 'insertRows', sheet: 0, at: 2, count: 1 }]],
      };
      try {
        expect(edits.state).toBe('failed');
        expect(() => edits.editCell(0, 2, 1, 'blocked')).toThrow(WorkbookEditPeerFailedError);
        await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
        await expect(edits.save()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
        expect(workbookEditPeerOperations(edits).applyRecoveryOp(input)).toMatchObject({ applied: true });
        expect(workbookEditPeerOperations(edits).applyRecoveryOp(rows)).toMatchObject({ applied: true });
        const version = peer.version();
        const expected = peer.save();
        expect(workbookEditPeerOperations(edits).applyRecoveryOp(input)).toMatchObject({ applied: true });
        expect(workbookEditPeerOperations(edits).applyRecoveryOp(rows)).toMatchObject({ applied: true });
        expect(peer.version()).toBe(version);
        expect(peer.save()).toEqual(expected);
        expect(peer.cell(0, 3, 1).input).toBe('queued before hydration');
        expect(edits.sentSequence).toBe(0);
        expect(edits.acknowledgedSequence).toBe(0);
        const recovery = edits.recoverySave();
        expect(recovery.recovery).toBe(true);
        expect<Uint8Array>(new Uint8Array(recovery.bytes)).toEqual(expected);
        expect(errors).toEqual([session.failure!]);
        const reopened = openWorkbook(new Uint8Array(recovery.bytes), { calculation });
        try {
          expect(reopened.cell(0, 3, 1).input).toBe('queued before hydration');
        } finally { reopened.dispose(); }
      } finally { edits.dispose(); }
    } finally {
      peer?.dispose();
      original.dispose();
      await session.dispose();
    }
  });

  test('finishes in-flight hydration when the worker fails', async () => {
    const host = crashableHost();
    const session = await createTestWorkbookSession(fixture, host.wrap, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    let peer: WorkbookHandle | undefined;
    try {
      const hydration = hydratePeer(session);
      const waiting = session.call.version();
      host.crash();
      await expect(waiting).rejects.toBeInstanceOf(SessionFailure);
      peer = await hydration;
      expect(await hydratePeer(session)).toBe(peer);
      const errors: Error[] = [];
      const edits = createWorkbookEditPeer({ session, peer, onError: (error) => { errors.push(error); } });
      try {
        expect<Uint8Array>(new Uint8Array(edits.recoverySave().bytes)).toEqual(peer.save());
        expect(errors).toEqual([session.failure!]);
      } finally { edits.dispose(); }
    } finally {
      peer?.dispose();
      await session.dispose();
    }
  });

  test('rejects hydration after disposal and cancels an in-flight hydration', async () => {
    const session = await createTestWorkbookSession(fixture, undefined, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    const hydration = hydratePeer(session);
    const disposal = session.dispose();
    await expect(hydration).rejects.toMatchObject({ code: 'disposed' });
    await disposal;
    await expect(hydratePeer(session)).rejects.toMatchObject({ code: 'disposed' });
  });
});
