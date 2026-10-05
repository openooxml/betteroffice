import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { isClientMessage, isHostMessage, type SessionTransport } from '../../../../shared/office-session';
import * as workbookWasm from '../wasm/loader';
import type { WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import { hydratePeer, openWorkbookSession, type WorkbookSession } from './client';
import { createWorkbookEditPeer, type WorkbookEditPeer } from './editPeer';
import { createTestWorkbookSession } from './testHelpers';

let bytes: Uint8Array;
let wasm: Uint8Array<ArrayBuffer>;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };
const editOptions = {
  now: () => (calculation.nowSerial - 25_569) * 86_400_000,
  randomSeed: () => calculation.randSeed,
};

beforeAll(async () => {
  wasm = new Uint8Array(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  await workbookWasm.initWasm(wasm);
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
  const rows = Array.from({ length: 800 }, (_, index) => {
    const row = index + 1;
    return `<row r="${row}"><c r="A${row}"><v>${row}</v></c><c r="B${row}"><f>A${row}*2</f><v>0</v></c></row>`;
  }).join('');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}<row r="801"><c r="A801"><f>NOW()</f><v>0</v></c><c r="B801"><f>RAND()</f><v>0</v></c></row></sheetData></worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>');
  bytes = await zip.generateAsync({ type: 'uint8array' });
});

function retainedSession(wrap?: (transport: SessionTransport) => SessionTransport) {
  return createTestWorkbookSession(bytes, wrap, {
    calculation, wasm: wasm.buffer, retainPeerHydration: true,
  });
}

async function equalPeer(peer: WorkbookHandle, session: WorkbookSession): Promise<void> {
  expect(peer.version()).toBe(await session.call.version());
  expect(peer.save()).toEqual(await session.save());
  expect(session.failure).toBeUndefined();
}

async function replayEdits(peer: WorkbookHandle, session: WorkbookSession): Promise<void> {
  const edits = createWorkbookEditPeer({ session, peer, ...editOptions });
  try {
    await edits.flush();
    expect(edits.state).toBe('ready');
    expect(edits.editCell(0, 1, 0, '91').applied).toBe(true);
    await edits.flush();
    await equalPeer(peer, session);
    edits.setActiveSheet(1);
    await edits.flush();
    expect(session.state.activeSheet).toBe(1);
    expect(edits.applyOps([{ type: 'insertRows', sheet: 1, at: 0, count: 1 }]).applied).toBe(true);
    await edits.flush();
    await equalPeer(peer, session);
    expect(edits.acknowledgedSequence).toBe(edits.sentSequence);
  } finally { edits.dispose(); }
}

test('dedicated XLSX worker hydrates the edit peer through snapshot RPCs', async () => {
  const worker = new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' });
  const posting = spyOn(worker, 'postMessage');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let session: WorkbookSession | undefined;
  let peer: WorkbookHandle | undefined;
  try {
    session = await openWorkbookSession(bytes, {
      calculation, wasm: wasm.buffer, worker: () => worker, retainPeerHydration: true,
    });
    peer = await hydratePeer(session);
    const calls = posting.mock.calls.flatMap(([message]) =>
      isClientMessage(message) && message.kind === 'call' ? [message] : []
    );
    expect(calls.find((message) => message.method === 'beginPeerSnapshot')?.args).toEqual([256, 16 * 1024]);
    expect(calls.filter((message) => message.method === 'pullPeerSnapshot').length).toBeGreaterThan(2);
    expect(calls.find((message) => message.method === 'endPeerSnapshot')?.args).toEqual([false]);
    expect(sourceOpening).not.toHaveBeenCalled();
    expect(warning).not.toHaveBeenCalled();
    await equalPeer(peer, session);
    await replayEdits(peer, session);
  } finally {
    posting.mockRestore();
    warning.mockRestore();
    sourceOpening.mockRestore();
    peer?.dispose();
    await session?.dispose();
    worker.terminate();
  }
});

test('snapshot-hydrated peer equals source hydration and replays cell and sheet edits identically', async () => {
  const open = workbookWasm.openWorkbook;
  let worker: WorkbookHandle | undefined;
  const opening = spyOn(workbookWasm, 'openWorkbook').mockImplementation((...args) => {
    worker = open(...args);
    return worker;
  });
  let session: WorkbookSession;
  try { session = await retainedSession(); } finally { opening.mockRestore(); }
  if (!worker) throw new Error('Missing worker workbook');
  const source = workbookWasm.openWorkbookPeer(bytes, { calculation }, workbookWasm.workbookPeerHydration(worker));
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  let peer: WorkbookHandle | undefined;
  let edits: WorkbookEditPeer | undefined;
  try {
    peer = await hydratePeer(session);
    expect(sourceOpening).not.toHaveBeenCalled();
    expect(warning).not.toHaveBeenCalled();
    expect(peer.version()).toBe(source.version());
    expect(peer.save()).toEqual(source.save());
    expect(peer.cell(0, 800, 0)).toEqual(source.cell(0, 800, 0));
    expect(peer.cell(0, 800, 1)).toEqual(source.cell(0, 800, 1));
    edits = createWorkbookEditPeer({ session, peer, ...editOptions });
    await edits.flush();
    source.setCalculationContext(calculation);
    expect(edits.editCell(0, 1, 0, '91')).toEqual(source.editCell(0, 1, 0, '91'));
    await edits.flush();
    expect(peer.save()).toEqual(source.save());
    await equalPeer(peer, session);
    source.setCalculationContext(calculation);
    expect(edits.applyOps([{ type: 'insertRows', sheet: 1, at: 0, count: 1 }]))
      .toEqual(source.applyOps([{ type: 'insertRows', sheet: 1, at: 0, count: 1 }]));
    await edits.flush();
    expect(peer.save()).toEqual(source.save());
    await equalPeer(peer, session);
  } finally {
    sourceOpening.mockRestore();
    warning.mockRestore();
    edits?.dispose();
    peer?.dispose();
    source.dispose();
    await session.dispose();
  }
});

test('xlsx worker editor falls back to source hydration when the snapshot is refused', async () => {
  let pullId: number | undefined;
  let refused = false;
  const discarded: boolean[] = [];
  const session = await retainedSession((transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call') {
        if (message.method === 'pullPeerSnapshot' && !refused) pullId = message.id;
        if (message.method === 'endPeerSnapshot') discarded.push(message.args[0] as boolean);
      }
      listener(message);
    }),
    post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && message.id === pullId && !refused) {
        refused = true;
        message = { protocol: 1, kind: 'reply', id: message.id, ok: false,
          error: { name: 'Error', message: 'Injected snapshot refusal' } };
      }
      transport.post(message, transfer);
    },
  }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(refused).toBe(true);
    expect(discarded).toEqual([true]);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: Injected snapshot refusal');
    expect(sourceOpening).toHaveBeenCalledTimes(1);
    await equalPeer(peer, session);
    await replayEdits(peer, session);
    expect(warning).toHaveBeenCalledTimes(1);
  } finally {
    warning.mockRestore();
    sourceOpening.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});

test('snapshot hydration transfers multiple chunks and yields between fixed-budget advances', async () => {
  const buffers: ArrayBuffer[] = [];
  const transfers: boolean[] = [];
  const begins: unknown[][] = [];
  const requests = new Map<number, string>();
  const session = await retainedSession((transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call') {
        requests.set(message.id, message.method);
        if (message.method === 'beginPeerSnapshot') begins.push(message.args);
      }
      listener(message);
    }),
    post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
        requests.get(message.id) === 'pullPeerSnapshot' && message.value instanceof ArrayBuffer) {
        buffers.push(message.value);
        transfers.push(transfer?.length === 1 && transfer[0] === message.value);
      }
      transport.post(message, transfer);
    },
  }));
  const createBuilder = workbookWasm.createWorkbookSnapshotBuilder;
  const advances: number[][] = [];
  const turns: boolean[] = [];
  let yielded = true;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const building = spyOn(workbookWasm, 'createWorkbookSnapshotBuilder').mockImplementation((options) => {
    const builder = createBuilder(options);
    return { ...builder, advance(records, bytes) {
      turns.push(yielded);
      yielded = false;
      timer = setTimeout(() => { yielded = true; }, 0);
      advances.push([records, bytes]);
      return builder.advance(records, bytes);
    } };
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(warning).not.toHaveBeenCalled();
    expect(begins).toEqual([[256, 16 * 1024]]);
    expect(buffers.length).toBeGreaterThan(1);
    expect(transfers.every(Boolean)).toBe(true);
    expect(buffers.every((buffer) => buffer.byteLength === 0)).toBe(true);
    expect(advances.length).toBeGreaterThan(1);
    expect(advances.every(([records, bytes]) => records === 256 && bytes === 16 * 1024)).toBe(true);
    expect(turns).toEqual(advances.map(() => true));
    await equalPeer(peer, session);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    building.mockRestore();
    warning.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['push', 'advance', 'finish'] as const)('snapshot %s errors fall back and keep edit replay ready', async (stage) => {
  const session = await retainedSession();
  const createBuilder = workbookWasm.createWorkbookSnapshotBuilder;
  let disposed = false;
  const building = spyOn(workbookWasm, 'createWorkbookSnapshotBuilder').mockImplementation((options) => {
    const builder = createBuilder(options);
    return {
      push(chunk) {
        if (stage === 'push') throw new Error('Injected snapshot push failure');
        builder.push(chunk);
      },
      advance(records, bytes) {
        if (stage === 'advance') throw new Error('Injected snapshot advance failure');
        return builder.advance(records, bytes);
      },
      finish() {
        if (stage === 'finish') throw new Error('Injected snapshot finish failure');
        return builder.finish();
      },
      dispose() {
        disposed = true;
        builder.dispose();
      },
    };
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(disposed).toBe(true);
    expect(warning).toHaveBeenCalledWith(`xlsx worker editor: snapshot hydration fell back: Injected snapshot ${stage} failure`);
    expect(sourceOpening).toHaveBeenCalledTimes(1);
    await equalPeer(peer, session);
    await replayEdits(peer, session);
  } finally {
    building.mockRestore();
    warning.mockRestore();
    sourceOpening.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['beginPeerSnapshot', 'endPeerSnapshot'] as const)('snapshot %s refusals fall back without failing the session', async (method) => {
  let requestId: number | undefined;
  let refused = false;
  const session = await retainedSession((transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === method && !refused) {
        requestId = message.id;
      }
      listener(message);
    }),
    post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && message.id === requestId && !refused) {
        refused = true;
        message = { protocol: 1, kind: 'reply', id: message.id, ok: false,
          error: { name: 'Error', message: 'Workbook changed during peer snapshot' } };
      }
      transport.post(message, transfer);
    },
  }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(refused).toBe(true);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: Workbook changed during peer snapshot');
    expect(sourceOpening).toHaveBeenCalledTimes(1);
    await replayEdits(peer, session);
  } finally {
    warning.mockRestore();
    sourceOpening.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['version', 'sequence'] as const)('snapshot rejects a changed %s pin and hydrates from the retained source', async (field) => {
  let beginId: number | undefined;
  const session = await retainedSession((transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'beginPeerSnapshot') {
        beginId = message.id;
      }
      listener(message);
    }),
    post(message, transfer) {
      if (isHostMessage(message) && message.kind === 'reply' && message.ok && message.id === beginId) {
        const pinned = message.value as { version: string; sequence: number };
        message = { ...message, value: { ...pinned, [field]: field === 'version' ? 'changed' : pinned.sequence + 1 } };
      }
      transport.post(message, transfer);
    },
  }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: Workbook snapshot differs from retained peer hydration');
    await equalPeer(peer, session);
    await replayEdits(peer, session);
  } finally {
    warning.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['begin', 'next', 'end'] as const)('snapshot %s traps remain recoverable RPC errors', async (stage) => {
  const session = await retainedSession();
  const snapshotAccess = workbookWasm.workbookPeerSnapshot;
  const snapshots = spyOn(workbookWasm, 'workbookPeerSnapshot').mockImplementation((handle) => {
    const snapshot = snapshotAccess(handle);
    return { ...snapshot, [stage]: () => { throw new WebAssembly.RuntimeError('Injected snapshot trap'); } };
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  let peer: WorkbookHandle | undefined;
  try {
    peer = await hydratePeer(session);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: Injected snapshot trap');
    await equalPeer(peer, session);
    await replayEdits(peer, session);
  } finally {
    snapshots.mockRestore();
    warning.mockRestore();
    peer?.dispose();
    await session.dispose();
  }
});
