import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { isClientMessage, isHostMessage, type SessionTransport } from '../../../../shared/office-session';
import * as workbookWasm from '../wasm/loader';
import type { WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { XlsxReadRequest } from '../edits';
import { hydratePeer, openWorkbookSession, type WorkbookSession } from './client';
import { workbookPeerSources } from './clientInternals';
import { createWorkbookEditPeer, type WorkbookEditPeer } from './editPeer';
import { workbookEditPeerOperations } from './editPeerInternals';
import { applyWorkbookReplayOp, workbookSessionInternals, type WorkbookReplayOp } from './replay';
import { createTestWorkbookSession } from './testHelpers';

let bytes: Uint8Array<ArrayBuffer>;
let wasm: Uint8Array<ArrayBuffer>;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };
const editMs = (calculation.nowSerial - 25_569) * 86_400_000;
const editOptions = {
  now: () => editMs,
  randomSeed: () => calculation.randSeed,
};
const editedCalculation: WorkbookCalculationContext = { ...calculation,
  nowSerial: (editMs - new Date(editMs).getTimezoneOffset() * 60_000) / 86_400_000 + 25_569 };

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
  bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
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
  let restoreReplay: (() => void) | undefined;
  try {
    peer = await hydratePeer(session);
    expect(sourceOpening).not.toHaveBeenCalled();
    expect(warning).not.toHaveBeenCalled();
    expect(peer.version()).toBe(source.version());
    expect(peer.save()).toEqual(source.save());
    expect(peer.cell(0, 800, 0)).toEqual(source.cell(0, 800, 0));
    expect(peer.cell(0, 800, 1)).toEqual(source.cell(0, 800, 1));
    const internal = workbookSessionInternals.get(session)!;
    const replaying = spyOn(internal, 'replay');
    restoreReplay = () => replaying.mockRestore();
    edits = createWorkbookEditPeer({ session, peer, ...editOptions });
    await edits.flush();
    source.setCalculationContext(editedCalculation);
    expect(edits.editCell(0, 1, 0, '91')).toEqual(source.editCell(0, 1, 0, '91'));
    await edits.flush();
    expect(peer.save()).toEqual(source.save());
    await equalPeer(peer, session);
    const initialReplays = replaying.mock.calls.length;
    const operations = workbookEditPeerOperations(edits);
    const readRequest: XlsxReadRequest = { ranges: [
      { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:B24' } },
      { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A801:B801' } },
    ] };
    const check = async () => {
      await edits!.flush();
      await equalPeer(peer!, session);
      expect(peer!.version()).toBe(source.version());
      expect(peer!.save()).toEqual(source.save());
      const read = peer!.readCells(readRequest);
      expect(read).toMatchObject({ ok: true });
      expect(await session.call.readCells(readRequest)).toEqual(read);
      expect(source.readCells(readRequest)).toEqual(read);
      expect(await session.call.calculationStatus()).toEqual(peer!.calculationStatus());
      expect(edits!.acknowledgedSequence).toBe(edits!.sentSequence);
    };
    let precedingCalculation = editedCalculation;
    const apply = async (op: WorkbookReplayOp, refused = false) => {
      if (!op.calculation) throw new Error('Missing operation calculation context');
      const before = edits!.sentSequence;
      const calls = replaying.mock.calls.length;
      source.setCalculationContext(op.calculation!);
      const expected = applyWorkbookReplayOp(source, op);
      const result = operations.applyQueuedOp(op);
      expect(result).toEqual(expected);
      expect(operations.applyQueuedOp(op)).toBe(result);
      if (refused) {
        source.setCalculationContext(precedingCalculation);
        expect(result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
        expect(edits!.sentSequence).toBe(before);
        expect(replaying).toHaveBeenCalledTimes(calls);
      } else {
        precedingCalculation = op.calculation!;
        expect(edits!.sentSequence).toBe(before + 1);
        expect(replaying).toHaveBeenCalledTimes(calls + 1);
        expect(replaying.mock.calls[calls][0]).toEqual({ sequence: before + 1, op, calculation: op.calculation });
      }
      await check();
      return result;
    };
    for (let index = 0; index < 24; index++) {
      const context = { nowSerial: calculation.nowSerial + index / 24, randSeed: calculation.randSeed + index };
      const op: WorkbookReplayOp = { method: 'editCell',
        args: index === 0 ? [0, 800, 0, '=NOW()+1'] : index === 1 ? [0, 800, 1, '=RAND()+1'] :
          [0, index - 2, 0, String(100 + index)], calculation: context };
      await apply(op);
    }
    const beforeNoop = edits.sentSequence;
    const noop: WorkbookReplayOp = { method: 'applyEdits', args: [{ expectVersion: peer.version(), steps: [{
      op: 'setCellInputs', target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } },
      inputs: [[peer.cell(0, 0, 0).input]],
    }] }],
      calculation: { nowSerial: 47_000, randSeed: 42 } };
    expect(await apply(noop)).toMatchObject({ ok: true, applied: false });
    expect(edits.sentSequence).toBe(beforeNoop + 1);
    expect(replaying).toHaveBeenCalledTimes(initialReplays + 25);
    await apply({ method: 'applyEdits', args: [{ expectVersion: 'stale', steps: [{
      op: 'setCellInputs', target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } },
      inputs: [['refused']],
    }] }], calculation: { nowSerial: 48_000, randSeed: 43 } }, true);
    await apply({ method: 'undo', args: [], calculation: { nowSerial: 48_001, randSeed: 44 } });
    await apply({ method: 'redo', args: [], calculation: { nowSerial: 48_002, randSeed: 45 } });
    expect(replaying).toHaveBeenCalledTimes(initialReplays + 27);
    expect(sourceOpening).not.toHaveBeenCalled();
    source.setCalculationContext(editedCalculation);
    expect(edits.applyOps([{ type: 'insertRows', sheet: 1, at: 0, count: 1 }]))
      .toEqual(source.applyOps([{ type: 'insertRows', sheet: 1, at: 0, count: 1 }]));
    await edits.flush();
    expect(peer.save()).toEqual(source.save());
    await equalPeer(peer, session);
  } finally {
    restoreReplay?.();
    sourceOpening.mockRestore();
    warning.mockRestore();
    edits?.dispose();
    peer?.dispose();
    source.dispose();
    await session.dispose();
  }
});

test.each(['snapshot', 'fallback'] as const)('attachment releases worker preview data after %s hydration', async (route) => {
  const session = await retainedSession();
  const internal = workbookSessionInternals.get(session)!;
  const source = workbookPeerSources.get(session)!;
  const opening = spyOn(workbookWasm, 'openWorkbookPeer');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const snapshot = route === 'fallback' ? spyOn(source.snapshot!, 'beginPeerSnapshot').mockImplementation(() => {
    throw new Error('Injected snapshot refusal');
  }) : undefined;
  let peer: WorkbookHandle | undefined;
  let edits: WorkbookEditPeer | undefined;
  const viewport = { x: 0, y: 0, width: 800, height: 600 };
  try {
    expect(await internal.preview!(viewport, 0, [{ method: 'editCell', args: [0, 0, 0, 'cold'], calculation }]))
      .toMatchObject({ sequence: 0, sheet: 0 });
    expect(opening).toHaveBeenCalledTimes(1);
    peer = await hydratePeer(session);
    expect(opening).toHaveBeenCalledTimes(route === 'fallback' ? 2 : 1);
    edits = createWorkbookEditPeer({ session, peer, ...editOptions });
    await edits.flush();
    await expect(internal.preview!(viewport, 0, [])).rejects.toMatchObject({
      name: 'Error', message: 'Worker preview requires retained hydration',
    });
    expect(opening).toHaveBeenCalledTimes(route === 'fallback' ? 2 : 1);
    await equalPeer(peer, session);
    expect(edits.editCell(0, 0, 0, 'ready').applied).toBe(true);
    await edits.flush();
    await equalPeer(peer, session);
    await internal.detachPeer!();
    await expect(internal.preview!(viewport, 0, [])).rejects.toMatchObject({
      name: 'Error', message: 'Worker preview requires retained hydration',
    });
    expect(opening).toHaveBeenCalledTimes(route === 'fallback' ? 2 : 1);
    expect(session.failure).toBeUndefined();
  } finally {
    snapshot?.mockRestore();
    opening.mockRestore();
    warning.mockRestore();
    edits?.dispose();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['explicit', 'worker-default'] as const)('successful snapshot releases fallback data and attaches with the exact %s calculation context', async (context) => {
  const session = await createTestWorkbookSession(bytes, undefined, {
    calculation: context === 'explicit' ? calculation : undefined,
    wasm: wasm.buffer, retainPeerHydration: true,
  });
  const source = workbookPeerSources.get(session);
  const internal = workbookSessionInternals.get(session);
  if (!source?.hydration || !internal) throw new Error('Missing retained peer source');
  expect(source.bytes?.byteLength).toBe(bytes.byteLength);
  const fallbackHydration = source.hydration;
  const initial = JSON.parse(fallbackHydration);
  const expected = { nowSerial: initial.calculation_context.now_serial, randSeed: initial.workbook.rand_seed };
  if (context === 'explicit') expect(expected).toEqual(calculation);
  expect(internal.initialCalculation).toEqual(expected);
  const parsing = spyOn(JSON, 'parse');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  let edits: WorkbookEditPeer | undefined;
  try {
    const hydration = hydratePeer(session);
    expect(hydratePeer(session)).toBe(hydration);
    peer = await hydration;
    expect(source.bytes).toBeUndefined();
    expect(source.hydration).toBeUndefined();
    expect(source.pending).toBe(hydration);
    expect(hydratePeer(session)).toBe(hydration);
    expect(await hydratePeer(session)).toBe(peer);
    expect(internal.initialCalculation).toEqual(expected);
    edits = createWorkbookEditPeer({ session, peer, ...editOptions });
    await edits.flush();
    expect(edits.state).toBe('ready');
    expect(parsing).not.toHaveBeenCalledWith(fallbackHydration);
    expect(sourceOpening).not.toHaveBeenCalled();
    expect(warning).not.toHaveBeenCalled();
    const setting = spyOn(peer, 'setCalculationContext');
    try {
      expect(edits.applyEdits({ expectVersion: 'stale', steps: [{
        op: 'setCellInputs', target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } },
        inputs: [['refused']],
      }] })).toMatchObject({ ok: false });
      await edits.flush();
      expect(setting.mock.calls).toEqual([[editedCalculation], [expected]]);
      expect(edits.sentSequence).toBe(0);
      await equalPeer(peer, session);
      expect(edits.editCell(0, 1, 0, '91').applied).toBe(true);
      await edits.flush();
      await equalPeer(peer, session);
      expect(new Uint8Array(await edits.save())).toEqual(new Uint8Array(await session.save()));
    } finally { setting.mockRestore(); }
  } finally {
    parsing.mockRestore();
    warning.mockRestore();
    sourceOpening.mockRestore();
    edits?.dispose();
    peer?.dispose();
    await session.dispose();
  }
});

test.each(['null', 'missing'] as const)('source-fallback attachment preserves %s initial calculation context', async (context) => {
  const session = await createTestWorkbookSession(bytes, (transport) => ({
    ...transport,
    post(message, transfer) {
      if (context === 'missing' && isHostMessage(message) && message.kind === 'event' && message.name === 'peerOpened') {
        const { initialCalculation, ...payload } = message.payload as { version: string; initialCalculation?: unknown };
        message = { ...message, payload };
      }
      transport.post(message, transfer);
    },
  }), { collaborative: true, clientId: 41, wasm: wasm.buffer, retainPeerHydration: true });
  const internal = workbookSessionInternals.get(session);
  if (!internal) throw new Error('Missing workbook session internals');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  let edits: WorkbookEditPeer | undefined;
  try {
    expect(internal.initialCalculation).toBe(context === 'null' ? null : undefined);
    peer = await hydratePeer(session);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: collaborative workbooks cannot be snapshotted');
    expect(sourceOpening).toHaveBeenCalledTimes(1);
    expect(workbookPeerSources.get(session)?.bytes).toEqual(bytes);
    expect(workbookPeerSources.get(session)?.hydration).toBeDefined();
    expect(internal.initialCalculation).toBe(context === 'null' ? null : undefined);
    edits = createWorkbookEditPeer({ session, peer });
    await edits.flush();
    expect(edits.state).toBe('ready');
    await equalPeer(peer, session);
  } finally {
    warning.mockRestore();
    sourceOpening.mockRestore();
    edits?.dispose();
    peer?.dispose();
    await session.dispose();
  }
});

test('real oversized-cell snapshot refusal discards the worker snapshot and replays source-hydrated edits', async () => {
  const zip = await JSZip.loadAsync(bytes);
  const text = 'x'.repeat(20_000);
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${text}</t></is></c></row></sheetData></worksheet>`);
  const oversized = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
  const open = workbookWasm.openWorkbook;
  let worker: WorkbookHandle | undefined;
  const opening = spyOn(workbookWasm, 'openWorkbook').mockImplementation((...args) => {
    worker = open(...args);
    return worker;
  });
  const discarded: boolean[] = [];
  let session: WorkbookSession;
  try {
    session = await createTestWorkbookSession(oversized, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'endPeerSnapshot') {
          discarded.push(message.args[0] as boolean);
        }
        listener(message);
      }),
    }), { calculation, wasm: wasm.buffer, retainPeerHydration: true });
  } finally { opening.mockRestore(); }
  if (!worker) throw new Error('Missing worker workbook');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const sourceOpening = spyOn(workbookWasm, 'openWorkbookPeer');
  let peer: WorkbookHandle | undefined;
  try {
    expect(worker.cell(1, 0, 0).input).toBe(text);
    peer = await hydratePeer(session);
    expect(warning).toHaveBeenCalledWith('xlsx worker editor: snapshot hydration fell back: Yrs snapshot record exceeds advance byte budget');
    expect(warning).toHaveBeenCalledTimes(1);
    expect(discarded).toEqual([true]);
    expect(() => workbookWasm.workbookPeerSnapshot(worker!).next()).toThrow('workbook peer snapshot is not active');
    expect(session.failure).toBeUndefined();
    expect(session.state.stage).toBe('ready');
    expect(sourceOpening).toHaveBeenCalledTimes(1);
    expect(sourceOpening.mock.calls[0]?.[0]).toEqual(oversized);
    expect(peer.cell(1, 0, 0).input).toBe(text);
    expect(workbookPeerSources.get(session)?.bytes).toEqual(oversized);
    expect(workbookPeerSources.get(session)?.hydration).toBeDefined();
    await equalPeer(peer, session);
    await replayEdits(peer, session);
  } finally {
    warning.mockRestore();
    sourceOpening.mockRestore();
    peer?.dispose();
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

test('snapshot hydration transfers bounded chunk batches and yields between fixed-budget slices', async () => {
  const buffers: ArrayBuffer[] = [];
  const transfers: boolean[] = [];
  const batches: number[][] = [];
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
        requests.get(message.id) === 'pullPeerSnapshot' && Array.isArray(message.value)) {
        const chunks = message.value as ArrayBuffer[];
        batches.push(chunks.map((chunk) => chunk.byteLength));
        buffers.push(...chunks);
        transfers.push(transfer?.length === chunks.length && chunks.every((chunk, index) =>
          chunk instanceof ArrayBuffer && transfer?.[index] === chunk
        ));
      }
      transport.post(message, transfer);
    },
  }));
  const createBuilder = workbookWasm.createWorkbookSnapshotBuilder;
  const advances: number[][] = [];
  const turns: boolean[] = [];
  let now = 0;
  const clock = spyOn(performance, 'now').mockImplementation(() => now);
  let yielded = true;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const building = spyOn(workbookWasm, 'createWorkbookSnapshotBuilder').mockImplementation((options) => {
    const builder = createBuilder(options);
    return { ...builder, advance(records, bytes) {
      turns.push(yielded);
      if (yielded) timer = setTimeout(() => { yielded = true; }, 0);
      yielded = false;
      advances.push([records, bytes]);
      now += 4;
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
    expect(batches.length).toBeGreaterThan(1);
    expect(batches.some((batch) => batch.length > 1)).toBe(true);
    expect(batches.every((batch) => batch.length <= 256 &&
      batch.reduce((total, bytes) => total + bytes, 0) <= 16 * 1024
    )).toBe(true);
    expect(transfers.every(Boolean)).toBe(true);
    expect(buffers.every((buffer) => buffer.byteLength === 0)).toBe(true);
    expect(advances.length).toBeGreaterThan(1);
    expect(advances.every(([records, bytes]) => records === 256 && bytes === 16 * 1024)).toBe(true);
    expect(turns.slice(0, 3)).toEqual([true, false, true]);
    expect(turns).toEqual(advances.map((_, index) => index % 2 === 0));
    await equalPeer(peer, session);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    clock.mockRestore();
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
