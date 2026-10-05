import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { isHostMessage, SessionFailure, type SessionTransport } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as workbookWasm from '../wasm/loader';
import { XlsxDocument } from '../wasm/generated/xlsx_wasm.js';
import { openWorkbook, type WorkbookCalculationContext, type WorkbookHandle } from '../wasm/loader';
import { createWorkbookSession, hydratePeer } from './client';
import { createWorkbookEditPeer, WorkbookEditPeerFailedError } from './editPeer';
import { workbookEditPeerOperations } from './editPeerInternals';
import { createWorkbookSessionHost } from './host';
import { WORKBOOK_SESSION_POLICIES } from './methods';
import { WorkbookPeerHydrationError } from './peerHydrationError';
import type { WorkbookReplayOp } from './replay';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let fixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };

beforeAll(async () => {
  ({ fixture, wasmBytes } = await loadWorkbookSessionFixtures());
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

describe('workbook peer hydration', () => {
  test('refuses retained worker mutations before attachment without changing document state', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(fixture, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    }, pair.client);
    try {
      const version = await session.call.version();
      const saved = await session.save();
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
      expect(mutators).toEqual(['applyEdits']);
      await expect(session.call.applyEdits(input)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
      await expect(session.call.applyEdits(input)).rejects.toMatchObject({ code: 'mutation-before-attachment' });
      const reply = new Promise<unknown>((resolve) => {
        const off = pair.client.listen((message) => {
          if (isHostMessage(message) && message.kind === 'reply' && message.id === 10_000) {
            off();
            resolve(message);
          }
        });
      });
      pair.client.post({ protocol: 1, kind: 'call', id: 10_000, method: 'applyEdits', args: [input] });
      expect(await reply).toMatchObject({
        ok: false, error: { name: 'WorkbookPeerHydrationError', refusal: { code: 'mutation-before-attachment' } },
      });
      expect(await session.call.version()).toBe(version);
      expect(await session.save()).toEqual(saved);
      expect(session.state).toMatchObject({ version: 0, dirty: false, stage: 'ready' });
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(0);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
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
          message = { ...message, hydration: JSON.stringify({ ...JSON.parse(message.hydration), sequence: 1 }) };
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

  test('adopts worker volatile values and version without recalculating during hydration', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Volatile" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>NOW()</f><v>0</v></c><c r="B1"><f>TODAY()</f><v>0</v></c><c r="C1"><f>RAND()</f><v>0</v></c></row></sheetData></worksheet>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
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
