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
import { createWorkbookEditPeer, WorkbookEditPeerFailedError, type WorkbookEditPeer } from './editPeer';
import { workbookEditPeerOperations } from './editPeerInternals';
import { createWorkbookSessionHost } from './host';
import { WORKBOOK_SESSION_POLICIES } from './methods';
import { WorkbookPeerHydrationError } from './peerHydrationError';
import { workbookSessionInternals, type WorkbookReplayEnvelope, type WorkbookReplayOp } from './replay';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

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
  test('refuses retained public mutations before attachment and after peer disposal without changing document state', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const session = await createWorkbookSession(fixture, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    }, pair.client);
    let peer: WorkbookHandle | undefined;
    let edits: WorkbookEditPeer | undefined;
    const applyEdits = session.call.applyEdits;
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
        expect(mutators).toEqual(['applyEdits']);
        await expect(session.call.applyEdits(input)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
        await expect(session.call.applyEdits(input)).rejects.toMatchObject({ code: 'mutation-outside-replay' });
        await expect(applyEdits(input)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
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
          ok: false, error: { name: 'WorkbookPeerHydrationError', refusal: { code: 'mutation-outside-replay' } },
        });
        const internal = workbookSessionInternals.get(session);
        if (!internal) throw new Error('Missing internal workbook replay helper');
        const envelope: WorkbookReplayEnvelope = {
          sequence: 1, calculation, op: { method: 'editCell', args: [0, 0, 0, 'refused replay'] },
        };
        await expect(internal.replay(envelope)).rejects.toBeInstanceOf(WorkbookPeerHydrationError);
        await expect(internal.replay(envelope)).rejects.toMatchObject({ code: 'mutation-outside-replay' });
        expect(await session.call.version()).toBe(version);
        expect(await session.save()).toEqual(saved);
        expect(session.state).toMatchObject({ version: 0, dirty: false, stage: 'ready' });
        expect(session.state).toEqual(state);
        expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBe(0);
        expect(session.failure).toBeUndefined();
      }
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
