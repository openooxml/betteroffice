import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { isClientMessage, isHostMessage, SessionFailure } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as wasm from '../wasm/loader';
import { createWorkbookSession, type WorkbookSession } from './client';
import type { WorkbookAck, WorkbookCommand, WorkbookCommandFrameRequest } from './commandProtocol';
import { createWorkbookSessionHost } from './host';
import { batchRequests, createTestWorkbookSession } from './testHelpers';

let bytes: Uint8Array;

beforeAll(async () => {
  await wasm.initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData></worksheet>');
  bytes = await zip.generateAsync({ type: 'uint8array' });
});

const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
const frameRequest: WorkbookCommandFrameRequest = {
  sheet: 0, viewport: { x: 0, y: 0, width: 384, height: 96 }, zoom: 1, dpr: 2,
  generation: 7, requestGeneration: 11,
};

function command(session: WorkbookSession, seq: number, kind: WorkbookCommand['kind'], payload: unknown,
  afterSeq?: number): WorkbookCommand {
  return { ...session.commandIdentity, basis: session.confirmedBasis, seq, kind, payload,
    ...(afterSeq === undefined ? {} : { afterSeq }) } as WorkbookCommand;
}

async function trackedSession(configure: (handle: wasm.WorkbookHandle) => void = () => {}) {
  const original = wasm.openWorkbook;
  let handle: wasm.WorkbookHandle | undefined;
  const opening = spyOn(wasm, 'openWorkbook').mockImplementation((...args) => {
    handle = original(...args);
    configure(handle);
    return handle;
  });
  try {
    const session = await createTestWorkbookSession(bytes);
    if (!handle) throw new Error('Missing worker workbook');
    return { session, handle };
  } finally { opening.mockRestore(); }
}

describe('workbook command swap compatibility', () => {
  test('queued native edit 65 refuses compatibility sequence 65 before retry 66 and save', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const sequences = new Map<number, number>();
    const releases = new Map<number, () => void>();
    const held = new Map<number, { promise: Promise<void>; resolve(): void }>();
    const events: string[] = [];
    function whenHeld(seq: number) {
      let waiter = held.get(seq);
      if (!waiter) {
        let resolve!: () => void;
        const promise = new Promise<void>((done) => { resolve = done; });
        waiter = { promise, resolve };
        held.set(seq, waiter);
      }
      return waiter;
    }
    function release(seq: number) {
      const deliver = releases.get(seq);
      if (!deliver) throw new Error(`Missing reply ${seq}`);
      releases.delete(seq);
      deliver();
    }
    const session = await createWorkbookSession(bytes, {}, {
      ...pair.client,
      post(message, transfer) {
        if (isClientMessage(message) && message.kind === 'call') {
          if (message.method === 'edit' || message.method === 'command') {
            const input = message.args[0];
            if (!input || typeof input !== 'object' || !('seq' in input) || typeof input.seq !== 'number') {
              throw new Error('Missing command sequence');
            }
            sequences.set(message.id, input.seq);
            events.push(`send:${input.seq}:${message.method}`);
          } else if (message.method === 'save') events.push('save');
        }
        pair.client.post(message, transfer);
      },
      listen: (listener) => pair.client.listen((message) => {
        if (isHostMessage(message) && message.kind === 'reply') {
          const seq = sequences.get(message.id);
          if (seq !== undefined) {
            releases.set(seq, () => listener(message));
            whenHeld(seq).resolve();
            return;
          }
        }
        listener(message);
      }),
    });
    try {
      const edits = session.editSession;
      if (!edits) throw new Error('Missing workbook edit session');
      const first = Array.from({ length: 64 }, (_, row) =>
        edits.submit('editCell', { payload: { args: [0, row, 0, `native:${row + 1}`] } }));
      await Promise.all(first.map((_, index) => whenHeld(index + 1).promise));
      const queued = edits.submit('editCell', { payload: { args: [0, 64, 0, 'queued native'] } });
      expect(queued.seq).toBeUndefined();
      await expect(session.submit(command(session, 65, 'editCell', { args: [0, 65, 0, 'compatibility'] })))
        .rejects.toThrow('Workbook command sequence is awaiting native dispatch');
      events.push('refused:65');
      expect(events).toEqual([...Array.from({ length: 64 }, (_, index) => `send:${index + 1}:edit`), 'refused:65']);
      release(1);
      await whenHeld(65).promise;
      expect(queued.seq).toBe(65);
      const compatibility = session.submit(command(session, 66, 'editCell', { args: [0, 65, 0, 'compatibility'] }));
      const saving = session.save();
      let saved = false;
      void saving.then(() => { saved = true; }, () => {});
      release(2);
      await whenHeld(66).promise;
      expect(saved).toBe(false);
      expect(events).toEqual([...Array.from({ length: 64 }, (_, index) => `send:${index + 1}:edit`),
        'refused:65', 'send:65:edit', 'send:66:command']);
      for (let seq = 3; seq <= 66; seq += 1) release(seq);
      await Promise.all(first.map((ticket) => ticket.ack));
      expect(await queued.ack).toMatchObject({ seq: 65, status: 'applied' });
      expect(await compatibility).toMatchObject({ seq: 66, status: 'applied', committed: true });
      const reopened = wasm.openWorkbook(await saving);
      try {
        expect(reopened.cell(0, 64, 0).input).toBe('queued native');
        expect(reopened.cell(0, 65, 0).input).toBe('compatibility');
      } finally { reopened.dispose(); }
      expect(saved).toBe(true);
      expect(session.throughSeq).toBe(66);
      expect(session.failure).toBeUndefined();
      expect(events).toEqual([...Array.from({ length: 64 }, (_, index) => `send:${index + 1}:edit`),
        'refused:65', 'send:65:edit', 'send:66:command', 'save']);
    } finally { await session.dispose(); }
  });

  test('keeps engine tokens opaque and separate from command and scheduler counters', async () => {
    let restore = () => {};
    const { session } = await trackedSession((handle) => {
      const version = handle.version.bind(handle);
      const tokens = new Map<string, string>();
      const opaque = spyOn(handle, 'version').mockImplementation(() => {
        const key = version();
        if (!tokens.has(key)) tokens.set(key, ['engine:initial/α', 'engine:next/β', 'engine:final/γ'][tokens.size]);
        return tokens.get(key)!;
      });
      restore = () => opaque.mockRestore();
    });
    try {
      expect(session.confirmedBasis.version).toBe('engine:initial/α');
      const basis = session.confirmedBasis;
      const first = session.submit(command(session, 1, 'patchRangeStyle', { args: [0, 'A1', { bold: true }], target }));
      const second = session.submit(command(session, 2, 'patchRangeStyle', { args: [0, 'A2', { italic: true }], target: {
        ...target, range: { kind: 'a1', a1: 'A2' },
      } }, 1));
      expect(basis.version).toBe('engine:initial/α');
      for (const outcome of await Promise.all([first, second])) {
        expect(outcome).toMatchObject({ status: 'applied', committed: true });
      }
      expect(session.confirmedBasis).toEqual({ epoch: basis.epoch, version: 'engine:final/γ' });
      expect(await session.call.version()).toBe('engine:final/γ');
      expect(session.throughSeq).toBe(2);
      expect(session.state.version).toBe(2);
      const saved = wasm.openWorkbook(await session.save());
      try {
        expect(saved.selectionFormatting(0, 'A1')).toMatchObject({ bold: true });
        expect(saved.selectionFormatting(0, 'A2')).toMatchObject({ italic: true });
      } finally { saved.dispose(); }
    } finally { restore(); await session.dispose(); }
  });

  test('selection mode and format capture succeed without document history or version changes', async () => {
    const { session, handle } = await trackedSession();
    try {
      const before = await session.save();
      const version = session.confirmedBasis.version;
      const history = handle.historyState();
      const selection = { ...target, range: { kind: 'a1' as const, a1: 'A2' }, activeCell: { row: 1, col: 0 } };
      expect(await session.submit(command(session, 1, 'select', { selection })))
        .toEqual({ seq: 1, status: 'applied', committed: false });
      expect(await session.submit(command(session, 2, 'select', { selection })))
        .toEqual({ seq: 2, status: 'noop', committed: false });
      expect(session.selection).toEqual(selection);
      expect(await session.submit(command(session, 3, 'setReadOnly', { readOnly: true })))
        .toEqual({ seq: 3, status: 'applied', committed: false });
      expect(session.editState.readOnly).toBe(true);
      expect(await session.submit(command(session, 4, 'setReadOnly', { readOnly: false })))
        .toEqual({ seq: 4, status: 'applied', committed: false });
      const capture = command(session, 5, 'applyOps', { toolbar: { id: 'paintFormat', args: null, target } });
      const first = await session.submit(capture);
      expect(first).toEqual({ seq: 5, status: 'applied', committed: false,
        result: { formatToken: expect.any(String) } });
      expect(await session.submit(capture)).toEqual(first);
      const second = await session.submit(command(session, 6, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target },
      }));
      expect(second.result).not.toEqual(first.result);
      expect(second).toMatchObject({ status: 'applied', committed: false });
      expect(session.throughSeq).toBe(6);
      expect(session.confirmedBasis.version).toBe(version);
      expect(session.state.version).toBe(0);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
    } finally { await session.dispose(); }
  });

  test('retains captured format results across later capture and source edits', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', { args: [0, 'A1', { bold: true }] }));
      const first = await session.submit(command(session, 2, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target },
      }));
      await session.submit(command(session, 3, 'patchRangeStyle', { args: [0, 'A1', { bold: false, italic: true }] }));
      const second = await session.submit(command(session, 4, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target },
      }));
      const firstToken = (first.result as { formatToken: string }).formatToken;
      const secondToken = (second.result as { formatToken: string }).formatToken;
      expect(firstToken).not.toBe(secondToken);
      await session.submit(command(session, 5, 'applyFormat', { formatToken: firstToken,
        target: { ...target, range: { kind: 'a1', a1: 'A2' } } }));
      await session.submit(command(session, 6, 'applyFormat', { formatToken: secondToken,
        target: { ...target, range: { kind: 'a1', a1: 'A3' } } }));
      const saved = wasm.openWorkbook(await session.save());
      try {
        expect(saved.selectionFormatting(0, 'A2')).toMatchObject({ bold: true, italic: false });
        expect(saved.selectionFormatting(0, 'A3')).toMatchObject({ bold: false, italic: true });
      } finally { saved.dispose(); }
    } finally { await session.dispose(); }
  });

  test('returns edit and proposal results intact including stale acceptance details', async () => {
    const { session, handle } = await trackedSession();
    const independent = wasm.openWorkbook(bytes);
    try {
      const expected = independent.editCell(0, 0, 0, '7');
      expect(await session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '7'] })))
        .toEqual({ seq: 1, status: 'applied', committed: true, result: expected });
      const proposed = await session.submit(command(session, 2, 'propose', {
        args: ['author', 'Review', [{ sheet: 0, row: 0, col: 0, input: '9' }]],
      }));
      expect(proposed).toMatchObject({ status: 'applied', committed: false });
      expect(proposed.result).toEqual(handle.listProposals()[0]);
      const id = (proposed.result as wasm.Proposal).id;
      await session.submit(command(session, 3, 'editCell', { args: [0, 0, 0, '8'] }));
      const before = await session.save();
      const version = session.confirmedBasis.version;
      const history = handle.historyState();
      expect(await session.submit(command(session, 4, 'acceptProposal', { args: [id] })))
        .toMatchObject({ seq: 4, status: 'refused', reason: 'stale-proposal', committed: false,
          affectedSeqs: [4], result: { cells: ['A1'] } });
      const refresh = await session.submit(command(session, 5, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null },
      }));
      expect(refresh).toEqual({ seq: 5, status: 'applied', committed: false,
        result: { proposals: handle.listProposals() } });
      expect(session.confirmedBasis.version).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
      expect(session.failure).toBeUndefined();
    } finally { independent.dispose(); await session.dispose(); }
  });

  test('strict batches refuse stale tokens and cancel descendants without blocking independent reads and save', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const basis = session.confirmedBasis;
      const first = session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '11'] }));
      const refused = session.submit(command(session, 2, 'applyEdits', { args: [{
        expectVersion: basis.version, steps: [{ op: 'setCellInputs', target, inputs: [['12']] }],
      }] }, 1));
      const dependent = session.submit(command(session, 3, 'editCell', { args: [0, 1, 0, 'dependent'] }, 2));
      const independent = session.submit(command(session, 4, 'editCell', { args: [0, 2, 0, 'independent'] }));
      const read = session.call.cellInputs(0, 'A1:A3');
      const saving = session.save();
      expect(await first).toMatchObject({ status: 'applied', committed: true });
      expect(await refused).toMatchObject({ seq: 2, status: 'refused', reason: 'stale-version',
        committed: false, affectedSeqs: [2] });
      expect(await dependent).toEqual({ seq: 3, status: 'cancelled', reason: 'dependency-refused',
        committed: false, affectedSeqs: [3] });
      expect(await independent).toMatchObject({ seq: 4, status: 'applied', committed: true });
      expect((await read).cells.map((row) => row[0].input)).toEqual(['11', '', 'independent']);
      const reopened = wasm.openWorkbook(await saving);
      try {
        expect([0, 1, 2].map((row) => reopened.cell(0, row, 0).input)).toEqual(['11', '', 'independent']);
      } finally { reopened.dispose(); }
      expect(session.throughSeq).toBe(4);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('bare and duplicate acks publish state while presentation requires explicit covering paint', async () => {
    const session = await createTestWorkbookSession(bytes);
    const acks: WorkbookAck[] = [];
    const events: string[] = [];
    const off = session.onAck((ack) => { acks.push(ack); events.push(`ack:${ack.statuses[0].seq}`); });
    try {
      const input = { ...command(session, 1, 'editCell', { args: [0, 0, 0, '7'] }), frame: frameRequest };
      const ack = await session.call.command(input);
      expect(ack.frame).toMatchObject({ ...session.commandIdentity, coversThroughSeq: 1,
        requestGeneration: 11, zoom: 1, dpr: 2 });
      expect(ack.frame?.displayList.commands.some((entry) => entry.op === 'text' && entry.text === '14')).toBe(true);
      expect(session.presentedThroughSeq).toBe(0);
      let presented = false;
      const waiting = session.whenPresented(1).then(() => { presented = true; events.push('presented:1'); });
      const duplicate = await session.call.command(input);
      expect(duplicate.frame).toBeUndefined();
      expect(duplicate.statuses).toEqual(ack.statuses);
      expect(await session.submit(command(session, 2, 'select', { selection: {
        ...target, range: { kind: 'a1', a1: 'A2' },
      } }))).toMatchObject({ status: 'applied', committed: false });
      expect(acks).toHaveLength(3);
      expect(acks[2].frame).toBeUndefined();
      expect(session.throughSeq).toBe(2);
      expect(presented).toBe(false);
      expect(session.markPresented({ ...session.commandIdentity, incarnation: 'retired', coversThroughSeq: 1 })).toBe(false);
      expect(session.markPresented({ ...session.commandIdentity, coversThroughSeq: 3 })).toBe(false);
      expect(session.markPresented(ack.frame!)).toBe(true);
      await waiting;
      expect(session.presentedThroughSeq).toBe(1);
      expect(session.markPresented({ ...session.commandIdentity, coversThroughSeq: 0 })).toBe(false);
      expect(session.markPresented({ ...session.commandIdentity, coversThroughSeq: 2 })).toBe(true);
      await session.whenPresented(2);
      events.push('presented:2');
      expect(events).toEqual(['ack:1', 'ack:1', 'ack:2', 'presented:1', 'presented:2']);
    } finally { off(); await session.dispose(); }
  });

  test('postcommit traps reject commands and paint waiters without recovery or late reply publication', async () => {
    const pair = createInProcessPair();
    createWorkbookSessionHost(pair.host);
    const requests: string[] = [];
    let signalFailure: ((error: unknown) => void) | undefined;
    let release = () => {};
    let held!: () => void;
    const replyHeld = new Promise<void>((resolve) => { held = resolve; });
    const commandIds = new Set<number>();
    const session = await createWorkbookSession(bytes, {}, {
      ...pair.client,
      post(message, transfer) {
        if (isClientMessage(message) && message.kind === 'call') {
          requests.push(message.method);
          if (message.method === 'command') commandIds.add(message.id);
        }
        pair.client.post(message, transfer);
      },
      listen: (listener) => pair.client.listen((message) => {
        if (isHostMessage(message) && message.kind === 'reply' && commandIds.has(message.id)) {
          release = () => listener(message);
          held();
        } else listener(message);
      }),
      onError(listener) { signalFailure = listener; return pair.client.onError(listener); },
    });
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      const editing = session.submit({ ...command(session, 1, 'editCell', {
        args: [0, 0, 0, 'committed before trap'],
      }), frame: frameRequest });
      const painting = session.whenPresented(1);
      const settled = Promise.allSettled([editing, painting]);
      await replyHeld;
      expect(session.state.dirty).toBe(true);
      const failure = new SessionFailure('trap', 'Worker trapped before reply');
      signalFailure!(failure);
      for (const outcome of await settled) {
        expect(outcome).toMatchObject({ status: 'rejected', reason: failure });
      }
      release();
      await Promise.resolve();
      expect(acks).toEqual([]);
      expect(session.state.stage).toBe('failed');
      expect(session.throughSeq).toBe(0);
      expect(session.presentedThroughSeq).toBe(0);
      expect(session.markPresented({ ...session.commandIdentity, coversThroughSeq: 0 })).toBe(false);
      await expect(session.save()).rejects.toBeInstanceOf(SessionFailure);
      await expect(session.submit(command(session, 2, 'editCell', { args: [0, 1, 0, 'blocked'] })))
        .rejects.toBeInstanceOf(SessionFailure);
      expect(requests).toEqual(['open', 'command']);
    } finally { await session.dispose(); }
  });
});
