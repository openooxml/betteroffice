import { beforeAll, describe, expect, test } from 'bun:test';
import { isClientMessage, SessionFailure, type SessionTransport } from '../../../../shared/office-session';
import type { XlsxEditRequest, XlsxRangeTarget } from '../edits';
import {
  openWorkbook, type Viewport, type WorkbookCalculationContext, type WorkbookHandle,
} from '../wasm/loader';
import type { WorkbookSession } from './client';
import {
  createWorkbookEditPeer, WorkbookEditPeerFailedError, type WorkbookEditPeer,
} from './editPeer';
import { WORKBOOK_SESSION_METHODS, WORKBOOK_SESSION_POLICIES } from './methods';
import {
  WORKBOOK_INTERNAL_SESSION_POLICIES,
  workbookSessionInternals,
  type WorkbookReplayEnvelope,
  type WorkbookReplayMethod,
} from './replay';
import { batchRequests, createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let fixture: Uint8Array;
let chartFixture: Uint8Array;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };
const viewport: Viewport = { x: 0, y: 0, width: 800, height: 800 };

beforeAll(async () => {
  ({ fixture, chartFixture } = await loadWorkbookSessionFixtures());
});

function target(a1: string): XlsxRangeTarget {
  return { sheetId: 'sheet:0', range: { kind: 'a1', a1 } };
}

function request(peer: WorkbookHandle, a1: string, input: string): XlsxEditRequest {
  return {
    expectVersion: peer.version(),
    steps: [{ op: 'setCellInputs', target: target(a1), inputs: [[input]] }],
  };
}

async function digest(bytes: Uint8Array | ArrayBuffer): Promise<string> {
  const owned = new Uint8Array(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const hash = await crypto.subtle.digest('SHA-256', owned);
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function matchingDigest(
  batch: number, method: WorkbookReplayMethod | 'open',
  edits: WorkbookEditPeer, peer: WorkbookHandle, session: WorkbookSession
): Promise<void> {
  await edits.flush();
  const [local, worker] = await Promise.all([digest(peer.save()), session.save().then(digest)]);
  if (local !== worker) throw new Error(`Workbook digest mismatch after batch ${batch} (${method}): ${local} !== ${worker}`);
  expect(edits.acknowledgedSequence).toBe(edits.sentSequence);
}

function deterministicOptions(): { now(): number; randomSeed(): number } {
  let tick = 0;
  let seed = 0;
  return {
    now: () => (calculation.nowSerial - 25_569) * 86_400_000 + tick++ * 1000,
    randomSeed: () => ++seed,
  };
}

function replay(session: WorkbookSession, envelope: WorkbookReplayEnvelope) {
  const internal = workbookSessionInternals.get(session);
  if (!internal) throw new Error('Missing internal workbook replay helper');
  return internal.replay(envelope);
}

function recordReplays(envelopes: WorkbookReplayEnvelope[]): (transport: SessionTransport) => SessionTransport {
  return (transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
        envelopes.push(structuredClone(message.args[0]) as WorkbookReplayEnvelope);
      }
      listener(message);
    }),
  });
}

describe('workbook edit peers', () => {
  test('keeps replay internal, ordered and never replaceable', () => {
    expect(WORKBOOK_SESSION_METHODS).not.toHaveProperty('replay');
    expect(WORKBOOK_SESSION_POLICIES).not.toHaveProperty('replay');
    expect(WORKBOOK_INTERNAL_SESSION_POLICIES.replay).toEqual({
      lane: 'input', mutates: true, userInput: true, reorderable: false,
    });
    expect(WORKBOOK_INTERNAL_SESSION_POLICIES.replay).not.toHaveProperty('replaceableBy');
  });

  test('rejects public mutations while an edit peer is attached and restores them on disposal', async () => {
    const peer = openWorkbook(fixture, { calculation });
    const session = await createTestWorkbookSession(fixture, batchRequests, { calculation });
    const applyEdits = session.call.applyEdits;
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    try {
      const initial = await digest(await session.save());
      const input = request(peer, 'B3', '901');
      await expect(session.call.applyEdits(input)).rejects.toThrow('edit peer is attached');
      await expect(applyEdits(input)).rejects.toThrow('edit peer is attached');
      expect(() => createWorkbookEditPeer({ session, peer })).toThrow('already has an attached edit peer');
      expect(await digest(await session.save())).toBe(initial);
      await matchingDigest(0, 'open', edits, peer, session);
      expect(session.state).toMatchObject({ version: 0, dirty: false, stage: 'ready' });
      expect(session.failure).toBeUndefined();
      expect(edits.editCell(0, 2, 1, '902').applied).toBe(true);
      await matchingDigest(1, 'editCell', edits, peer, session);
      const beforeDispose = await digest(await session.save());
      edits.dispose();
      const restored = { ...request(peer, 'B3', '903'), expectVersion: await session.call.version() };
      expect(await applyEdits(restored)).toMatchObject({ ok: true, applied: true });
      expect((await session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('903');
      expect(await digest(await session.save())).not.toBe(beforeDispose);
      expect(session.failure).toBeUndefined();
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('serves an immediate frame after a peer sheet change without awaiting flush', async () => {
    for (const wrapHost of [undefined, batchRequests]) {
      const peer = openWorkbook(fixture, { calculation });
      const session = await createTestWorkbookSession(fixture, wrapHost, { calculation });
      const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
      try {
        expect(peer.sheetCount()).toBeGreaterThan(1);
        edits.setActiveSheet(1);
        const pendingFrame = session.call.frame(viewport);
        const sentSequence = edits.sentSequence;
        expect(edits.acknowledgedSequence).toBe(0);
        const frame = await pendingFrame;
        expect(frame.sequence).toBe(sentSequence);
        expect(frame.sheet).toBe(1);
        expect(frame.displayList).toEqual(peer.displayList(viewport));
        expect(session.state.activeSheet).toBe(1);
        await matchingDigest(1, 'setActiveSheet', edits, peer, session);
      } finally {
        edits.dispose();
        peer.dispose();
        await session.dispose();
      }
    }
  });

  test('matches saved digests after every mutator batch, volatile formula and history boundary', async () => {
    const peer = openWorkbook(fixture, { calculation });
    const session = await createTestWorkbookSession(fixture, undefined, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    let batch = 0;
    async function check(method: WorkbookReplayMethod): Promise<void> {
      await matchingDigest(++batch, method, edits, peer, session);
      expect((await session.call.frame(viewport)).sequence).toBe(edits.acknowledgedSequence);
    }
    try {
      await matchingDigest(0, 'open', edits, peer, session);
      expect((await session.call.frame(viewport)).sequence).toBe(0);
      for (const [index, input] of [
        '42', 'plain text', 'TRUE', '=1+2', '=NOW()', '=TODAY()', '=RANDBETWEEN(1,1000000)', '=RANDBETWEEN(1,100)',
      ].entries()) {
        const result = edits.editCell(0, 9, index + 1, input);
        expect(result).not.toBeInstanceOf(Promise);
        expect(result.applied).toBe(true);
        expect(peer.cell(0, 9, index + 1).input).toBe(input);
        await check('editCell');
      }
      const volatile = peer.readCells({ ranges: [target('F10:I10')] });
      if (!volatile.ok) throw new Error(volatile.failure.message);
      for (const cell of volatile.ranges[0].cells[0]) expect(cell.value.kind).toBe('number');
      expect(edits.editCells(0, [
        { row: 10, col: 1, input: '17' },
        { row: 10, col: 2, input: '=B11*2' },
        { row: 10, col: 3, input: '=RANDBETWEEN(1,1000000)+RANDBETWEEN(1,100)+NOW()+TODAY()' },
      ]).applied).toBe(true);
      await check('editCells');
      const applied = edits.applyEdits({
        expectVersion: peer.version(),
        steps: [
          { op: 'setCellInputs', target: target('B12:C12'), inputs: [['12', '=B12+RANDBETWEEN(1,1000000)']] },
          { op: 'patchStyle', target: target('B12'), patch: { bold: true } },
          { op: 'setNumberFormat', target: target('C12'), format: 'percent' },
          { op: 'setFormulas', target: target('D12'), formulas: [['NOW()+TODAY()+RANDBETWEEN(1,10)']] },
        ],
      });
      expect(applied).toMatchObject({ ok: true, applied: true });
      await check('applyEdits');
      expect(edits.applyOps([
        { type: 'insertRows', sheet: 0, at: 60, count: 1 },
        { type: 'mergeCells', sheet: 0, range: {
          start: { row: 19, col: 9 }, end: { row: 19, col: 10 },
        } },
      ]).applied).toBe(true);
      await check('applyOps');
      expect(edits.patchRangeStyle(0, 'B10:D12', {
        italic: true, fontSize: 14, fillColor: '#ffcc00', border: { preset: 'outer' },
      }).applied).toBe(true);
      await check('patchRangeStyle');
      for (const format of ['percent', 'increaseDecimal', { type: 'custom', pattern: '0.000' }] as const) {
        edits.setNumberFormat(0, 'B10:C11', format);
        await check('setNumberFormat');
      }
      expect(edits.applyFormat(0, 'F12:G13', peer.captureFormat(0, 'B10:C11')).applied).toBe(true);
      await check('applyFormat');
      for (let index = 0; index < 6; index += 1) {
        expect(edits.undo().applied).toBe(true);
        await check('undo');
      }
      for (let index = 0; index < 6; index += 1) {
        expect(edits.redo().applied).toBe(true);
        await check('redo');
      }
      expect(peer.sheetCount()).toBeGreaterThan(1);
      for (const sheet of [1, 0, 0]) {
        expect(edits.setActiveSheet(sheet)).toBeUndefined();
        await check('setActiveSheet');
        expect(session.state.activeSheet).toBe(sheet);
        expect((await session.call.frame(viewport)).sheet).toBe(sheet);
      }
      const saved = await edits.save();
      expect(saved).toBeInstanceOf(ArrayBuffer);
      expect(await digest(saved)).toBe(await digest(peer.save()));
      expect(edits.state).toBe('ready');
      expect(edits.error).toBeUndefined();
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('replays movable charts and their undo and redo with equal saved digests', async () => {
    const peer = openWorkbook(chartFixture, { calculation });
    const session = await createTestWorkbookSession(chartFixture, undefined, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    try {
      const chart = peer.displayList(viewport).charts?.find((item) => item.movable);
      if (!chart) throw new Error('Chart fixture has no movable chart');
      expect(edits.moveChart(0, chart.id, 24, 12).applied).toBe(true);
      await matchingDigest(1, 'moveChart', edits, peer, session);
      expect(edits.undo().applied).toBe(true);
      await matchingDigest(2, 'undo', edits, peer, session);
      expect(edits.redo().applied).toBe(true);
      await matchingDigest(3, 'redo', edits, peer, session);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('orders synchronous edit bursts and flushes them before a normal worker save', async () => {
    const sequences: number[] = [];
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          sequences.push((message.args[0] as WorkbookReplayEnvelope).sequence);
        }
        listener(message);
      }),
    }), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const internal = workbookSessionInternals.get(session);
    if (!internal) throw new Error('Missing internal workbook replay helper');
    const submittedSequences: number[] = [];
    const submitReplay = internal.replay;
    internal.replay = (envelope) => {
      submittedSequences.push(envelope.sequence);
      return submitReplay(envelope);
    };
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    try {
      expect(edits.applyEdits(request(peer, 'B3', '301'))).toMatchObject({ ok: true, applied: true });
      expect(edits.applyEdits(request(peer, 'C3', '302'))).toMatchObject({ ok: true, applied: true });
      expect(edits.editCells(0, [{ row: 2, col: 3, input: '=B3+C3+1' }]).applied).toBe(true);
      expect(edits.patchRangeStyle(0, 'B3:D3', { italic: true }).applied).toBe(true);
      expect(edits.sentSequence).toBe(4);
      expect(edits.acknowledgedSequence).toBe(0);
      expect(submittedSequences).toEqual([1, 2, 3, 4]);
      const pendingFrame = session.call.frame(viewport);
      const [saved, frame] = await Promise.all([edits.save(), pendingFrame, edits.flush(), edits.flush()]);
      expect(frame.sequence).toBe(edits.sentSequence);
      expect(frame.displayList).toEqual(peer.displayList(viewport));
      expect(await digest(saved)).toBe(await digest(peer.save()));
      await matchingDigest(1, 'patchRangeStyle', edits, peer, session);
      expect(sequences).toEqual([1, 2, 3, 4]);
      expect((await session.call.frame(viewport)).sequence).toBe(4);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('replays listener overwrites in local commit order before an immediate frame', async () => {
    const envelopes: WorkbookReplayEnvelope[] = [];
    const session = await createTestWorkbookSession(fixture, recordReplays(envelopes), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    let nested = false;
    let nestedApplied = false;
    let sentDuringUpdate = -1;
    const offUpdate = peer.onUpdate(() => {
      if (nested) return;
      nested = true;
      nestedApplied = edits.editCell(0, 0, 0, '2').applied;
      sentDuringUpdate = edits.sentSequence;
    });
    try {
      expect(edits.editCell(0, 0, 0, '1').applied).toBe(true);
      const pendingFrame = session.call.frame(viewport);
      expect(nestedApplied).toBe(true);
      expect(sentDuringUpdate).toBe(0);
      expect(peer.cell(0, 0, 0).input).toBe('2');
      expect(peer.readCells({ ranges: [target('A1')] })).toMatchObject({
        ok: true, ranges: [{ cells: [[{ value: { kind: 'number', value: 2 } }]] }],
      });
      expect(edits.sentSequence).toBe(2);
      expect(edits.acknowledgedSequence).toBe(0);
      await matchingDigest(1, 'editCell', edits, peer, session);
      expect(envelopes.map(({ sequence, op }) => ({ sequence, op }))).toEqual([
        { sequence: 1, op: { method: 'editCell', args: [0, 0, 0, '1'] } },
        { sequence: 2, op: { method: 'editCell', args: [0, 0, 0, '2'] } },
      ]);
      const frame = await pendingFrame;
      expect(frame.sequence).toBe(2);
      expect(frame.displayList).toEqual(peer.displayList(viewport));
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('2');
    } finally {
      offUpdate();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('waits for nested listener edits in a listener-started flush', async () => {
    const session = await createTestWorkbookSession(fixture, batchRequests, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    let nested = false;
    let nestedApplied = false;
    let pendingFlush: Promise<void> | undefined;
    const offUpdate = peer.onUpdate(() => {
      if (nested) return;
      nested = true;
      nestedApplied = edits.editCell(0, 0, 1, '2').applied;
      pendingFlush = edits.flush();
    });
    try {
      expect(edits.editCell(0, 0, 0, '1').applied).toBe(true);
      expect(nestedApplied).toBe(true);
      if (!pendingFlush) throw new Error('Missing listener flush');
      await pendingFlush;
      expect(edits.acknowledgedSequence).toBe(2);
      expect(edits.sentSequence).toBe(2);
      expect((await session.call.cellInputs(0, 'B1')).cells[0][0].input).toBe('2');
      expect(await digest(await edits.save())).toBe(await digest(peer.save()));
    } finally {
      offUpdate();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('rejects a listener-started flush when the worker refuses the nested edit', async () => {
    const session = await createTestWorkbookSession(fixture, (transport) => {
      const batched = batchRequests(transport);
      return { ...batched, listen: (listener) => batched.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          const envelope = message.args[0] as WorkbookReplayEnvelope;
          if (envelope.op.method === 'applyEdits') {
            const input = envelope.op.args[0];
            envelope.op.args = [{ ...input, steps: [{
              op: 'setCellInputs', target: { ...target('B3'), sheetId: 'missing-sheet' }, inputs: [['2']],
            }] }];
          }
        }
        listener(message);
      }) };
    }, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    let nested = false;
    let nestedResult: ReturnType<WorkbookEditPeer['applyEdits']> | undefined;
    let pendingFlush: Promise<void> | undefined;
    const offUpdate = peer.onUpdate(() => {
      if (nested) return;
      nested = true;
      nestedResult = edits.applyEdits(request(peer, 'B3', '2'));
      pendingFlush = edits.flush();
    });
    try {
      expect(edits.editCell(0, 0, 0, '1').applied).toBe(true);
      expect(nestedResult).toMatchObject({ ok: true, applied: true });
      if (!pendingFlush) throw new Error('Missing listener flush');
      await expect(pendingFlush).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(edits.sentSequence).toBe(2);
      expect(edits.state).toBe('failed');
      expect(edits.error).toBe(session.failure);
      expect(session.failure?.diagnostics).toContain('sequence 2 (applyEdits)');
      expect(session.failure?.diagnostics).toContain('missing-target');
    } finally {
      offUpdate();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('waits for acknowledgement in a replay-started flush', async () => {
    const session = await createTestWorkbookSession(fixture, batchRequests, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const internal = workbookSessionInternals.get(session);
    if (!internal) throw new Error('Missing internal workbook replay helper');
    const submitReplay = internal.replay;
    let pendingFlush: Promise<void> | undefined;
    internal.replay = (envelope) => {
      if (!pendingFlush) pendingFlush = edits.flush();
      return submitReplay(envelope);
    };
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    try {
      expect(edits.editCell(0, 2, 1, '901').applied).toBe(true);
      if (!pendingFlush) throw new Error('Missing replay flush');
      await pendingFlush;
      expect(edits.acknowledgedSequence).toBe(1);
      expect(edits.sentSequence).toBe(1);
      expect((await session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('901');
    } finally {
      internal.replay = submitReplay;
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('omits nested peer refusals and throws without sequence gaps', async () => {
    const envelopes: WorkbookReplayEnvelope[] = [];
    const session = await createTestWorkbookSession(fixture, recordReplays(envelopes), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    let nested = false;
    let refusal: ReturnType<WorkbookEditPeer['applyEdits']> | undefined;
    let thrown: unknown;
    let nestedApplied = false;
    const offUpdate = peer.onUpdate(() => {
      if (nested) return;
      nested = true;
      refusal = edits.applyEdits({ ...request(peer, 'B1', '2'), expectVersion: 'stale' });
      try { edits.moveChart(0, 'missing-chart', 1, 0); } catch (cause) { thrown = cause; }
      nestedApplied = edits.editCell(0, 0, 1, '3').applied;
    });
    try {
      expect(edits.editCell(0, 0, 0, '1').applied).toBe(true);
      expect(refusal).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(thrown).toBeInstanceOf(Error);
      expect(nestedApplied).toBe(true);
      expect(edits.sentSequence).toBe(2);
      await matchingDigest(1, 'editCell', edits, peer, session);
      expect(envelopes.map(({ sequence, op }) => ({ sequence, op }))).toEqual([
        { sequence: 1, op: { method: 'editCell', args: [0, 0, 0, '1'] } },
        { sequence: 2, op: { method: 'editCell', args: [0, 0, 1, '3'] } },
      ]);
      expect(edits.editCell(0, 0, 2, '4').applied).toBe(true);
      await matchingDigest(2, 'editCell', edits, peer, session);
      expect(envelopes.map((envelope) => envelope.sequence)).toEqual([1, 2, 3]);
      expect(edits.state).toBe('ready');
    } finally {
      offUpdate();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('snapshots outer arguments before a listener can mutate them', async () => {
    const envelopes: WorkbookReplayEnvelope[] = [];
    const session = await createTestWorkbookSession(fixture, recordReplays(envelopes), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    const input = request(peer, 'A1', '1');
    const original = structuredClone(input);
    let mutated = false;
    const offUpdate = peer.onUpdate(() => {
      const first = input.steps[0];
      if (first.op !== 'setCellInputs') throw new Error('Missing input step');
      first.inputs[0][0] = '2';
      first.target.sheetId = 'missing-sheet';
      input.expectVersion = 'stale';
      mutated = true;
    });
    try {
      expect(edits.applyEdits(input)).toMatchObject({ ok: true, applied: true });
      expect(mutated).toBe(true);
      expect(input).not.toEqual(original);
      expect(peer.cell(0, 0, 0).input).toBe('1');
      expect(edits.sentSequence).toBe(1);
      await matchingDigest(1, 'applyEdits', edits, peer, session);
      expect(envelopes).toHaveLength(1);
      expect(envelopes[0].op).toEqual({ method: 'applyEdits', args: [original] });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('1');
    } finally {
      offUpdate();
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('sends successful no-ops, snapshots arguments and omits peer refusals and throws', async () => {
    const envelopes: WorkbookReplayEnvelope[] = [];
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          envelopes.push(structuredClone(message.args[0]) as WorkbookReplayEnvelope);
        }
        listener(message);
      }),
    }), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const edits = createWorkbookEditPeer({ session, peer, ...deterministicOptions() });
    const changes: number[] = [];
    session.on('changed', (event) => { changes.push(event.version); });
    let batch = 0;
    async function check(method: WorkbookReplayMethod): Promise<void> {
      await matchingDigest(++batch, method, edits, peer, session);
    }
    try {
      expect(() => edits.recoverySave()).toThrow('requires a failed');
      expect(edits.undo().applied).toBe(false);
      await check('undo');
      expect(edits.redo().applied).toBe(false);
      await check('redo');
      expect(edits.applyEdits(request(peer, 'B3', '100'))).toMatchObject({ ok: true, applied: false });
      await check('applyEdits');
      expect(changes).toEqual([]);
      expect(session.state).toMatchObject({ version: 0, dirty: false });
      expect(edits.sentSequence).toBe(3);
      const initialHistory = peer.historyState();
      expect(edits.applyEdits({ ...request(peer, 'B3', '900'), expectVersion: 'stale' }))
        .toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      await check('applyEdits');
      expect(() => edits.moveChart(0, 'missing-chart', 1, 0)).toThrow();
      await check('moveChart');
      expect(edits.sentSequence).toBe(3);
      expect(peer.historyState()).toEqual(initialHistory);
      const input = request(peer, 'B3', '900');
      expect(edits.applyEdits(input)).toMatchObject({ ok: true, applied: true });
      const first = input.steps[0];
      if (first.op !== 'setCellInputs') throw new Error('Missing input step');
      first.inputs[0][0] = 'tampered';
      await check('applyEdits');
      expect(edits.applyEdits(request(peer, 'B3', '900'))).toMatchObject({ ok: true, applied: false });
      await check('applyEdits');
      expect(edits.undo().applied).toBe(true);
      await check('undo');
      expect(peer.cell(0, 2, 1).input).toBe('100');
      expect(edits.redo().applied).toBe(true);
      await check('redo');
      expect(peer.cell(0, 2, 1).input).toBe('900');
      expect(envelopes.map((envelope) => envelope.sequence)).toEqual([1, 2, 3, 4, 5, 6, 7]);
      expect(envelopes.map((envelope) => envelope.op.method)).toEqual([
        'undo', 'redo', 'applyEdits', 'applyEdits', 'applyEdits', 'undo', 'redo',
      ]);
      expect(envelopes[3].op).toMatchObject({ args: [{ steps: [{ inputs: [['900']] }] }] });
      expect(new Set(envelopes.map((envelope) => envelope.calculation.randSeed)).size).toBe(7);
      expect(new Set(envelopes.map((envelope) => envelope.calculation.nowSerial)).size).toBe(7);
      expect(changes).toEqual([1, 2, 3]);
      expect((await session.call.frame(viewport)).sequence).toBe(7);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('rejects gaps, duplicates and malformed envelopes without mutation and remains usable', async () => {
    const session = await createTestWorkbookSession(fixture, batchRequests, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    try {
      const initial = await digest(await session.save());
      const initialVersion = await session.call.version();
      const envelope: WorkbookReplayEnvelope = {
        sequence: 1, calculation, op: { method: 'setActiveSheet', args: [0] },
      };
      await expect(replay(session, { ...envelope, sequence: 2 })).rejects.toMatchObject({
        name: 'WorkbookReplayOrderError',
      });
      expect(await digest(await session.save())).toBe(initial);
      peer.setCalculationContext(calculation);
      peer.setActiveSheet(0);
      expect(await replay(session, envelope)).toMatchObject({ sequence: 1, revision: 0, version: 0 });
      for (const sequence of [1, 3]) {
        await expect(replay(session, { ...envelope, sequence })).rejects.toMatchObject({
          name: 'WorkbookReplayOrderError',
        });
      }
      const malformedEnvelopes: unknown[] = [
        null,
        { ...envelope, sequence: 0 },
        { ...envelope, sequence: 2, calculation: { nowSerial: Infinity, randSeed: 0 } },
        { ...envelope, sequence: 2, calculation: { nowSerial: 46000, randSeed: -1 } },
        { ...envelope, sequence: 2, op: { method: 'save', args: [] } },
        { ...envelope, sequence: 2, op: { method: 'editCell', args: [0, 1, 1] } },
        { ...envelope, sequence: 2, op: { method: 'editCell', args: [0, , 1, 'invalid'] } },
        { ...envelope, sequence: 2, op: { method: 'editCells', args: [0, [{ row: 1, col: 1, input: 4 }]] } },
        { ...envelope, sequence: 2, op: { method: 'applyEdits', args: [{ expectVersion: '', steps: [null] }] } },
        { ...envelope, sequence: 2, op: { method: 'patchRangeStyle', args: [0, 'B3', { bold: 1 }] } },
        { ...envelope, sequence: 2, op: { method: 'undo', args: [0] } },
        { ...envelope, sequence: 2, op: { method: 'applyOps', args: [[{ op: 'insertRows' }]] } },
        { ...envelope, sequence: 2, op: { method: 'applyOps', args: [[{ type: 42 }]] } },
        { ...envelope, sequence: 2, op: { method: 'applyOps', args: [[new Date(0)]] } },
        { ...envelope, sequence: 2, op: { method: 'applyFormat', args: [0, 'B3', {
          rows: 1, columns: 1, formats: [new Map()],
        }] } },
      ];
      for (const entry of [null, 42, {}]) {
        const operations = [
          { method: 'editCells', args: [0, [entry]] },
          { method: 'editCells', args: [0, entry] },
          { method: 'applyOps', args: [[entry]] },
          { method: 'applyOps', args: [entry] },
          { method: 'applyEdits', args: [{ expectVersion: '', steps: [entry] }] },
          { method: 'applyEdits', args: [entry] },
          { method: 'applyFormat', args: [0, 'B3', { rows: 1, columns: 1, formats: [entry] }] },
          { method: 'applyFormat', args: [0, 'B3', entry] },
          { method: 'patchRangeStyle', args: [0, 'B3', [entry]] },
          { method: 'patchRangeStyle', args: [0, 'B3', { clear: [entry] }] },
          { method: 'setNumberFormat', args: [0, 'B3', entry] },
          { method: 'setNumberFormat', args: [0, 'B3', [entry]] },
          { method: 'applyEdits', args: [{ expectVersion: '', steps: [{
            op: 'setCellInputs', target: target('B3'), inputs: [[entry]],
          }] }] },
          { method: 'applyEdits', args: [{ expectVersion: '', steps: [{
            op: 'setFormulas', target: target('B3'), formulas: [[entry]],
          }] }] },
          { method: 'applyEdits', args: [{ expectVersion: '', steps: [{
            op: 'setCellInputs', target: entry, inputs: [['invalid']],
          }] }] },
        ];
        malformedEnvelopes.push(...operations.map((op) => ({ ...envelope, sequence: 2, op })));
      }
      for (const malformed of malformedEnvelopes) {
        await expect(replay(session, malformed as WorkbookReplayEnvelope)).rejects.toMatchObject({
          name: 'WorkbookReplayValidationError',
        });
        expect(await digest(await session.save())).toBe(initial);
        expect(await session.call.version()).toBe(initialVersion);
        expect((await session.call.frame(viewport)).sequence).toBe(1);
        expect(session.state).toMatchObject({ version: 0, dirty: false, stage: 'ready' });
        expect(session.failure).toBeUndefined();
        expect(await digest(peer.save())).toBe(initial);
      }
      const accepted = request(peer, 'B3', '777');
      expect(peer.applyEdits(accepted)).toMatchObject({ ok: true, applied: true });
      const changed = await replay(session, {
        sequence: 2, calculation, op: { method: 'applyEdits', args: [accepted] },
      });
      expect(changed).toMatchObject({ sequence: 2, revision: 1, version: 1, result: { ok: true, applied: true } });
      expect(await digest(await session.save())).toBe(await digest(peer.save()));
      expect((await session.call.frame(viewport)).sequence).toBe(2);
      expect(session.failure).toBeUndefined();
      expect(await replay(session, {
        sequence: 3, calculation, op: { method: 'applyEdits', args: [request(peer, 'B3', '777')] },
      })).toMatchObject({ sequence: 3, revision: 1, version: 1, result: { ok: true, applied: false } });
      expect((await session.call.frame(viewport)).sequence).toBe(3);
    } finally {
      peer.dispose();
      await session.dispose();
    }
  });

  test('fails terminally on worker refusal and preserves all accepted peer edits for one recovery save', async () => {
    const peer = openWorkbook(fixture, { calculation });
    const session = await createTestWorkbookSession(fixture, (transport) => {
      const batched = batchRequests(transport);
      return { ...batched, listen: (listener) => batched.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          const envelope = message.args[0] as WorkbookReplayEnvelope;
          if (envelope.op.method === 'applyEdits') {
            const input = envelope.op.args[0];
            envelope.op.args = [{ ...input, steps: [{
              op: 'setCellInputs', target: { ...target('B3'), sheetId: 'missing-sheet' }, inputs: [['saved edit']],
            }] }];
          }
        }
        listener(message);
      }) };
    }, { calculation });
    const errors: Error[] = [];
    const edits = createWorkbookEditPeer({
      session, peer, ...deterministicOptions(), onError: (error) => { errors.push(error); },
    });
    try {
      expect(edits.applyEdits(request(peer, 'B3', 'saved edit'))).toMatchObject({ ok: true, applied: true });
      expect(edits.editCell(0, 2, 2, 'queued edit').applied).toBe(true);
      await Promise.resolve();
      const pending = Promise.allSettled([session.call.readCells({ ranges: [] }), session.save()]);
      await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      for (const result of await pending) {
        if (result.status !== 'rejected') throw new Error('Pending work survived replay divergence');
        expect(result.reason).toBeInstanceOf(SessionFailure);
      }
      expect(edits.state).toBe('failed');
      expect(edits.sentSequence).toBe(2);
      expect(edits.acknowledgedSequence).toBe(0);
      expect(edits.error).toBe(session.failure);
      expect(session.state.stage).toBe('failed');
      expect(session.failure?.diagnostics).toContain('sequence 1 (applyEdits)');
      expect(session.failure?.diagnostics).toContain('missing-target');
      const error = edits.error;
      if (!error) throw new Error('Missing workbook edit peer failure');
      expect(errors).toEqual([error]);
      expect(() => edits.editCell(0, 2, 1, 'lost edit')).toThrow(WorkbookEditPeerFailedError);
      expect(() => edits.undo()).toThrow(WorkbookEditPeerFailedError);
      await expect(edits.save()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      const recovery = edits.recoverySave();
      expect(recovery.recovery).toBe(true);
      expect(await digest(recovery.bytes)).toBe(await digest(peer.save()));
      const reopened = openWorkbook(new Uint8Array(recovery.bytes), { calculation });
      try {
        expect(reopened.cell(0, 2, 1).input).toBe('saved edit');
        expect(reopened.cell(0, 2, 2).input).toBe('queued edit');
      } finally { reopened.dispose(); }
      expect(() => edits.recoverySave()).toThrow('already saved');
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('fails the peer on a nonterminal replay rejection while the worker remains usable', async () => {
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          (message.args[0] as WorkbookReplayEnvelope).sequence += 1;
        }
        listener(message);
      }),
    }), { calculation });
    const peer = openWorkbook(fixture, { calculation });
    const errors: Error[] = [];
    const edits = createWorkbookEditPeer({ session, peer, onError: (error) => { errors.push(error); } });
    try {
      const unchanged = await digest(await session.save());
      edits.editCell(0, 2, 1, 'recover rejected replay');
      await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(edits.error?.name).toBe('WorkbookReplayOrderError');
      expect(edits.state).toBe('failed');
      const error = edits.error;
      if (!error) throw new Error('Missing workbook edit peer failure');
      expect(errors).toEqual([error]);
      expect(session.failure).toBeUndefined();
      expect(session.state.stage).toBe('ready');
      expect(await digest(await session.save())).toBe(unchanged);
      expect(() => edits.editCells(0, [])).toThrow(WorkbookEditPeerFailedError);
      await expect(edits.save()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
      expect(await digest(edits.recoverySave().bytes)).toBe(await digest(peer.save()));
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('fails on a worker mutator throw and surfaces unrelated session failures', async () => {
    for (const mode of ['throw', 'crash'] as const) {
      let crash: ((error: unknown) => void) | undefined;
      const wrapHost = (transport: SessionTransport): SessionTransport => ({
        ...transport,
        onError(listener) { crash = listener; return transport.onError(listener); },
        listen: (listener) => transport.listen((message) => {
          if (mode === 'throw' && isClientMessage(message) && message.kind === 'call' &&
            message.method === 'replay') {
            const envelope = message.args[0] as WorkbookReplayEnvelope;
            envelope.op = { method: 'moveChart', args: [0, 'missing-chart', 1, 0] };
          }
          listener(message);
        }),
      });
      const session = await createTestWorkbookSession(fixture, wrapHost, { calculation });
      const peer = openWorkbook(fixture, { calculation });
      const errors: Error[] = [];
      const edits = createWorkbookEditPeer({ session, peer, onError: (error) => { errors.push(error); } });
      try {
        edits.editCell(0, 2, 1, 'recover me');
        if (mode === 'crash') crash!(new SessionFailure('crash', 'Worker stopped'));
        await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
        expect(edits.state).toBe('failed');
        const failure = session.failure;
        if (!failure) throw new Error('Missing workbook session failure');
        expect(errors).toEqual([failure]);
        expect(() => edits.redo()).toThrow(WorkbookEditPeerFailedError);
        if (mode === 'throw') expect(session.failure?.diagnostics).toContain('sequence 1 (moveChart)');
        expect(await digest(edits.recoverySave().bytes)).toBe(await digest(peer.save()));
      } finally {
        edits.dispose();
        peer.dispose();
        await session.dispose();
      }
    }
  });
});
