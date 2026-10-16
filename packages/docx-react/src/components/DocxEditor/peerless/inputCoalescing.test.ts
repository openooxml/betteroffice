import { expect, test } from 'bun:test';
import type {
  DocxCommandFramePayload, DocxCommandPayload, DocxCommandSelection, DocxEditState,
} from '@betteroffice/docx/peerless';
import { createOutbox } from '../../../../../../shared/office-session/commands/outbox';
import { createWorkerEditSession } from '../../../../../../shared/office-session/commands/session';
import type {
  CommandAck, CommandConnection, CommandEnvelope, CommandTicket,
} from '../../../../../../shared/office-session/commands/types';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import { admitIntent, createInputAdmission, submitFormattingIntent } from './inputAdmission';
import { coalesceInputDrafts, type PeerlessCommandDraft } from './inputCoalescing';
import { createPeerlessViewStore, type PeerlessViewSnapshot } from './viewStore';

function ticket(id: number): CommandTicket {
  return { id, seq: undefined, ack: Promise.resolve({ status: 'noop' }), presented: Promise.resolve('presented') };
}

function draft(text: string, options: PeerlessCommandDraft['options'] = {}): PeerlessCommandDraft {
  const intent = { kind: 'insertText' as const, payload: { text } };
  return { ...intent, options: {
    coalesceKey: 'typing', ...options,
    echo: { intent, originalIntentCount: 1, contextFrameEpoch: 3 },
  } };
}

test('d2b_coalescing_preserves_aliases_and_fences', () => {
  const after = ticket(90);
  const a = draft('🦊', { after, group: 'text' });
  const b = draft('é', { after, group: 'text' });
  const merged = coalesceInputDrafts(a, b)!;
  expect(merged.payload).toEqual({ text: '🦊é' });
  expect(merged.options.after).toBe(after);
  expect(merged.options.echo?.originalIntentCount).toBe(2);
  const outbox = createOutbox({ coalesce: coalesceInputDrafts });
  const tickets = [ticket(1), ticket(2)];
  outbox.enqueue(a, tickets[0]!);
  const entry = outbox.enqueue(b, tickets[1]!);
  expect(entry.tickets).toEqual(tickets);
  expect(outbox.size).toBe(1);
  for (const options of [
    { after: ticket(91) }, { group: 'other' }, { fence: true },
    { lane: 'background' as const }, { expectVersion: true }, { coalesceKey: 'paste' },
  ]) expect(coalesceInputDrafts(a, draft('b', { after, group: 'text', ...options }))).toBeUndefined();
  for (const next of [
    { kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } },
    { kind: 'toggleMark', payload: { mark: { type: 'bold' } } },
    { kind: 'replaceSelection', payload: { text: 'paste' } },
  ] satisfies Array<Pick<PeerlessCommandDraft, 'kind' | 'payload'>>) {
    expect(coalesceInputDrafts(a, { ...next, options: b.options })).toBeUndefined();
  }
  expect(coalesceInputDrafts(draft('a', { fence: true }), draft('b'))).toBeUndefined();
});

function typingSession() {
  type Ack = CommandAck<DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload>;
  const calls: { envelope: CommandEnvelope<DocxCommandPayload>; ack: Promise<Ack>; resolve(value: Ack): void }[] = [];
  const queued: (() => void)[] = [];
  const connection: CommandConnection<DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload> = {
    initial: { session: 'test', incarnation: 1, epoch: 1, version: 0, throughSeq: 0, outcomes: [] },
    client: {
      call: {
        edit(envelope) {
          let resolve!: (value: Ack) => void;
          const ack = new Promise<Ack>((settle) => { resolve = settle; });
          calls.push({ envelope, ack, resolve });
          return ack;
        },
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
      },
      failure: undefined, onFailure: () => () => {}, dispose: async () => {},
    },
  };
  const session = createWorkerEditSession<string, DocxCommandPayload, DocxCommandSelection | null, DocxEditState,
    DocxCommandFramePayload, PeerlessEchoIntent>({
    source: 'test', connection, recovery: { reopen: async () => { throw new Error('unexpected recovery'); } },
    maxRecoveryAttempts: 0, checkpointPolicy: () => false, coalesce: coalesceInputDrafts,
    requestFullFrame: async () => connection.initial, now: () => 0,
    enqueue(callback) { queued.push(callback); },
    echo: { paint() {}, drop() {}, present: () => true },
  });
  const snapshot: PeerlessViewSnapshot = {
    identity: { session: 'test', incarnation: 1, epoch: 1 }, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, geometry: null, editState: undefined,
  };
  const viewStore = createPeerlessViewStore(snapshot);
  const echoPainter: PeerlessEchoPainter = {
    installContext() {}, restoreBackground: () => true,
    preview: (intent) => ({ intent, originalIntentCount: 1, contextFrameEpoch: 1 }),
    paint() {}, drop() {}, reconcile() {}, disable() {}, dispose() {},
  };
  const admission = createInputAdmission({ session, viewStore, echoPainter });
  const pump = () => { while (queued.length) queued.shift()!(); };
  let acknowledged = 0;
  return { session, admission, calls, pump,
    installGeometry(coversThroughSeq: number) {
      viewStore.installFrameGeometry({
        identity: snapshot.identity, coversThroughSeq, frameEpoch: coversThroughSeq, layoutRevision: 1,
        pages: [], caret: { x: 0, y: 0, width: 1, height: 16 }, selectionRects: [],
      });
    },
    type(text: string, onSubmitted?: (ticket: CommandTicket) => void) {
      return [...text].map((char) => {
        const result = admitIntent(admission, { kind: 'insertText', payload: { text: char } }, { onSubmitted });
        pump();
        if (result.status !== 'submitted') throw new Error(`Expected submitted input, got ${result.status}`);
        return result.ticket;
      });
    },
    async acknowledge(count: number) {
      const batch = calls.slice(acknowledged, acknowledged + count);
      expect(batch).toHaveLength(count);
      acknowledged += count;
      for (const { envelope, resolve } of batch) {
        resolve({ ...connection.initial, version: envelope.seq, throughSeq: envelope.seq,
          outcomes: [{ seq: envelope.seq, status: 'applied' }] });
      }
      await Promise.all(batch.map(({ ack }) => ack));
      pump();
    },
  };
}

test('chained typing behind held acks sends one coalesced insertText per in-flight window', async () => {
  const h = typingSession();
  const typed = h.type('abcdefghijklmnop');
  const order: number[] = [];
  typed.forEach((ticket, index) => { void ticket.ack.then(() => { order.push(index); }); });
  expect(h.calls.map(({ envelope }) => envelope.payload)).toEqual([...'abcd'].map((text) => ({ text })));
  await h.acknowledge(1);
  expect(h.calls).toHaveLength(5);
  expect(h.calls[4]!.envelope).toMatchObject({ kind: 'insertText', payload: { text: 'efghijklmnop' }, afterSeq: 4 });
  const later = h.type('qrstu');
  expect(h.calls).toHaveLength(5);
  await h.acknowledge(1);
  expect(h.calls).toHaveLength(6);
  expect(h.calls[5]!.envelope).toMatchObject({ kind: 'insertText', payload: { text: 'qrstu' }, afterSeq: 5 });
  await h.acknowledge(4);
  const outcomes = await Promise.all([...typed, ...later].map((ticket) => ticket.ack));
  expect(outcomes.map((outcome) => outcome.seq)).toEqual([1, 2, 3, 4, ...Array(12).fill(5), ...Array(5).fill(6)]);
  expect(outcomes.every((outcome) => outcome.status === 'applied')).toBe(true);
  expect(order).toEqual(typed.map((_, index) => index));
  h.admission.dispose();
  await h.session.dispose();
});

test('coalesced typing carries the input times of its first and last keystrokes', async () => {
  const h = typingSession();
  [...'abcdefg'].forEach((text, index) => {
    admitIntent(h.admission, { kind: 'insertText', payload: { text } }, { inputTime: 100 + index * 150 });
    h.pump();
  });
  expect(h.calls.map(({ envelope }) => [envelope.inputTime, envelope.lastInputTime]))
    .toEqual([[100, undefined], [250, undefined], [400, undefined], [550, undefined]]);
  await h.acknowledge(1);
  expect(h.calls[4]!.envelope).toMatchObject({ payload: { text: 'efg' }, inputTime: 700, lastInputTime: 1000 });
});

test('chained typing never coalesces across a selection or format change', async () => {
  const h = typingSession();
  h.type('abcdef');
  admitIntent(h.admission, { kind: 'navigate', payload: { unit: 'character', direction: 'backward', extend: false } });
  h.pump();
  h.type('gh');
  submitFormattingIntent(h.admission, { kind: 'toggleMark', payload: { mark: { type: 'bold' } } });
  h.pump();
  h.type('ij');
  await h.acknowledge(4);
  await h.acknowledge(4);
  expect(h.calls.slice(4).map(({ envelope }) => [envelope.kind, 'text' in envelope.payload ? envelope.payload.text : null]))
    .toEqual([['insertText', 'ef'], ['navigate', null], ['insertText', 'gh'], ['toggleMark', null], ['insertText', 'ij']]);
  h.admission.dispose();
  await h.session.dispose();
});

test.each([
  { name: 'cancelled empty', text: null },
  { name: 'committed empty', text: '' },
])('deferred typing stays in separate transactions across a $name composition', async ({ text }) => {
  const h = typingSession();
  try {
    h.type('wxyz');
    expect(h.calls).toHaveLength(4);
    const navigation = admitIntent(h.admission, {
      kind: 'navigate', payload: { unit: 'line', direction: 'down', extend: false, goalX: 0 },
    }, { onSubmitted: h.pump });
    expect(navigation.status).toBe('submitted');
    const before = h.type('abcde', h.pump);
    const reservation = h.admission.reserveComposition();
    const composition = h.admission.settleComposition(reservation.slot,
      text === null ? null : { text, target: reservation.target });
    expect(composition.status).toBe('ignored');
    let submitted!: () => void;
    const whenSubmitted = new Promise<void>((resolve) => { submitted = resolve; });
    const [f] = h.type('f', () => { h.pump(); submitted(); });
    expect(h.calls).toHaveLength(4);
    expect([...before, f!].every((ticket) => ticket.seq === undefined)).toBe(true);

    await h.acknowledge(4);
    expect(h.calls).toHaveLength(4);
    h.installGeometry(4);
    await whenSubmitted;
    expect(h.calls.slice(4).map(({ envelope }) =>
      [envelope.kind, 'text' in envelope.payload ? envelope.payload.text : null]))
      .toEqual([['navigate', null], ['insertText', 'a'], ['insertText', 'b'], ['insertText', 'c']]);
    await h.acknowledge(1);
    expect(h.calls[8]!.envelope).toMatchObject({ kind: 'insertText', payload: { text: 'de' }, afterSeq: 8 });
    await h.acknowledge(1);
    expect(h.calls[9]!.envelope).toMatchObject({ kind: 'insertText', payload: { text: 'f' }, afterSeq: 9 });
    expect(before.at(-1)!.seq).not.toBe(f!.seq);
    await h.acknowledge(4);
    const outcomes = await Promise.all([...before, f!].map((ticket) => ticket.ack));
    expect(outcomes.map((outcome) => outcome.seq)).toEqual([6, 7, 8, 9, 9, 10]);
    expect(outcomes.every((outcome) => outcome.status === 'applied')).toBe(true);
  } finally {
    h.admission.dispose();
    await h.session.dispose();
  }
});

test.each([
  { name: 'cancelled empty', text: null },
  { name: 'committed empty', text: '' },
  { name: 'committed nonempty', text: '文' },
])('chained typing stays in separate transactions across a $name composition with echo disabled', async ({ text }) => {
  const h = typingSession();
  try {
    h.type('wxyz');
    expect(h.calls).toHaveLength(4);
    const navigation = admitIntent(h.admission, {
      kind: 'navigate', payload: { unit: 'character', direction: 'backward', extend: false },
    });
    h.pump();
    expect(navigation.status).toBe('submitted');
    expect(h.admission.getSnapshot().echoEnabled).toBe(false);
    const [a] = h.type('a');
    const reservation = h.admission.reserveComposition();
    const composition = h.admission.settleComposition(reservation.slot,
      text === null ? null : { text, target: reservation.target });
    expect(composition.status).toBe(text ? 'submitted' : 'ignored');
    h.pump();
    const [b] = h.type('b');
    expect(h.calls).toHaveLength(4);

    await h.acknowledge(4);
    const expected = [
      ['navigate', null], ['insertText', 'a'],
      ...(text ? [['replaceSelection', text]] : []),
      ['insertText', 'b'],
    ];
    expect(h.calls.slice(4).map(({ envelope }) =>
      [envelope.kind, 'text' in envelope.payload ? envelope.payload.text : null])).toEqual(expected);
    expect(a!.seq).not.toBe(b!.seq);
    expect(h.calls.slice(4).map(({ envelope }) => envelope.afterSeq))
      .toEqual(expected.map((_, index) => index + 4));
    await h.acknowledge(expected.length);
    const outcomes = await Promise.all([a!.ack, b!.ack]);
    expect(outcomes.every((outcome) => outcome.status === 'applied')).toBe(true);
    expect(outcomes[0]!.seq).not.toBe(outcomes[1]!.seq);
  } finally {
    h.admission.dispose();
    await h.session.dispose();
  }
});
