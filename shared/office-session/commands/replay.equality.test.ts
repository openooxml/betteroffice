import { expect, test } from 'bun:test';
import { freezeDeep } from '../testing/frozen';
import { SessionFailure } from '../types';
import { createWorkerEditSession } from './session';
import type { CommandAck, CommandConnection, CommandEnvelope } from './types';

type Payload = { text: string; bytes: Uint8Array; marks: { bold: boolean; spans: number[] } };
type Ack = CommandAck<string, string, string>;

function worker(incarnation: number) {
  const edits: CommandEnvelope<Payload>[] = [];
  const replies: Array<(ack: Ack) => void> = [];
  const listeners = new Set<(failure: SessionFailure) => void>();
  let failure: SessionFailure | undefined;
  const initial: Ack = { session: 'replay', incarnation, epoch: 1, version: 0, throughSeq: 0, outcomes: [],
    frame: { coversThroughSeq: 0, revision: 0, payload: 'open' } };
  const connection: CommandConnection<Payload, string, string, string> = {
    initial,
    client: {
      call: {
        edit(envelope) {
          edits.push(envelope);
          return new Promise<Ack>((resolve) => { replies.push(resolve); });
        },
        checkpoint: () => Promise.reject(new Error('Unexpected checkpoint')),
      },
      get failure() { return failure; },
      onFailure(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
      dispose: () => Promise.resolve(),
    },
  };
  return { connection, edits,
    ack(index: number) {
      const seq = edits[index]!.seq;
      replies[index]!({ ...initial, throughSeq: seq, version: index + 1, outcomes: [{ seq, status: 'applied' }],
        frame: undefined });
    },
    crash() {
      failure = new SessionFailure('crash', 'worker crashed');
      for (const listener of [...listeners]) listener(failure);
    } };
}

async function turns() {
  for (let index = 0; index < 16; index += 1) await Promise.resolve();
}

test('replays envelopes deep-equal to the frozen ones first sent', async () => {
  const first = worker(0);
  const reopened: ReturnType<typeof worker>[] = [];
  const session = createWorkerEditSession<string, Payload, string, string, string, string>({
    source: 'source', connection: first.connection,
    recovery: { reopen: async ({ incarnation }) => {
      const next = worker(incarnation);
      reopened.push(next);
      return next.connection;
    } },
    enqueue: (callback) => { queueMicrotask(callback); }, now: () => 0,
    requestFullFrame: (connection) => Promise.resolve(connection.initial),
  });
  try {
    const head = session.submit('insert', { text: 'a', bytes: Uint8Array.of(1, 2, 3), marks: { bold: true, spans: [0, 1] } },
      { inputTime: 5, group: 'typing' });
    const tail = session.submit('insert', { text: 'b', bytes: Uint8Array.of(4), marks: { bold: false, spans: [] } },
      { after: head });
    await turns();
    expect(first.edits.map((envelope) => envelope.seq)).toEqual([1, 2]);
    const sent = first.edits.map((envelope) => structuredClone(envelope));
    for (const envelope of first.edits) freezeDeep(envelope);
    first.crash();
    await turns();
    const next = reopened[0]!;
    expect(next.edits).toHaveLength(1);
    next.ack(0);
    await turns();
    expect(next.edits).toHaveLength(2);
    next.ack(1);
    expect(await head.ack).toMatchObject({ seq: 1, status: 'applied' });
    expect(await tail.ack).toMatchObject({ seq: 2, status: 'applied' });
    expect(sent[1]).toMatchObject({ afterSeq: 1 });
    expect(next.edits).toEqual(sent.map((envelope) => ({ ...envelope, incarnation: 1, replay: true })));
    expect(session.state.throughSeq).toBe(2);
  } finally { await session.dispose(); }
});
