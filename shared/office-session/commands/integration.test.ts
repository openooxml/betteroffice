import { describe, expect, it, spyOn } from 'bun:test';
import { createSessionClient } from '../client';
import { createSessionHost } from '../host';
import { isClientMessage, isHostMessage, transferable } from '../protocol';
import { createInProcessPair } from '../testing/inProcessTransport';
import type { SessionTransport } from '../transport';
import { SessionFailure } from '../types';
import { createCommandHost } from './host';
import { createWorkerEditSession } from './session';
import {
  isCommandAck, type Checkpoint, type CommandAck, type CommandConnection,
  type CommandEnvelope, type CommandMethods, type EchoHooks, type FrameEnvelope,
  type RecoveryEvent, type RecoveryHooks,
} from './types';

type Payload = { text: string };
type Frame = { text: string; pixels: ArrayBuffer };
type Methods = CommandMethods<Payload, number, boolean, Frame>;
type Ack = CommandAck<number, boolean, Frame>;
type Connection = CommandConnection<Payload, number, boolean, Frame>;
type Recovery = RecoveryHooks<string, Payload, number, boolean, Frame>;

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function fakeClock() {
  const tasks = new Map<number, { at: number; run(): void }>();
  let nextId = 0;
  const now = 0;
  const date = spyOn(Date, 'now').mockImplementation(() => now);
  const schedule = (
    ...args: Parameters<typeof setTimeout>
  ): ReturnType<typeof setTimeout> => {
    const [callback, delay, ...values] = args;
    if (typeof callback !== 'function') throw new Error('Expected a timer callback');
    const id = ++nextId;
    tasks.set(id, { at: now + (delay ?? 0), run: () => { callback(...values); } });
    return id as unknown as ReturnType<typeof setTimeout>;
  };
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(schedule as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((id) => {
    tasks.delete(id as unknown as number);
  });
  async function microtasks(): Promise<void> {
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
  }
  function ready(): boolean {
    return [...tasks.values()].some((task) => task.at <= now);
  }
  async function turn(): Promise<void> {
    for (const [id, task] of [...tasks]) {
      if (task.at > now || !tasks.has(id)) continue;
      tasks.delete(id);
      task.run();
    }
    await microtasks();
  }
  async function until(predicate: () => boolean): Promise<void> {
    for (let index = 0; index < 200; index += 1) {
      await turn();
      if (predicate()) return;
    }
    throw new Error('Controlled event queue did not reach the expected state');
  }
  return {
    now: () => now,
    enqueue: (callback: () => void) => { setTimeout(callback, 0); },
    async drain(): Promise<void> {
      await microtasks();
      for (let index = 0; index < 200; index += 1) {
        await turn();
        if (!ready()) return;
      }
      throw new Error('Controlled event queue did not drain');
    },
    until,
    async settle<T>(promise: Promise<T>): Promise<T> {
      let settled = false;
      const observed = promise.then(
        (value) => { settled = true; return value; },
        (error: unknown) => { settled = true; throw error; }
      );
      void observed.catch(() => {});
      await until(() => settled);
      return await observed;
    },
    restore() {
      tasks.clear();
      clear.mockRestore();
      timeout.mockRestore();
      date.mockRestore();
    },
  };
}

function openWorker(input: {
  incarnation?: number;
  checkpoint?: Checkpoint;
  source?: string;
  refusedSeqs?: readonly number[];
  replayAppliedSeqs?: readonly number[];
  applyGate?: Promise<void>;
  frameGates?: ReadonlyMap<number, Promise<void>>;
} = {}) {
  const pair = createInProcessPair();
  const faults = new Set<(failure: unknown) => void>();
  const incoming: CommandEnvelope<Payload>[] = [];
  const acks: Ack[] = [];
  const applied: Array<{ seq: number; kind: string; replay: boolean }> = [];
  const frameBuffers: ArrayBuffer[] = [];
  const framed: number[] = [];
  const restored: { text: string; version: number; order: string[] } = input.checkpoint
    ? JSON.parse(new TextDecoder().decode(input.checkpoint.state))
    : { text: input.source ?? '', version: 0, order: [] };
  let text = restored.text;
  let version = restored.version;
  const order = [...restored.order];
  let dropReplies = false;
  function endpoint(transport: SessionTransport, host = false): SessionTransport {
    return {
      post(message, transfer) {
        if (host && dropReplies && isHostMessage(message) && message.kind === 'reply') return;
        transport.post(message, transfer);
      },
      listen: (listener) => transport.listen(listener),
      onError(listener) {
        faults.add(listener);
        const off = transport.onError(listener);
        return () => { faults.delete(listener); off(); };
      },
      close: () => transport.close(),
    };
  }
  const commands = createCommandHost<Payload, number, boolean, Frame>({
    session: 'integration', incarnation: input.incarnation ?? 0,
    epoch: input.checkpoint?.epoch ?? 1, throughSeq: input.checkpoint?.throughSeq ?? 0,
    refusedSeqs: input.refusedSeqs, replayAppliedSeqs: input.replayAppliedSeqs,
    getVersion: () => version,
    async apply(envelope) {
      if (input.applyGate) await input.applyGate;
      text += envelope.payload.text;
      version += 1;
      order.push(envelope.kind);
      applied.push({ seq: envelope.seq, kind: envelope.kind, replay: envelope.replay === true });
      return { status: 'applied', selection: text.length, editState: true };
    },
    async frame({ coversThroughSeq }) {
      framed.push(coversThroughSeq);
      const gate = input.frameGates?.get(coversThroughSeq);
      if (gate) await gate;
      const pixels = new TextEncoder().encode(text).buffer;
      frameBuffers.push(pixels);
      const frame: FrameEnvelope<Frame> = {
        coversThroughSeq, revision: version, payload: { text, pixels },
      };
      return transferable(frame, [pixels]);
    },
    snapshot: () => new TextEncoder().encode(JSON.stringify({ text, version, order })),
  });
  const host = createSessionHost<Methods, {}, null>(endpoint(pair.host, true), {
    context: null,
    handlers: { edit: commands.handler, checkpoint: commands.checkpoint },
    policies: { edit: commands.policy, checkpoint: commands.checkpointPolicy },
  });
  pair.host.listen((message) => {
    if (isClientMessage(message) && message.kind === 'call' && message.method === 'edit') {
      incoming.push(message.args[0] as CommandEnvelope<Payload>);
    }
  });
  pair.client.listen((message) => {
    if (isHostMessage(message) && message.kind === 'reply' && message.ok && isCommandAck(message.value)) {
      acks.push(message.value as Ack);
    }
  });
  const client = createSessionClient<Methods, {}>(endpoint(pair.client), {
    methods: { edit: true, checkpoint: true }, now: () => 0,
    timer(callback, ms) {
      const id = setTimeout(callback, ms);
      return () => { clearTimeout(id); };
    },
  });
  const initial: Ack = {
    session: 'integration', incarnation: input.incarnation ?? 0,
    epoch: input.checkpoint?.epoch ?? 1, version,
    throughSeq: input.checkpoint?.throughSeq ?? 0, outcomes: [],
    selection: text.length, editState: order.length > 0,
    frame: { coversThroughSeq: input.checkpoint?.throughSeq ?? 0, revision: version,
      payload: { text, pixels: new TextEncoder().encode(text).buffer } },
  };
  const connection: Connection = { initial, client };
  return {
    connection, host, incoming, acks, applied, frameBuffers, framed, order,
    get text() { return text; },
    set dropReplies(value: boolean) { dropReplies = value; },
    crash() {
      const failure = new SessionFailure('crash', 'Simulated worker crash');
      for (const listener of [...faults]) listener(failure);
    },
  };
}

function unexpectedReopen(): Promise<Connection> {
  return Promise.reject(new Error('Unexpected recovery'));
}

function unexpectedFrameRequest(): Promise<Ack> {
  return Promise.reject(new Error('Unexpected full frame request'));
}

describe('command layer over the real RPC and resident scheduler', () => {
  it('preserves pipelined edit arrival order while the first apply is blocked', async () => {
    const clock = fakeClock();
    try {
      const gate = deferred<void>();
      const worker = openWorker({ applyGate: gate.promise });
      const session = createWorkerEditSession<string, Payload, number, boolean, Frame, string>({
        source: '', connection: worker.connection, recovery: { reopen: unexpectedReopen },
        requestFullFrame: unexpectedFrameRequest, enqueue: clock.enqueue, now: clock.now,
      });
      const tickets = ['a', 'b', 'c'].map((text) => session.submit('insert', { text }));
      await clock.drain();
      expect(worker.incoming.map((envelope) => envelope.seq)).toEqual([1, 2, 3]);
      expect(worker.incoming.map((envelope) => envelope.basis.version)).toEqual([0, 0, 0]);
      expect(worker.applied).toEqual([]);
      expect(worker.host.scheduler.pending().foreground).toBe(2);
      gate.resolve();
      await clock.settle(session.flush());
      expect(worker.applied.map((entry) => entry.seq)).toEqual([1, 2, 3]);
      expect(worker.text).toBe('abc');
      expect(await Promise.all(tickets.map((ticket) => ticket.ack))).toEqual([
        { seq: 1, status: 'applied' }, { seq: 2, status: 'applied' }, { seq: 3, status: 'applied' },
      ]);
      expect(session.state.version).toBe(3);
      expect(worker.host.scheduler.version).toBe(3);
      await clock.settle(session.dispose());
    } finally { clock.restore(); }
  });

  it('omits intermediate frames for queued edits and transfers the final covering frame', async () => {
    const clock = fakeClock();
    try {
      const worker = openWorker();
      const session = createWorkerEditSession<string, Payload, number, boolean, Frame, string>({
        source: '', connection: worker.connection, recovery: { reopen: unexpectedReopen },
        requestFullFrame: unexpectedFrameRequest, enqueue: clock.enqueue, now: clock.now,
      });
      const tickets = ['a', 'b', 'c'].map((text) => session.submit('insert', { text }));
      await clock.settle(session.flush());
      expect(worker.acks.map((ack) => ack.throughSeq)).toEqual([1, 2, 3]);
      expect(worker.acks.map((ack) => ack.frame?.coversThroughSeq)).toEqual([undefined, undefined, 3]);
      expect(worker.framed).toEqual([3]);
      expect(worker.frameBuffers.map((buffer) => buffer.byteLength)).toEqual([0]);
      const frame = worker.acks[2]!.frame!;
      expect(frame.payload.text).toBe('abc');
      expect(new TextDecoder().decode(frame.payload.pixels)).toBe('abc');
      expect(frame.payload.pixels.byteLength).toBe(3);
      expect(await clock.settle(Promise.all(tickets.map((ticket) => ticket.presented))))
        .toEqual(['presented', 'presented', 'presented']);
      expect(session.state.presentedThroughSeq).toBe(3);
      await clock.settle(session.dispose());
    } finally { clock.restore(); }
  });

  it('keeps remote/local seq order through replay and redraws later local echo over remote coverage', async () => {
    const clock = fakeClock();
    try {
      const remoteFrame = deferred<void>();
      const localFrame = deferred<void>();
      const original = openWorker({ frameGates: new Map([
        [2, remoteFrame.promise], [3, localFrame.promise],
      ]) });
      const workers = [original];
      const events: RecoveryEvent[] = [];
      const presentations: Array<{ coverage: number; text: string; retired: number[];
        remaining: Array<{ seq: number | undefined; value: string }> }> = [];
      const echo: EchoHooks<string, Frame> = {
        paint: () => {}, drop: () => {},
        present(frame, retired, remaining) {
          presentations.push({ coverage: frame.coversThroughSeq, text: frame.payload.text,
            retired: retired.map((entry) => entry.seq ?? entry.ticket.seq!),
            remaining: remaining.map((entry) => ({ seq: entry.seq ?? entry.ticket.seq, value: entry.value })) });
        },
      };
      const recovery: Recovery = {
        async reopen(input) {
          const worker = openWorker(input);
          workers.push(worker);
          return worker.connection;
        },
        onRecovery: (event) => { events.push(event); },
      };
      const session = createWorkerEditSession<string, Payload, number, boolean, Frame, string>({
        source: '', connection: original.connection, recovery, echo, checkpointPolicy: () => false,
        requestFullFrame: unexpectedFrameRequest, enqueue: clock.enqueue, now: clock.now,
      });
      const first = session.submit('local-a', { text: 'a' }, { echo: 'a' });
      const remote = session.submit('remote', { text: 'R' }, { lane: 'collab', after: first });
      await clock.until(() => original.framed.includes(2));
      const later = session.submit('local-b', { text: 'b' }, { echo: 'b', after: remote });
      await clock.until(() => original.incoming.length === 3);
      remoteFrame.resolve();
      await clock.settle(remote.presented);
      expect(presentations.some((entry) => entry.coverage === 2 && entry.text === 'aR' &&
        entry.remaining.length === 1 && entry.remaining.some((pending) => pending.seq === 3 && pending.value === 'b')))
        .toBe(true);
      await clock.until(() => original.applied.length === 3);
      let laterAcked = false;
      void later.ack.then(() => { laterAcked = true; });
      await clock.drain();
      expect(laterAcked).toBe(false);
      original.crash();
      await clock.settle(later.ack);
      await clock.until(() => events.some((event) => event.phase === 'recovered'));
      const reopened = workers[1]!;
      expect(reopened.incoming.map((entry) => [entry.seq, entry.kind, entry.replay, entry.afterSeq]))
        .toEqual([[1, 'local-a', true, undefined], [2, 'remote', true, 1], [3, 'local-b', true, 2]]);
      expect(reopened.applied.map((entry) => entry.kind)).toEqual(['local-a', 'remote', 'local-b']);
      expect(reopened.text).toBe('aRb');
      expect(await clock.settle(later.presented)).toBe('presented');
      expect(presentations.some((entry) => entry.coverage === 3 && entry.text === 'aRb' &&
        entry.retired.includes(3) && entry.remaining.length === 0)).toBe(true);
      const next = session.submit('local-c', { text: 'c' }, { after: later });
      expect(await clock.settle(next.ack)).toEqual({ seq: 4, status: 'applied' });
      expect(reopened.text).toBe('aRbc');
      await clock.settle(session.dispose());
      localFrame.resolve();
      await clock.drain();
    } finally { clock.restore(); }
  });

  it('reopens a fresh host from a checkpoint and applies an unacked command exactly once', async () => {
    const clock = fakeClock();
    try {
      const original = openWorker();
      const workers = [original];
      const checkpoints: Checkpoint[] = [];
      const reopenedInputs: Parameters<Recovery['reopen']>[0][] = [];
      const events: RecoveryEvent[] = [];
      const recovery: Recovery = {
        async checkpoint(connection) {
          const checkpoint = await connection.client.call.checkpoint();
          checkpoints.push(structuredClone(checkpoint));
          return checkpoint;
        },
        async reopen(input) {
          reopenedInputs.push(input);
          const worker = openWorker(input);
          workers.push(worker);
          return worker.connection;
        },
        onRecovery: (event) => { events.push(event); },
      };
      const session = createWorkerEditSession<string, Payload, number, boolean, Frame, string>({
        source: '', connection: original.connection, recovery,
        checkpointPolicy: ({ applied }) => applied >= 2,
        requestFullFrame: unexpectedFrameRequest, enqueue: clock.enqueue, now: clock.now,
      });
      const settled: number[] = [];
      const first = session.submit('one', { text: '1' });
      const second = session.submit('two', { text: '2' });
      void first.ack.then(() => { settled.push(1); });
      void second.ack.then(() => { settled.push(2); });
      await clock.settle(session.flush());
      await clock.until(() => checkpoints.length === 1);
      await clock.drain();
      expect(checkpoints[0]?.throughSeq).toBe(2);
      expect(checkpoints[0]?.version).toBe(2);
      original.dropReplies = true;
      const third = session.submit('three', { text: '3' });
      let thirdAcked = false;
      void third.ack.then(() => { thirdAcked = true; settled.push(3); });
      await clock.until(() => original.applied.length === 3);
      await clock.drain();
      expect(original.text).toBe('123');
      expect(thirdAcked).toBe(false);
      original.crash();
      expect(original.connection.client.failure?.code).toBe('crash');
      expect(await clock.settle(third.ack)).toEqual({ seq: 3, status: 'applied' });
      await clock.until(() => events.some((event) => event.phase === 'recovered'));
      const reopened = workers[1]!;
      expect(reopenedInputs[0]?.checkpoint?.throughSeq).toBe(2);
      expect(reopenedInputs[0]?.incarnation).toBe(1);
      expect(reopened.incoming.map((entry) => [entry.seq, entry.replay])).toEqual([[3, true]]);
      expect(reopened.applied).toEqual([{ seq: 3, kind: 'three', replay: true }]);
      expect(reopened.order).toEqual(['one', 'two', 'three']);
      expect(reopened.text).toBe('123');
      expect(settled).toEqual([1, 2, 3]);
      expect(events.find((event) => event.phase === 'recovered')?.historyReset).toBe(true);
      const duplicate = { ...reopened.incoming[0]! };
      const duplicateAck = await clock.settle(reopened.connection.client.call.edit(duplicate));
      expect(duplicateAck.throughSeq).toBe(3);
      expect(duplicateAck.version).toBe(3);
      expect(reopened.applied).toHaveLength(1);
      expect(reopened.text).toBe('123');
      expect(await clock.settle(third.presented)).toBe('presented');
      const fourth = session.submit('four', { text: '4' });
      expect(await clock.settle(fourth.ack)).toEqual({ seq: 4, status: 'applied' });
      expect(reopened.order).toEqual(['one', 'two', 'three', 'four']);
      expect(reopened.text).toBe('1234');
      expect(settled).toEqual([1, 2, 3]);
      await clock.settle(session.dispose());
    } finally { clock.restore(); }
  });
});
