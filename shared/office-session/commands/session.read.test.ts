import { describe, expect, it } from 'bun:test';
import { SessionFailure } from '../types';
import { createWorkerEditSession } from './session';
import { query, type At, type QueryFence, type QueryReply } from './spec';
import type { CommandAck, CommandConnection, CommandEnvelope, WorkerEditSessionOptions } from './types';

const queries = {
  text: query<[], string>()({ stale: 'answer-at-fence' }),
  doc: query<[], string>()({ stale: { reject: ['doc'] } }),
  selection: query<[], string>()({ stale: { reject: ['selection'] } }),
  layout: query<[], string>()({ stale: { reject: ['layout'] } }),
};
type Ack = CommandAck<number, boolean, never>;
type Options = WorkerEditSessionOptions<string, string, number, boolean, never, never, typeof queries>;
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}
async function drain() {
  for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
}
function port(incarnation = 0) {
  const edits: Array<{ envelope: CommandEnvelope<string>; reply: ReturnType<typeof deferred<Ack>> }> = [];
  const reads: Array<{ fence: QueryFence; kind: string; args: unknown[]; reply: ReturnType<typeof deferred<QueryReply>> }> = [];
  const failures = new Set<(failure: SessionFailure) => void>();
  const acks = new Set<(ack: Ack) => void>();
  const connection: CommandConnection<string, number, boolean, never> = {
    initial: { session: 'reads', incarnation, epoch: 1, throughSeq: 0, version: 0,
      outcomes: [], revisions: { selection: 0, layout: 0 } },
    client: {
      call: {
        edit(envelope) { const reply = deferred<Ack>(); edits.push({ envelope, reply }); return reply.promise; },
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, throughSeq: 0, version: 0 }),
        query(fence, kind, args) { const reply = deferred<QueryReply>(); reads.push({ fence, kind, args, reply }); return reply.promise; },
      }, failure: undefined,
      onFailure(listener) { failures.add(listener); return () => { failures.delete(listener); }; },
      onAck(listener) { acks.add(listener); return () => { acks.delete(listener); }; },
      dispose: async () => {},
    },
  };
  return { connection, edits, reads, acks,
    crash(failure = new SessionFailure('crash', 'crashed')) { for (const listener of failures) listener(failure); },
    ack(index: number, extra: Partial<Ack> = {}) {
      const seq = edits[index]!.envelope.seq;
      edits[index]!.reply.resolve({ ...connection.initial, throughSeq: seq, version: seq,
        outcomes: [{ seq, status: 'applied' }], ...extra });
    },
    emit(extra: Partial<Ack>) { for (const listener of acks) listener({ ...connection.initial, ...extra }); },
    answer(index: number, at: Partial<At> = {}, text = 'answer') {
      const fence = reads[index]!.fence;
      reads[index]!.reply.resolve({ ok: true, value: text, at: { incarnation: fence.incarnation,
        epoch: fence.epoch, throughSeq: fence.afterSeq, doc: fence.afterSeq, selection: 0, layout: 0, ...at } });
    },
  };
}
function harness(extra: Partial<Options> = {}) {
  const worker = port();
  const queue: Array<() => void> = [];
  const timers = new Map<number, { at: number; run(): void }>();
  let id = 0;
  let now = 0;
  const session = createWorkerEditSession({
    source: '', connection: worker.connection, queries,
    recovery: { reopen: async () => { throw new Error('unexpected recovery'); } },
    maxRecoveryAttempts: 0, requestFullFrame: async () => { throw new Error('unexpected frame'); },
    enqueue(callback) { queue.push(callback); }, now: () => now,
    timer(callback, ms) { const key = ++id; timers.set(key, { at: now + ms, run: callback }); return () => { timers.delete(key); }; },
    ...extra,
  });
  return { session, worker, timers,
    pump() { while (queue.length) queue.shift()!(); },
    elapse(ms: number) { now += ms; },
    tick(ms: number) {
      now += ms;
      for (const [key, task] of timers) if (task.at <= now) { timers.delete(key); task.run(); }
    },
  };
}

describe('fenced session reads and holds', () => {
  it('fences every earlier ticket after send without waiting for its ack', async () => {
    const h = harness();
    const ticket = h.session.submit('insert', 'a');
    const read = h.session.read('text', []);
    expect(h.worker.reads).toHaveLength(0);
    h.pump();
    expect(h.worker.reads[0]!.fence.afterSeq).toBe(ticket.seq!);
    h.worker.answer(0);
    expect(await read).toMatchObject({ ok: true, at: { throughSeq: 1 } });
    await h.session.dispose();
  });

  it('waits for captured holds and includes the holder commit while ignoring later holds', async () => {
    const h = harness();
    const owner = h.session.hold('draft');
    const read = h.session.read('text', []);
    const later = h.session.hold('gesture');
    const ticket = h.session.submit('insert', 'commit');
    h.pump();
    expect(h.worker.reads).toHaveLength(0);
    owner.release();
    expect(h.worker.reads[0]!.fence.afterSeq).toBe(ticket.seq!);
    h.worker.answer(0, {}, 'commit');
    expect(await read).toMatchObject({ ok: true, value: 'commit' });
    expect(later.released).toBe(false);
    owner.release();
    await h.session.dispose();
    expect(later.released).toBe(true);
  });

  it('lets within bypass only its owner hold and keeps submit and flush live', async () => {
    const h = harness();
    const owner = h.session.hold('composition');
    const other = h.session.hold('mode');
    const read = h.session.read('text', [], { within: owner });
    const ticket = h.session.submit('insert', 'a');
    const flush = h.session.flush();
    h.pump(); h.worker.ack(0);
    expect(await ticket.ack).toMatchObject({ status: 'applied' });
    await flush;
    expect(h.worker.reads).toHaveLength(0);
    other.release();
    h.worker.answer(0);
    expect(await read).toMatchObject({ ok: true });
    expect(owner.released).toBe(false);
    await h.session.dispose();
  });

  it('times busy from hold opening and cancels timers on settlement', async () => {
    const h = harness({ holdTimeoutMs: 100 });
    const hold = h.session.hold('draft');
    h.tick(70);
    const read = h.session.read('text', []);
    h.tick(29);
    expect(h.worker.reads).toHaveLength(0);
    h.tick(1);
    expect(await read).toEqual({ ok: false, refusal: { code: 'busy' } });
    expect(h.timers.size).toBe(0);
    expect(h.worker.reads).toHaveLength(0);
    hold.release();
    await h.session.dispose();
  });

  it('refuses an already expired hold immediately and defaults to 2000ms', async () => {
    const h = harness();
    h.session.hold('draft');
    h.tick(2000);
    expect(await h.session.read('text', [])).toEqual({ ok: false, refusal: { code: 'busy' } });
    await h.session.dispose();
  });

  it('returns busy when a late release runs before its delayed timeout callback', async () => {
    const h = harness();
    try {
      const hold = h.session.hold('draft');
      const read = h.session.read('text', []);
      h.elapse(2000);
      hold.release();
      expect(h.worker.reads).toHaveLength(0);
      expect(await read).toEqual({ ok: false, refusal: { code: 'busy' } });
      expect(h.timers.size).toBe(0);
    } finally { await h.session.dispose(); }
  });

  it('releases a signal-owned hold and handles a signal already aborted', async () => {
    const h = harness();
    const signal = new AbortController();
    const hold = h.session.hold('gesture', { signal: signal.signal });
    const read = h.session.read('text', []);
    signal.abort();
    expect(hold.released).toBe(true);
    h.worker.answer(0);
    expect(await read).toMatchObject({ ok: true });
    expect(h.session.hold('draft', { signal: signal.signal }).released).toBe(true);
    await h.session.dispose();
  });

  it('aborts without posting when one signal owns both the hold and the read', async () => {
    const h = harness();
    try {
      const controller = new AbortController();
      const hold = h.session.hold('draft', { signal: controller.signal });
      const read = h.session.read('text', [], { signal: controller.signal });
      controller.abort();
      expect(hold.released).toBe(true);
      expect(h.worker.reads).toHaveLength(0);
      expect(await read).toEqual({ ok: false, refusal: { code: 'superseded', detail: 'aborted' } });
    } finally { await h.session.dispose(); }
  });

  it('waits for queued tickets at the captured watermark across capacity and cancellation', async () => {
    const h = harness({ maxInFlight: 1 });
    const first = h.session.submit('insert', 'a');
    h.session.submit('insert', 'b', { after: first });
    const read = h.session.read('text', []);
    h.pump();
    expect(h.worker.reads).toHaveLength(0);
    h.worker.ack(0, { outcomes: [{ seq: 1, status: 'refused', reason: 'read-only' }] });
    await drain();
    h.pump();
    expect(h.worker.reads[0]!.fence.afterSeq).toBe(1);
    h.worker.answer(0);
    expect(await read).toMatchObject({ ok: true });
    await h.session.dispose();
  });

  it('handles coalesced aliases at the watermark without gating on future input', async () => {
    const h = harness({ maxInFlight: 1, coalesce: (a, b) => ({ ...b, payload: a.payload + b.payload }) });
    h.session.submit('insert', 'a', { coalesceKey: 'text' });
    const read = h.session.read('text', []);
    h.session.submit('insert', 'b', { coalesceKey: 'text' });
    h.pump();
    expect(h.worker.edits[0]!.envelope.payload).toBe('ab');
    expect(h.worker.reads[0]!.fence.afterSeq).toBe(1);
    h.worker.answer(0);
    await read;
    await h.session.dispose();
  });

  for (const kind of ['doc', 'selection', 'layout'] as const) {
    it(`supersedes only declared ${kind} dependencies using pushed acknowledgements`, async () => {
      const h = harness();
      const read = h.session.read(kind, []);
      h.worker.emit({ version: kind === 'doc' ? 1 : 0,
        revisions: { selection: kind === 'selection' ? 1 : 0, layout: kind === 'layout' ? 1 : 0 } });
      h.worker.answer(0);
      expect(await read).toEqual({ ok: false, refusal: { code: 'superseded' } });
      const unaffected = h.session.read(kind, []);
      h.worker.answer(1, { doc: 1, selection: 1, layout: 1 });
      expect(await unaffected).toMatchObject({ ok: true });
      await h.session.dispose();
    });
  }

  it('ignores reversed ack revisions and passes host refusals through', async () => {
    const h = harness();
    h.session.submit('insert', 'a'); h.session.submit('insert', 'b'); h.pump();
    h.worker.ack(1, { revisions: { selection: 2, layout: 2 } }); await drain();
    h.worker.ack(0, { revisions: { selection: 1, layout: 1 } }); await drain();
    const read = h.session.read('selection', []);
    h.worker.answer(0, { selection: 1 });
    expect(await read).toEqual({ ok: false, refusal: { code: 'superseded' } });
    const refused = h.session.read('text', []);
    h.worker.reads[1]!.reply.resolve({ ok: false, refusal: { code: 'not-ready', detail: 'loading' } });
    expect(await refused).toEqual({ ok: false, refusal: { code: 'not-ready', detail: 'loading' } });
    await h.session.dispose();
    expect(h.worker.acks.size).toBe(0);
  });

  for (const phase of ['before', 'after'] as const) {
    it(`resolves abort ${phase} posting and discards late replies`, async () => {
      const h = harness();
      const signal = new AbortController();
      const hold = phase === 'before' ? h.session.hold('draft') : undefined;
      const read = h.session.read('text', [], { signal: signal.signal });
      signal.abort();
      expect(await read).toEqual({ ok: false, refusal: { code: 'superseded', detail: 'aborted' } });
      expect(h.timers.size).toBe(0);
      if (phase === 'after') h.worker.reads[0]!.reply.resolve({} as QueryReply);
      hold?.release(); await drain();
      expect(h.worker.reads).toHaveLength(phase === 'before' ? 0 : 1);
      expect(h.session.stage).toBe('ready');
      await h.session.dispose();
    });
  }

  it('refuses unsupported queries and disposed reads without posting', async () => {
    const h = harness();
    expect(await h.session.read('missing' as 'text', [])).toEqual({ ok: false, refusal: { code: 'unsupported' } });
    delete h.worker.connection.client.call.query;
    expect(await h.session.read('text', [])).toEqual({ ok: false, refusal: { code: 'unsupported' } });
    const hold = h.session.hold('draft');
    await h.session.dispose();
    expect(h.session.stage).toBe('disposed');
    expect(hold.released).toBe(true);
    expect(await h.session.read('text', [])).toEqual({ ok: false, refusal: { code: 'disposed' } });
    expect(h.worker.reads).toHaveLength(0);
  });

  for (const at of [null, { incarnation: 99 }, { epoch: 2 }, { throughSeq: 0 }]) {
    it(`fails malformed or fence-violating replies ${JSON.stringify(at)}`, async () => {
      const h = harness();
      h.session.submit('insert', 'a'); h.pump();
      const read = h.session.read('text', []);
      if (at === null) h.worker.reads[0]!.reply.resolve({} as QueryReply);
      else h.worker.answer(0, at);
      await expect(read).rejects.toMatchObject({ code: 'message' });
      expect(h.session.stage).toBe('failed');
      await h.session.dispose();
    });
  }

  it('rejects pending reads and releases holds on terminal failure', async () => {
    const h = harness();
    const hold = h.session.hold('draft');
    const held = h.session.read('text', []);
    const posted = h.session.read('text', [], { within: hold });
    const failure = new SessionFailure('crash', 'failed');
    h.worker.crash(failure);
    await expect(held).rejects.toBe(failure);
    await expect(posted).rejects.toBe(failure);
    await expect(h.session.read('text', [])).rejects.toBe(failure);
    expect(hold.released).toBe(true);
    expect(h.session.stage).toBe('failed');
    await h.session.dispose();
  });

  it('reissues a posted read only after replay on a fresh connection', async () => {
    const replacement = port(1);
    const h = harness({ maxRecoveryAttempts: 1, recovery: { reopen: async () => replacement.connection } });
    h.session.submit('insert', 'a'); h.pump(); h.worker.ack(0); await drain();
    const read = h.session.read('text', []);
    h.worker.crash();
    expect(h.session.stage).toBe('recovering');
    await drain();
    expect(replacement.reads).toHaveLength(0);
    expect(replacement.edits[0]!.envelope.replay).toBe(true);
    h.worker.reads[0]!.reply.resolve({} as QueryReply);
    replacement.ack(0); await drain(); h.pump();
    expect(replacement.reads[0]!.fence).toEqual({ session: 'reads', incarnation: 1, epoch: 1, afterSeq: 1 });
    replacement.answer(0);
    expect(await read).toMatchObject({ ok: true, at: { incarnation: 1 } });
    expect(h.session.stage).toBe('ready');
    await h.session.dispose();
  });

  for (const applied of [false, true]) for (const posted of [false, true]) {
    it(`fences ${posted ? 'a reissued' : 'a new'} read at the replayed frontier when the last sent command was refused${applied ? ' after an applied one' : ''}`, async () => {
      const replacement = port(1);
      const h = harness({ maxRecoveryAttempts: 1, recovery: { reopen: async () => replacement.connection } });
      if (applied) { h.session.submit('insert', 'a'); h.pump(); h.worker.ack(0); await drain(); }
      const refused = h.session.submit('insert', 'b'); h.pump();
      h.worker.ack(applied ? 1 : 0, { version: applied ? 1 : 0,
        outcomes: [{ seq: refused.seq!, status: 'refused', reason: 'locked' }] });
      expect(await refused.ack).toMatchObject({ status: 'refused' });
      const early = posted ? h.session.read('text', []) : undefined;
      if (posted) expect(h.worker.reads[0]!.fence.afterSeq).toBe(refused.seq!);
      h.worker.crash(); await drain();
      expect(replacement.edits.map((edit) => edit.envelope.seq)).toEqual(applied ? [1] : []);
      if (applied) { replacement.ack(0); await drain(); }
      expect(h.session.stage).toBe('ready');
      const read = early ?? h.session.read('text', []);
      expect(replacement.reads[0]!.fence.afterSeq).toBe(applied ? 1 : 0);
      replacement.answer(0);
      expect(await read).toMatchObject({ ok: true, at: { incarnation: 1 } });
      const next = h.session.submit('insert', 'c'); h.pump();
      const later = h.session.read('text', []);
      expect(replacement.reads[1]!.fence.afterSeq).toBe(next.seq!);
      replacement.answer(1);
      expect(await later).toMatchObject({ ok: true });
      await h.session.dispose();
    });
  }
});
