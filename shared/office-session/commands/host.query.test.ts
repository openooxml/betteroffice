import { describe, expect, it } from 'bun:test';
import { isDeferredReply, isTransferResult, transferable, type DeferredReply, type TransferResult } from '../protocol';
import { SessionFailure } from '../types';
import { createCommandHost, commandSessionMethods } from './host';
import { command, query, refuse, type QueryFence, type QueryReply } from './spec';
import { isCommandAck, type CommandEnvelope, type CommandHostOptions } from './types';

const commands = {
  insert: command<string>()({ admission: 'local-write', refusals: ['locked'] }),
  select: command<string>()({ admission: 'view' }),
  remote: command<string>()({ admission: 'remote-update' }),
  mode: command<string>()({ admission: 'mode' }),
};
const specs = {
  text: query<[], string>()({ stale: 'answer-at-fence', refusals: ['locked'] }),
  draft: query<[], string>()({ stale: { reject: ['doc', 'selection'] } }),
};
const fence: QueryFence = { session: 'query', incarnation: 2, epoch: 1, afterSeq: 0 };
function envelope(seq: number, kind = 'insert', replay?: true): CommandEnvelope<string> {
  return { session: 'query', incarnation: 2, seq, basis: { epoch: 1, version: 0 }, kind, payload: 'a', replay };
}
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}
async function drain() {
  for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
}
function observe(promise: Promise<unknown>) {
  const state: { done: boolean; result?: unknown } = { done: false };
  void promise.then((result) => { state.done = true; state.result = result; },
    (error: unknown) => { state.done = true; state.result = error; });
  return state;
}
async function value<T>(result: Promise<T | TransferResult<T>> | T | TransferResult<T> | DeferredReply<T>): Promise<T> {
  const reply = await (isDeferredReply(result) ? result.promise : result);
  return isTransferResult(reply) ? (reply as TransferResult<T>).value : reply;
}
function model(extra: Partial<CommandHostOptions<string, number, boolean, never>> = {}) {
  let version = 0;
  let layout = 0;
  let readOnly = false;
  let text = '';
  const applied: string[] = [];
  const host = createCommandHost<string, number, boolean, never>({
    session: 'query', incarnation: 2, epoch: 1, commands, queries: {
      specs, handlers: { text: () => text, draft: () => text },
    }, readOnly: () => readOnly, layoutRevision: () => layout, checkRefusals: true,
    getVersion: () => version,
    apply(command) {
      applied.push(command.kind);
      text += command.payload;
      version += 1;
      return { status: 'applied', selection: text.length, editState: true };
    }, frame: () => undefined, snapshot: () => new Uint8Array(), ...extra,
  });
  return { host, applied,
    readOnly(value: boolean) { readOnly = value; },
    layout(value: number) { layout = value; },
    version(value: number) { version = value; },
    edit(seq: number, kind = 'insert', replay?: true) {
      const command = envelope(seq, kind, replay);
      host.policy(command);
      return value(host.handler(null, command));
    },
    async read(afterSeq = 0, kind = 'text'): Promise<QueryReply> {
      const result = host.query(null, { ...fence, afterSeq }, kind, []);
      return isDeferredReply(result) ? value(result.promise) : value(result);
    },
  };
}

describe('worker query and admission contract', () => {
  it('refuses unknown and read-only local writes while admitting other classes', async () => {
    const m = model();
    m.readOnly(true);
    expect((await m.edit(1)).outcomes).toEqual([{ seq: 1, status: 'refused', reason: 'read-only', committed: false }]);
    expect((await m.edit(2, 'missing')).outcomes[0]?.reason).toBe('unsupported');
    for (const [index, kind] of ['select', 'remote', 'mode'].entries()) {
      expect((await m.edit(index + 3, kind)).outcomes[0]?.status).toBe('applied');
    }
    expect(m.applied).toEqual(['select', 'remote', 'mode']);
  });

  it('allows previously applied replay through read-only admission', async () => {
    const m = model({ replayAppliedSeqs: [1] });
    m.readOnly(true);
    expect((await m.edit(1, 'insert', true)).outcomes[0]?.status).toBe('applied');
  });

  it('keeps acknowledgements unchanged without query options', async () => {
    const m = model({ queries: undefined });
    expect(await m.edit(1)).toEqual({ session: 'query', incarnation: 2, epoch: 1, throughSeq: 1,
      version: 1, outcomes: [{ seq: 1, status: 'applied' }], selection: 1, editState: true });
  });

  it('stamps answers and preserves transfers and declared refusal values', async () => {
    const bytes = new Uint8Array([1]);
    const m = model({ queries: { specs, handlers: {
      text: () => transferable(bytes, [bytes.buffer]), draft: () => refuse('not-ready'),
    } } });
    await m.edit(1);
    m.layout(7);
    const result = await m.host.query(null, { ...fence, afterSeq: 1 }, 'text', []);
    expect(isTransferResult(result)).toBe(true);
    if (!isTransferResult(result)) throw new Error('Expected transfer');
    expect(result.transfer).toEqual([bytes.buffer]);
    expect(result.value).toEqual({ ok: true, value: bytes,
      at: { incarnation: 2, epoch: 1, throughSeq: 1, doc: 1, selection: 1, layout: 7 } });
    expect(await m.read(1, 'draft')).toEqual({ ok: false, refusal: { code: 'not-ready' } });
  });

  it('rejects invalid and unseen fences as values and ignores inherited query kinds', async () => {
    const m = model();
    for (const changed of [{ session: 'other' }, { incarnation: 3 }, { epoch: 2 }, { afterSeq: 1 }]) {
      expect(await value(m.host.query(null, { ...fence, ...changed }, 'text', [])))
        .toEqual({ ok: false, refusal: { code: 'stale' } });
    }
    for (const changed of [{ afterSeq: -1 }, { epoch: NaN }, { session: null }]) {
      expect(await value(m.host.query(null, { ...fence, ...changed } as QueryFence, 'text', [])))
        .toEqual({ ok: false, refusal: { code: 'invalid-payload' } });
    }
    expect(await value(m.host.query(null, fence, 'text', null as unknown as unknown[])))
      .toEqual({ ok: false, refusal: { code: 'invalid-payload' } });
    expect(await m.read(0, 'toString')).toEqual({ ok: false, refusal: { code: 'unsupported' } });
  });

  it('parks behind admitted input without blocking a later handler and releases in fence order', async () => {
    const m = model();
    m.host.policy(envelope(1));
    m.host.policy(envelope(2));
    const order: number[] = [];
    const second = m.host.query(null, { ...fence, afterSeq: 2 }, 'text', []);
    const first = m.host.query(null, { ...fence, afterSeq: 1 }, 'text', []);
    expect(isDeferredReply(first)).toBe(true);
    expect(isDeferredReply(second)).toBe(true);
    if (!isDeferredReply(first) || !isDeferredReply(second)) throw new Error('Expected parked reads');
    const a = first.promise.then((reply) => { order.push(1); return reply; });
    const b = second.promise.then((reply) => { order.push(2); return reply; });
    await m.host.handler(null, envelope(1));
    expect(await value(a)).toMatchObject({ ok: true, value: 'a' });
    expect(order).toEqual([1]);
    await m.host.handler(null, envelope(2));
    expect(await value(b)).toMatchObject({ ok: true, value: 'aa' });
    expect(order).toEqual([1, 2]);
  });

  it('tracks live sequences from handler entry and serializes queries after async apply', async () => {
    const gate = deferred<void>();
    const entered = deferred<void>();
    let applying = false;
    let version = 0;
    const m = model({ getVersion: () => version,
      async apply() { applying = true; entered.resolve(); await gate.promise; applying = false; version += 1; return { status: 'applied' }; },
      queries: { specs, handlers: { text: () => { expect(applying).toBe(false); return 'done'; }, draft: () => '' } },
    });
    const edit = m.host.handler(null, envelope(1));
    await entered.promise;
    const read = m.read(1);
    gate.resolve();
    await edit;
    expect(await read).toMatchObject({ ok: true, value: 'done', at: { throughSeq: 1 } });
  });

  it('chooses FIFO only for answer-at-fence reads', () => {
    const m = model();
    expect(m.host.queryPolicy(fence, 'text', [])).toEqual({ lane: 'interactive', reorderable: false });
    expect(m.host.queryPolicy(fence, 'draft', [])).toEqual({ lane: 'interactive', reorderable: true });
  });

  it('publishes only changed revisions and validates optional ack revision counters', async () => {
    const m = model();
    expect((await m.edit(1)).revisions).toEqual({ selection: 1, layout: 0 });
    expect(m.host.revise({ selection: 1, editState: true })).toBeUndefined();
    m.layout(4);
    expect(m.host.revise()).toMatchObject({ outcomes: [], throughSeq: 1, version: 1,
      revisions: { selection: 1, layout: 4 } });
    expect(m.host.revise()).toBeUndefined();
    m.version(2);
    expect(m.host.revise({ selection: 2 })).toMatchObject({ outcomes: [], version: 2,
      selection: 2, revisions: { selection: 2, layout: 4 } });
    const ack = await m.edit(2);
    expect(isCommandAck(ack)).toBe(true);
    expect(isCommandAck({ ...ack, revisions: { selection: -1, layout: 0 } })).toBe(false);
    expect(isCommandAck({ ...ack, revisions: { selection: 0 } })).toBe(false);
  });

  it('converts ordinary handler throws to invalid-payload and poisons terminal throws', async () => {
    const m = model({ queries: { specs, handlers: { text: () => { throw new Error('bad args'); }, draft: () => '' } } });
    expect(await m.read()).toEqual({ ok: false, refusal: { code: 'invalid-payload', detail: 'bad args' } });
    const failure = new SessionFailure('trap', 'terminal');
    const poisoned = model({ queries: { specs, handlers: { text: () => { throw failure; }, draft: () => '' } } });
    await expect(poisoned.read()).rejects.toBe(failure);
    await expect(poisoned.edit(1)).rejects.toBe(failure);
  });

  it('poisons undeclared command and query refusals with a TypeError', async () => {
    const m = model({ apply: () => ({ status: 'refused', reason: 'undeclared' }) });
    await expect(m.edit(1)).rejects.toBeInstanceOf(TypeError);
    const q = model({ queries: { specs, handlers: { text: () => refuse('undeclared'), draft: () => '' } } });
    await expect(q.read()).rejects.toBeInstanceOf(TypeError);
    await expect(q.edit(1)).rejects.toBeInstanceOf(TypeError);
    const declared = model({ queries: { specs, handlers: { text: () => refuse('locked'), draft: () => '' } } });
    expect(await declared.read()).toEqual({ ok: false, refusal: { code: 'locked' } });
  });

  it('rejects every parked query with the same command poison', async () => {
    const m = model({ apply() { throw new Error('unknown'); } });
    m.host.policy(envelope(1));
    const read = m.read(1).catch((error: unknown) => error);
    const failure = await m.host.handler(null, envelope(1)).catch((error: unknown) => error);
    expect(await read).toBe(failure);
  });

  for (const overflow of ['query', 'checkpoint'] as const) {
    for (const parked of ['query', 'frame'] as const) {
      it(`rejects a parked ${parked} when a ${overflow} exceeds handler capacity behind a held apply`, async () => {
        const entered = deferred<void>();
        const m = model({ apply: () => { entered.resolve(); return new Promise<never>(() => {}); } });
        void m.edit(1).catch(() => {});
        await entered.promise;
        if (parked === 'frame') m.host.policy(envelope(2));
        const waiter = observe(parked === 'query' ? m.read(1) : value(m.host.frame()));
        for (let index = 0; index < 4095; index += 1) void m.read(0).catch(() => {});
        const failure = await (overflow === 'query' ? m.read(0) : m.host.checkpoint(null)).catch((error: unknown) => error);
        expect(failure).toBeInstanceOf(SessionFailure);
        expect(failure).toMatchObject({ code: 'out-of-memory', message: 'Command handler capacity exceeded' });
        await drain();
        expect(waiter.done).toBe(true);
        expect(waiter.result).toBe(failure);
        await expect(m.read(0)).rejects.toBe(failure);
      });
    }
  }

  it('acknowledges the edit when a read it releases exhausts capacity and rejects later boundaries', async () => {
    const entered = deferred<void>();
    const gate = deferred<void>();
    let version = 0;
    const m = model({ getVersion: () => version,
      async apply() { entered.resolve(); await gate.promise; version += 1; return { status: 'applied' }; } });
    m.host.policy(envelope(2));
    const first = observe(m.read(1));
    const second = observe(m.read(2));
    const edit = value(m.host.handler(null, envelope(2)));
    await entered.promise;
    for (let index = 0; index < 4095; index += 1) void m.read(0).catch(() => {});
    gate.resolve();
    expect((await edit).outcomes).toEqual([{ seq: 2, status: 'applied' }]);
    await drain();
    expect(first.result).toBeInstanceOf(SessionFailure);
    expect(first.result).toMatchObject({ code: 'out-of-memory', message: 'Command handler capacity exceeded' });
    expect(second).toMatchObject({ done: true });
    expect(second.result).toBe(first.result);
    await expect(m.read(0)).rejects.toBe(first.result);
  });

  it('rejects parked queries and frames when revise poisons the host behind a held apply', async () => {
    const entered = deferred<void>();
    const m = model({ apply: () => { entered.resolve(); return new Promise<never>(() => {}); } });
    void m.edit(1).catch(() => {});
    await entered.promise;
    m.host.policy(envelope(2));
    const query = observe(m.read(1));
    const frame = observe(value(m.host.frame()));
    m.layout(-1);
    let failure: unknown;
    try { m.host.revise(); } catch (error) { failure = error; }
    expect(failure).toBeInstanceOf(SessionFailure);
    expect(failure).toMatchObject({ code: 'crash', message: 'Invalid layout revision' });
    await drain();
    expect(query).toMatchObject({ done: true });
    expect(query.result).toBe(failure);
    expect(frame).toMatchObject({ done: true });
    expect(frame.result).toBe(failure);
    await expect(m.read(0)).rejects.toBe(failure);
  });

  it('provides edit checkpoint and query method maps', () => {
    const m = model();
    const methods = commandSessionMethods(m.host);
    expect(Object.keys(methods.handlers)).toEqual(['edit', 'checkpoint', 'query']);
    expect(Object.keys(methods.policies)).toEqual(['edit', 'checkpoint', 'query']);
  });
});
