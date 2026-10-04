import { describe, expect, it } from 'bun:test';
import { createSessionClient } from './client';
import { createSessionHost } from './host';
import { transferable } from './protocol';
import { createInProcessPair } from './testing/inProcessTransport';
import { SessionFailure, type MethodPolicies } from './types';

type Methods = {
  echo(value: unknown): unknown;
  add(a: number, b: number): number;
  slow(): string;
  mutate(value: string): string;
  throw(): never;
  trap(): never;
  transfer(buffer: ArrayBuffer): ArrayBuffer;
};
type Events = { tick: number };
const methods = ['echo', 'add', 'slow', 'mutate', 'throw', 'trap', 'transfer'] as const;
const policies: MethodPolicies<Methods> = {
  echo: { lane: 'interactive' }, add: { lane: 'interactive' },
  slow: { lane: 'interactive' }, mutate: { lane: 'input', mutates: true },
  throw: { lane: 'interactive' }, trap: { lane: 'interactive' }, transfer: { lane: 'interactive' },
};

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function session() {
  const pair = createInProcessPair();
  const held = deferred<string>();
  const started = deferred<void>();
  const disposed = deferred<void>();
  const order: string[] = [];
  let disposeCount = 0;
  let hostBuffer: ArrayBuffer | undefined;
  const host = createSessionHost<Methods, Events, string[]>(pair.host, {
    policies, context: order,
    handlers: {
      echo: (_, value) => value,
      add: (_, a, b) => a + b,
      slow: async (context) => { context.push('slow'); started.resolve(); return await held.promise; },
      mutate: (context, value) => { context.push(value); return value; },
      throw: () => { throw Object.assign(new Error('refused'), { refusal: { reason: 'busy' } }); },
      trap: () => { throw new WebAssembly.RuntimeError('unreachable'); },
      transfer: (_, buffer) => { hostBuffer = buffer; return transferable(buffer, [buffer]); },
    },
    onDispose: () => { disposeCount += 1; disposed.resolve(); },
  });
  const client = createSessionClient<Methods, Events>(pair.client, { methods });
  return { pair, host, client, held, started, disposed, order,
    get disposeCount() { return disposeCount; },
    get hostBuffer() { return hostBuffer; },
  };
}

describe('session host and cloned transport', () => {
  it('round trips calls, serializes handlers and delivers events in order', async () => {
    const s = session();
    const ticks: number[] = [];
    s.client.on('tick', (tick) => { ticks.push(tick); });
    s.host.emit('tick', 1);
    s.host.emit('tick', 2);
    expect(await s.client.call.add(2, 3)).toBe(5);
    const value = { text: 'echo', nested: [1, 2] };
    const echoed = await s.client.call.echo(value);
    expect(echoed).toEqual(value);
    expect(echoed).not.toBe(value);
    expect(ticks).toEqual([1, 2]);
    const completions: string[] = [];
    const slow = s.client.call.slow().then((value) => { completions.push(value); });
    await s.started.promise;
    const first = s.client.call.mutate('first').then((value) => { completions.push(value); });
    const second = s.client.call.mutate('second').then((value) => { completions.push(value); });
    s.held.resolve('done');
    await Promise.all([slow, first, second]);
    expect(s.order).toEqual(['slow', 'first', 'second']);
    expect(completions).toEqual(['done', 'first', 'second']);
    await s.client.dispose();
  });

  it('rejects non-cloneable arguments without failing the session', async () => {
    const s = session();
    const error = await s.client.call.echo(() => {}).catch((error: Error) => error);
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('DataCloneError');
    expect(error.message.length).toBeGreaterThan(0);
    expect(s.client.failure).toBeUndefined();
    expect(await s.client.call.add(1, 2)).toBe(3);
    await s.client.dispose();
  });

  it('detaches buffers in both directions and transfers event payloads', async () => {
    const s = session();
    const buffer = new Uint8Array([1, 2, 3]).buffer;
    const reply = s.client.callWithTransfer('transfer', [buffer], [buffer]);
    expect(buffer.byteLength).toBe(0);
    const received = await reply;
    expect([...new Uint8Array(received)]).toEqual([1, 2, 3]);
    expect(s.hostBuffer?.byteLength).toBe(0);
    await s.client.dispose();
    const pair = createInProcessPair();
    const events = deferred<ArrayBuffer>();
    const host = createSessionHost<{}, { buffer: ArrayBuffer }, null>(pair.host, {
      handlers: {}, policies: {}, context: null,
    });
    const client = createSessionClient<{}, { buffer: ArrayBuffer }>(pair.client, { methods: [] });
    client.on('buffer', (value) => events.resolve(value));
    const payload = new Uint8Array([4]).buffer;
    host.emit('buffer', payload, [payload]);
    expect(payload.byteLength).toBe(0);
    expect([...new Uint8Array(await events.promise)]).toEqual([4]);
    await client.dispose();
  });

  it('keeps ordinary errors and unknown methods non-terminal', async () => {
    const s = session();
    const error = await s.client.call.throw().catch((error) => error);
    expect(error.message).toBe('refused');
    expect(error.refusal).toEqual({ reason: 'busy' });
    expect(await s.client.call.add(3, 4)).toBe(7);
    const reply = deferred<unknown>();
    const off = s.pair.client.listen((message) => reply.resolve(message));
    s.pair.client.post({ protocol: 1, kind: 'call', id: 999, method: 'toString', args: [] });
    expect(await reply.promise).toEqual({ protocol: 1, kind: 'reply', id: 999, ok: false,
      error: { name: 'Error', message: 'Unknown session method: toString' } });
    off();
    expect(await s.client.call.add(4, 5)).toBe(9);
    await s.client.dispose();
  });

  it('fails once on a trap, rejects queued and later calls, and skips queued mutations', async () => {
    const s = session();
    const failures: SessionFailure[] = [];
    s.client.onFailure((error) => { failures.push(error); });
    const trap = s.client.call.trap().catch((error) => error);
    const queued = s.client.call.mutate('skipped').catch((error) => error);
    expect(await trap).toBeInstanceOf(SessionFailure);
    expect(s.client.failure?.code).toBe('trap');
    expect(await queued).toBe(s.client.failure);
    expect(await s.client.call.add(1, 2).catch((error) => error)).toBe(s.client.failure);
    expect(failures).toEqual([s.client.failure!]);
    expect(s.order).toEqual([]);
  });

  it('disposes once, rejects pending and later calls, and suppresses failure notifications', async () => {
    const s = session();
    const failures: SessionFailure[] = [];
    s.client.onFailure((error) => { failures.push(error); });
    const pending = s.client.call.slow().catch((error) => error);
    await s.started.promise;
    const disposal = s.client.dispose();
    expect(s.client.dispose()).toBe(disposal);
    expect((await pending).code).toBe('disposed');
    expect(await s.client.call.add(1, 2).catch((error) => error)).toBe(s.client.failure);
    await disposal;
    await s.disposed.promise;
    s.held.resolve('ignored');
    expect(s.disposeCount).toBe(1);
    expect(failures).toEqual([]);
  });

  it('makes malformed client messages terminal', async () => {
    const s = session();
    s.pair.client.post({ protocol: 2, kind: 'call', id: 1, method: 'echo', args: [] });
    const pending = s.client.call.add(1, 2).catch((error) => error);
    expect((await pending).code).toBe('message');
    expect(s.client.failure?.code).toBe('message');
  });

  it('closing either endpoint stops already queued delivery', async () => {
    for (const side of ['client', 'host'] as const) {
      const pair = createInProcessPair();
      const messages: unknown[] = [];
      pair.host.listen((message) => { messages.push(message); });
      pair.client.post({ value: 1 });
      pair[side].close();
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(messages).toEqual([]);
    }
  });
});
