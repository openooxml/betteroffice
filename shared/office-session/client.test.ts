import { describe, expect, it } from 'bun:test';
import { createSessionClient, type SessionClientOptions } from './client';
import { SessionFailure } from './types';
import type { SessionTransport } from './transport';

function harness(onWasmModule?: (url: string, module: WebAssembly.Module) => void) {
  let now = 0;
  let next = 0;
  const timers = new Map<number, { at: number; callback: () => void }>();
  const messages: unknown[] = [];
  let receive = (_message: unknown) => {};
  let error = (_error: unknown) => {};
  let closes = 0;
  const transport: SessionTransport = {
    post: (message) => { messages.push(message); },
    listen: (listener) => { receive = listener; return () => { receive = () => {}; }; },
    onError: (listener) => { error = listener; return () => { error = () => {}; }; },
    close: () => { closes += 1; },
  };
  const client = createSessionClient<{ echo(value: string): string }, { tick: number }>(transport, {
    methods: { echo: true }, silenceMs: 100, now: () => now, onWasmModule,
    timer: (callback, ms) => {
      const id = next++;
      timers.set(id, { at: now + ms, callback });
      return () => { timers.delete(id); };
    },
  });
  return {
    client, messages,
    receive: (message: unknown) => receive(message),
    error: (value: unknown) => error(value),
    get closes() { return closes; },
    get timerCount() { return timers.size; },
    advance(ms: number) {
      now += ms;
      for (const [id, timer] of [...timers]) {
        if (timer.at > now || !timers.has(id)) continue;
        timers.delete(id);
        timer.callback();
      }
    },
  };
}

describe('session client', () => {
  it('receives compiled modules independently of pending replies', async () => {
    const module = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    const modules: Array<{ url: string; module: WebAssembly.Module }> = [];
    const h = harness((url, module) => { modules.push({ url, module }); });
    const reply = h.client.call.echo('pending');
    h.advance(80);
    h.receive({ protocol: 1, kind: 'wasm-module', url: 'https://example.test/module.wasm', module });
    h.advance(80);
    expect(h.client.failure).toBeUndefined();
    expect(modules).toEqual([{ url: 'https://example.test/module.wasm', module }]);
    expect(h.timerCount).toBe(1);
    h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'pending' });
    expect(await reply).toBe('pending');
    expect(h.timerCount).toBe(0);
    await h.client.dispose();
    h.receive({ protocol: 1, kind: 'wasm-module', url: 'https://example.test/late.wasm', module });
    expect(modules).toHaveLength(1);
  });

  it('ignores valid compiled modules when no cache listener is installed', async () => {
    const module = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    const h = harness();
    h.receive({ protocol: 1, kind: 'wasm-module', url: 'https://example.test/module.wasm', module });
    expect(h.client.failure).toBeUndefined();
    await h.client.dispose();
  });

  it('requires every method in the method record', () => {
    type Methods = { echo(value: string): string; add(a: number, b: number): number };
    // @ts-expect-error
    const methods: SessionClientOptions<Methods>['methods'] = { echo: true };
    expect(methods.echo).toBe(true);
  });

  it('skips event listeners removed during notification', async () => {
    const h = harness();
    const ticks: number[] = [];
    let off = () => {};
    h.client.on('tick', () => { off(); });
    off = h.client.on('tick', (value) => { ticks.push(value); });
    h.receive({ protocol: 1, kind: 'event', name: 'tick', payload: 1 });
    expect(ticks).toEqual([]);
    await h.client.dispose();
  });

  it('stops event notification when a listener disposes the client', async () => {
    const h = harness();
    const ticks: number[] = [];
    h.client.on('tick', () => { void h.client.dispose(); });
    h.client.on('tick', (value) => { ticks.push(value); });
    h.receive({ protocol: 1, kind: 'event', name: 'tick', payload: 1 });
    expect(ticks).toEqual([]);
    await h.client.dispose();
  });

  it('skips failure listeners removed during notification', () => {
    const h = harness();
    const failures: SessionFailure[] = [];
    let off = () => {};
    h.client.onFailure(() => { off(); });
    off = h.client.onFailure((error) => { failures.push(error); });
    h.error(new SessionFailure('crash', 'failed'));
    expect(failures).toEqual([]);
  });

  it('arms only for pending calls and rejects every call on silence', async () => {
    const h = harness();
    h.advance(1000);
    expect(h.client.failure).toBeUndefined();
    expect(h.timerCount).toBe(0);
    const failures: SessionFailure[] = [];
    h.client.onFailure((error) => { failures.push(error); });
    const first = h.client.call.echo('first').catch((error) => error);
    const second = h.client.call.echo('second').catch((error) => error);
    h.advance(99);
    expect(h.client.failure).toBeUndefined();
    h.advance(1);
    expect(h.client.failure?.code).toBe('silence');
    expect(await first).toBe(h.client.failure);
    expect(await second).toBe(h.client.failure);
    expect(await h.client.call.echo('later').catch((error) => error)).toBe(h.client.failure);
    h.error(new Error('another error'));
    expect(failures).toEqual([h.client.failure!]);
    expect(h.closes).toBe(1);
    expect(h.timerCount).toBe(0);
  });

  it('rearms on events and replies without resetting for additional calls', async () => {
    const h = harness();
    const ticks: number[] = [];
    h.client.on('tick', (value) => { ticks.push(value); });
    const first = h.client.call.echo('first');
    h.advance(80);
    const second = h.client.call.echo('second').catch((error) => error);
    h.receive({ protocol: 1, kind: 'event', name: 'tick', payload: 1 });
    h.advance(80);
    expect(h.client.failure).toBeUndefined();
    h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'first' });
    expect(await first).toBe('first');
    h.advance(99);
    expect(h.client.failure).toBeUndefined();
    h.advance(1);
    expect((await second).code).toBe('silence');
    expect(ticks).toEqual([1]);
    expect(h.messages).toEqual([
      { protocol: 1, kind: 'call', id: 1, method: 'echo', args: ['first'] },
      { protocol: 1, kind: 'call', id: 2, method: 'echo', args: ['second'] },
    ]);
  });

  it('does not extend silence when more calls are posted', async () => {
    const h = harness();
    const first = h.client.call.echo('first').catch((error) => error);
    h.advance(99);
    const second = h.client.call.echo('second').catch((error) => error);
    h.advance(1);
    expect((await first).code).toBe('silence');
    expect(await second).toBe(h.client.failure);
  });

  it('disarms after the last reply and supports event unsubscription', async () => {
    const h = harness();
    const ticks: number[] = [];
    const off = h.client.on('tick', (tick) => { ticks.push(tick); });
    const reply = h.client.call.echo('ok');
    h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'ok' });
    expect(await reply).toBe('ok');
    h.receive({ protocol: 1, kind: 'event', name: 'tick', payload: 1 });
    off();
    h.receive({ protocol: 1, kind: 'event', name: 'tick', payload: 2 });
    h.advance(1000);
    expect(h.client.failure).toBeUndefined();
    expect(ticks).toEqual([1]);
    expect(h.timerCount).toBe(0);
    await h.client.dispose();
  });

  it('makes malformed inbound messages terminal', async () => {
    for (const message of [null, { kind: 'event', name: 'tick', payload: 1 },
      { protocol: 1, kind: 'wasm-compile' },
      { protocol: 1, kind: 'reply', id: 1, ok: true },
      { protocol: 1, kind: 'reply', id: 1, ok: false, error: {} }]) {
      const h = harness();
      const reply = h.client.call.echo('pending').catch((error) => error);
      h.receive(message);
      expect(await reply).toBeInstanceOf(SessionFailure);
      expect(h.client.failure?.code).toBe('message');
    }
  });

  it('preserves transport failure codes and isolates throwing listeners', async () => {
    const h = harness();
    const errors: SessionFailure[] = [];
    h.client.onFailure(() => { throw new Error('listener'); });
    h.client.onFailure((error) => { errors.push(error); });
    const reply = h.client.call.echo('pending').catch((error) => error);
    h.error(new SessionFailure('message', 'unreadable'));
    expect((await reply).code).toBe('message');
    expect(errors).toEqual([h.client.failure!]);
  });
});
