import { describe, expect, it } from 'bun:test';
import { createScopeTransport, createWorkerTransport, type SessionScope } from './transport';
import { SessionFailure } from './types';

function scope() {
  const listeners = new Map<string, Set<EventListener>>();
  const posts: Array<{ message: unknown; transfer: Transferable[] }> = [];
  let closes = 0;
  const port: SessionScope = {
    postMessage: (message, transfer) => { posts.push({ message, transfer }); },
    addEventListener(type, listener) {
      const set = listeners.get(type) ?? new Set<EventListener>();
      set.add(listener);
      listeners.set(type, set);
    },
    removeEventListener: (type, listener) => { listeners.get(type)?.delete(listener); },
    close: () => { closes += 1; },
  };
  return { port, posts,
    get closes() { return closes; },
    dispatch(type: string, fields: object) {
      for (const listener of [...(listeners.get(type) ?? [])]) listener(fields as Event);
    },
  };
}

describe('session transports', () => {
  it('forwards messages and transfers, unsubscribes, and closes a scope once', () => {
    const s = scope();
    const transport = createScopeTransport(s.port);
    const received: unknown[] = [];
    const off = transport.listen((message) => { received.push(message); });
    s.dispatch('message', { data: 1 });
    off();
    s.dispatch('message', { data: 2 });
    const buffer = new ArrayBuffer(1);
    transport.post({ buffer }, [buffer]);
    expect(s.posts).toEqual([{ message: { buffer }, transfer: [buffer] }]);
    expect(received).toEqual([1]);
    transport.close();
    transport.close();
    expect(s.closes).toBe(1);
    expect(() => transport.post(null)).toThrow(SessionFailure);
  });

  it('maps worker errors to crash and message failures and terminates once', () => {
    const s = scope();
    const worker = Object.assign(s.port, { terminate: () => s.port.close?.() });
    const transport = createWorkerTransport(worker as unknown as Worker);
    const errors: SessionFailure[] = [];
    const off = transport.onError((error) => { errors.push(error as SessionFailure); });
    s.dispatch('error', { message: 'crashed' });
    s.dispatch('messageerror', {});
    expect(errors.map((error) => error.code)).toEqual(['crash', 'message']);
    expect(errors[0]?.message).toBe('crashed');
    off();
    s.dispatch('error', { message: 'ignored' });
    transport.close();
    transport.close();
    expect(errors).toHaveLength(2);
    expect(s.closes).toBe(1);
  });
});
