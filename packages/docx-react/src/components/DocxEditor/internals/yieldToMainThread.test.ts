import { afterEach, expect, mock, test } from 'bun:test';
import { yieldToMainThread } from './yieldToMainThread';

const restores = new Set<() => void>();

afterEach(() => {
  for (const restore of [...restores].reverse()) restore();
});

function stubGlobal(name: string, value: unknown): () => void {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, name);
  const restore = () => {
    if (!restores.delete(restore)) return;
    if (descriptor) Object.defineProperty(globalThis, name, descriptor);
    else Reflect.deleteProperty(globalThis, name);
  };
  restores.add(restore);
  Object.defineProperty(globalThis, name, { configurable: true, writable: true, value });
  return restore;
}

test('yieldToMainThread uses scheduler.yield when available', async () => {
  let release!: () => void;
  const pending = new Promise<void>((resolve) => { release = resolve; });
  const scheduler = { yield: mock(() => pending) };
  const channel = mock(() => { throw new Error('Unexpected channel'); });
  const restoreScheduler = stubGlobal('scheduler', scheduler);
  const restoreChannel = stubGlobal('MessageChannel', channel);
  try {
    expect(yieldToMainThread()).toBe(pending);
    expect(scheduler.yield).toHaveBeenCalledTimes(1);
    expect(channel).not.toHaveBeenCalled();
    release();
    await pending;
  } finally {
    restoreChannel();
    restoreScheduler();
  }
});

test('yieldToMainThread waits for a MessageChannel round trip and closes both ports', async () => {
  const channel = {
    port1: { onmessage: null as ((event: MessageEvent) => void) | null, close: mock(() => {}) },
    port2: { postMessage: mock(() => {}), close: mock(() => {}) },
  };
  const createChannel = mock(function () { return channel; });
  const timeout = mock(() => { throw new Error('Unexpected timer'); });
  const restoreScheduler = stubGlobal('scheduler', undefined);
  const restoreChannel = stubGlobal('MessageChannel', createChannel);
  const restoreTimeout = stubGlobal('setTimeout', timeout);
  try {
    let settled = false;
    const pending = yieldToMainThread();
    void pending.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);
    expect(createChannel).toHaveBeenCalledTimes(1);
    expect(channel.port2.postMessage).toHaveBeenCalledTimes(1);
    expect(timeout).not.toHaveBeenCalled();
    channel.port1.onmessage!({} as MessageEvent);
    await pending;
    expect(settled).toBe(true);
    expect(channel.port1.close).toHaveBeenCalledTimes(1);
    expect(channel.port2.close).toHaveBeenCalledTimes(1);
  } finally {
    restoreTimeout();
    restoreChannel();
    restoreScheduler();
  }
});

test('yieldToMainThread uses a zero-delay timer only without MessageChannel', async () => {
  let release!: () => void;
  const timeout = mock((callback: () => void) => { release = callback; });
  const restoreScheduler = stubGlobal('scheduler', undefined);
  const restoreChannel = stubGlobal('MessageChannel', undefined);
  const restoreTimeout = stubGlobal('setTimeout', timeout);
  try {
    let settled = false;
    const pending = yieldToMainThread();
    void pending.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);
    expect(timeout).toHaveBeenCalledWith(expect.any(Function), 0);
    release();
    await pending;
    expect(settled).toBe(true);
  } finally {
    restoreTimeout();
    restoreChannel();
    restoreScheduler();
  }
});
