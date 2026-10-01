import { afterEach, expect, spyOn, test } from 'bun:test';
import { scheduleIdlePageBuild } from './pageBuildScheduler';

const originalIdle = globalThis.requestIdleCallback;
const originalCancelIdle = globalThis.cancelIdleCallback;
const originalFrame = globalThis.requestAnimationFrame;
const originalCancelFrame = globalThis.cancelAnimationFrame;

afterEach(() => {
  globalThis.requestIdleCallback = originalIdle;
  globalThis.cancelIdleCallback = originalCancelIdle;
  globalThis.requestAnimationFrame = originalFrame;
  globalThis.cancelAnimationFrame = originalCancelFrame;
});

test('idle work respects both the idle deadline and an eight millisecond budget', () => {
  let callback!: IdleRequestCallback;
  globalThis.requestIdleCallback = ((run: IdleRequestCallback) => {
    callback = run;
    return 1;
  }) as typeof requestIdleCallback;
  let now = 100;
  const clock = spyOn(performance, 'now').mockImplementation(() => now);
  const remaining: number[] = [];
  try {
    scheduleIdlePageBuild((deadline) => {
      remaining.push(deadline.timeRemaining());
      now += 6;
      remaining.push(deadline.timeRemaining());
      now += 3;
      remaining.push(deadline.timeRemaining());
    });
    callback({ didTimeout: false, timeRemaining: () => 5 });
    expect(remaining).toEqual([5, 2, 0]);
  } finally {
    clock.mockRestore();
  }
});

test('the frame fallback skips busy frames and cancellation stops the deferred frame', () => {
  globalThis.requestIdleCallback = undefined as unknown as typeof requestIdleCallback;
  const callbacks = new Map<number, FrameRequestCallback>();
  let next = 1;
  globalThis.requestAnimationFrame = (run) => {
    const id = next++;
    callbacks.set(id, run);
    return id;
  };
  globalThis.cancelAnimationFrame = (id) => { callbacks.delete(id); };
  let now = 110;
  const clock = spyOn(performance, 'now').mockImplementation(() => now);
  let runs = 0;
  try {
    const task = scheduleIdlePageBuild(() => { runs += 1; });
    const first = callbacks.get(1)!;
    callbacks.delete(1);
    first(100);
    expect(runs).toBe(0);
    expect([...callbacks.keys()]).toEqual([2]);
    task.cancel();
    expect(callbacks.size).toBe(0);
    now = 201;
    scheduleIdlePageBuild((deadline) => {
      expect(deadline.timeRemaining()).toBe(7);
      runs += 1;
    });
    callbacks.get(3)!(200);
    expect(runs).toBe(1);
  } finally {
    clock.mockRestore();
  }
});
