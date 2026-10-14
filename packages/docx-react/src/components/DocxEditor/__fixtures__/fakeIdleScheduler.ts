import { expect, spyOn } from 'bun:test';

export function fakeIdleScheduler({
  timers: fakeTimers = false, noIdle = false, now,
}: { timers?: boolean; noIdle?: boolean; now?: () => number } = {}) {
  const originalIdle = globalThis.requestIdleCallback;
  const originalCancelIdle = globalThis.cancelIdleCallback;
  const idleWork = new Map<number, { callback: IdleRequestCallback; timeout?: number }>();
  const timers = new Map<number, { at: number; run: () => void }>();
  let next = 0;
  let time = 0;
  const clock = now || fakeTimers
    ? spyOn(performance, 'now').mockImplementation(now ?? (() => time)) : null;
  const timeout = fakeTimers ? spyOn(globalThis, 'setTimeout').mockImplementation(
    ((run: () => void, delay = 0) => {
      const id = ++next;
      timers.set(id, { at: performance.now() + delay, run });
      return id;
    }) as unknown as typeof setTimeout
  ) : null;
  const clear = fakeTimers ? spyOn(globalThis, 'clearTimeout').mockImplementation((id) => {
    timers.delete(id as unknown as number);
  }) : null;
  globalThis.requestIdleCallback = noIdle
    ? undefined as unknown as typeof requestIdleCallback
    : (callback, options) => {
      const id = ++next;
      idleWork.set(id, { callback, timeout: options?.timeout });
      return id;
    };
  globalThis.cancelIdleCallback = (id) => { idleWork.delete(id); };
  const flushIdle = (deadline: IdleDeadline = { didTimeout: false, timeRemaining: () => 40 }) => {
    for (const [id, { callback }] of [...idleWork]) {
      if (idleWork.delete(id)) callback(deadline);
    }
  };
  return {
    idleWork, timers, flushIdle,
    get now() { return time; },
    set now(value: number) { time = value; },
    setNow(value: number) { time = value; },
    flushOneIdle(deadline?: IdleDeadline) {
      expect(idleWork.size).toBe(1);
      flushIdle(deadline);
    },
    flushTimer() {
      expect(timers.size).toBe(1);
      const [id, timer] = timers.entries().next().value!;
      expect(timer.at).toBeLessThanOrEqual(performance.now());
      timers.delete(id);
      timer.run();
    },
    flushDueTimers() {
      for (const [id, timer] of [...timers]) {
        if (timer.at <= performance.now() && timers.delete(id)) timer.run();
      }
    },
    restore() {
      timeout?.mockRestore();
      clear?.mockRestore();
      clock?.mockRestore();
      idleWork.clear();
      timers.clear();
      globalThis.requestIdleCallback = originalIdle;
      globalThis.cancelIdleCallback = originalCancelIdle;
    },
  };
}

export function checkIdleContinuation(
  scheduler: ReturnType<typeof fakeIdleScheduler>,
  scenario: string,
  {
    setNow, dispose, checkFirst, checkTimer, timerSlices, noIdleStart = 0,
    offset = () => 0, run = (callback) => callback(), replayIdle = false, resumeIdle = false, deadline,
  }: {
    setNow: (now: number) => void; dispose: () => void;
    checkFirst: () => void; checkTimer: (index: number) => void; timerSlices: number;
    noIdleStart?: number; offset?: () => number; run?: (callback: () => void) => void;
    replayIdle?: boolean; resumeIdle?: boolean; deadline?: IdleDeadline;
  }
) {
  const noIdle = scenario.includes('no idle');
  const expiryDuringWork = scenario.startsWith('expiry during');
  setNow(noIdle ? noIdleStart : expiryDuringWork ? 4999 : 4990);
  run(() => noIdle ? scheduler.flushTimer() : scheduler.flushOneIdle(deadline));
  checkFirst();
  expect([...scheduler.timers.values()].map(({ at }) => at)).toEqual([
    noIdle ? performance.now() : Math.max(5000, performance.now()),
  ]);
  if (scenario.startsWith('dispose')) dispose();
  else {
    if (!noIdle && !expiryDuringWork && !resumeIdle) {
      setNow(4999 - offset());
      expect([...scheduler.timers.values()].every(({ at }) => at > performance.now())).toBe(true);
      setNow(5000 - offset());
    }
    const held = Array.from(scheduler.idleWork.values(), ({ callback }) => callback);
    for (let index = 0; index < timerSlices; index++) {
      run(() => {
        if (resumeIdle) scheduler.flushOneIdle();
        else scheduler.flushTimer();
        if (replayIdle && index === timerSlices - 1) {
          held.forEach((callback) => callback({ didTimeout: true, timeRemaining: () => 0 }));
        }
      });
      checkTimer(index);
    }
  }
  expect(scheduler.timers.size).toBe(0);
  expect(scheduler.idleWork.size).toBe(0);
}
