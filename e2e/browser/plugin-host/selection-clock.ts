export function selectionClock() {
  const timeout = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const dateNow = Date.now;
  const start = dateNow();
  let now = 0;
  let serial = -1;
  let holdWaits = false;
  const timers = new Map<number, { at: number; run(): void }>();
  globalThis.setTimeout = ((callback: (...args: unknown[]) => void, delay = 0, ...args: unknown[]) => {
    if (delay <= 0 || (!holdWaits && delay < 350)) return timeout(callback, delay, ...args);
    const id = serial--;
    timers.set(id, { at: now + delay, run: () => callback(...args) });
    return id;
  }) as unknown as typeof setTimeout;
  globalThis.clearTimeout = ((id: ReturnType<typeof setTimeout>) => {
    if (!timers.delete(id as unknown as number)) cancel(id);
  }) as typeof clearTimeout;
  Date.now = () => start + now;
  return {
    holdWaits() { holdWaits = true; },
    async tick(ms: number) {
      const end = now + ms;
      for (;;) {
        const next = [...timers].filter(([, timer]) => timer.at <= end)
          .sort((a, b) => a[1].at - b[1].at)[0];
        if (!next) break;
        now = next[1].at;
        timers.delete(next[0]);
        next[1].run();
        for (let turn = 0; turn < 40; turn++) await Promise.resolve();
      }
      now = end;
      for (let turn = 0; turn < 40; turn++) await Promise.resolve();
    },
    async expire() {
      for (let turn = 0; turn < 16 && timers.size; turn++) {
        await this.tick(Math.max(0, Math.min(...[...timers.values()].map(timer => timer.at)) - now));
      }
    },
    restore() {
      timers.clear();
      globalThis.setTimeout = timeout;
      globalThis.clearTimeout = cancel;
      Date.now = dateNow;
    },
  };
}
