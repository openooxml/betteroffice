import { afterEach, beforeEach, expect } from 'bun:test';

type Act = <T>(callback: () => T | Promise<T>) => unknown;
type Scope = {
  test: number;
  ended: boolean;
  polls: Set<Promise<unknown>>;
  acts: Set<Promise<unknown>>;
};

async function unsettledAfter(pending: Promise<unknown>[], ms: number): Promise<number> {
  let left = pending.length;
  for (const promise of pending) {
    void promise.then(
      () => { left -= 1; },
      () => { left -= 1; }
    );
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  await Promise.race([
    Promise.allSettled(pending),
    new Promise((done) => { timer = setTimeout(done, ms); }),
  ]);
  clearTimeout(timer);
  return left;
}

export function testScope(realAct: Act, { drainMs = 2000 }: { drainMs?: number } = {}) {
  let started = 0;
  let current: Scope | null = null;

  afterEach(async () => {
    if (!current) return;
    current.ended = true;
    await Promise.allSettled(current.polls);
  });
  beforeEach(async () => {
    const previous = current;
    let unfinished = 0;
    if (previous) {
      previous.ended = true;
      unfinished = await unsettledAfter([...previous.polls, ...previous.acts], drainMs);
      await new Promise((done) => setTimeout(done, 0));
    }
    if (current === previous) {
      current = { test: ++started, ended: false, polls: new Set(), acts: new Set() };
    }
    if (previous && unfinished > 0) {
      throw new Error(
        `test #${previous.test} of this file left ${unfinished} unfinished act() or poll call(s) ${drainMs} ms after it ended; later tests no longer wait for them`
      );
    }
  });

  const live = (scope: Scope | null, call: string): Scope => {
    if (!scope) throw new Error(`${call}() needs a running test`);
    if (scope.ended) {
      throw new Error(
        `${call}() from test #${scope.test} of this file ran after that test ended: its body outlived the test, likely past its timeout`
      );
    }
    return scope;
  };
  const run = async <T>(
    scope: Scope,
    tracked: Set<Promise<unknown>>,
    callback: () => T | Promise<T>,
    call: string
  ): Promise<T> => {
    const pending = Promise.resolve(realAct(callback)) as Promise<T>;
    tracked.add(pending);
    let value: T;
    try {
      value = await pending;
    } finally {
      tracked.delete(pending);
    }
    live(scope, call);
    return value;
  };
  const sleep = (ms: number) => async () => {
    await new Promise((done) => setTimeout(done, ms));
  };

  return {
    guard<T extends unknown[], R>(callback: (...args: T) => R) {
      const scope = live(current, 'guard');
      return (...args: T): R => {
        live(scope, 'guard');
        return callback(...args);
      };
    },
    act<T>(callback: () => T | Promise<T>): Promise<T> {
      const scope = live(current, 'act');
      return run(scope, scope.acts, callback, 'act');
    },
    async settle(ms = 20) {
      const scope = live(current, 'settle');
      await run(scope, scope.polls, sleep(ms), 'settle');
    },
    async until(done: () => boolean) {
      const scope = live(current, 'until');
      for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
        await run(scope, scope.polls, sleep(10), 'until');
      }
      expect(done()).toBe(true);
    },
  };
}
