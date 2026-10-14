import { afterEach, beforeEach, expect } from 'bun:test';

type Act = (callback: () => Promise<void>) => Promise<unknown>;
type Scope = {
  test: number;
  ended: boolean;
  polls: Set<Promise<unknown>>;
  acts: Set<Promise<unknown>>;
};

// Bun keeps running a timed-out body; once its test ends, act/settle/until reject instead of acting.
// Call before the file's afterEach hooks: Bun runs them in registration order, so the scope ends and
// its polls drain before cleanup, and the next beforeEach waits for its acts that cleanup unblocks.
export function testScope(realAct: Act) {
  let started = 0;
  let current: Scope | null = null;

  afterEach(async () => {
    if (!current) return;
    current.ended = true;
    await Promise.allSettled(current.polls);
  });
  beforeEach(async () => {
    const previous = current;
    if (previous) {
      previous.ended = true;
      await Promise.allSettled([...previous.polls, ...previous.acts]);
      // A macrotask turn lets the ended body's continuation (finally blocks, follow-up calls) run while its scope is current.
      await new Promise((done) => setTimeout(done, 0));
    }
    if (current === previous) {
      current = { test: ++started, ended: false, polls: new Set(), acts: new Set() };
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
  const run = async (
    scope: Scope,
    tracked: Set<Promise<unknown>>,
    callback: () => Promise<void>,
    call: string
  ) => {
    // React's act thenable must be subscribed once; the hooks wait on this promise instead.
    const pending = Promise.resolve(realAct(callback));
    tracked.add(pending);
    try {
      await pending;
    } finally {
      tracked.delete(pending);
    }
    live(scope, call);
  };
  const sleep = (ms: number) => async () => {
    await new Promise((done) => setTimeout(done, ms));
  };

  return {
    async act(callback: () => Promise<void>) {
      const scope = live(current, 'act');
      await run(scope, scope.acts, callback, 'act');
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
