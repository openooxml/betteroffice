import { afterEach, beforeEach, expect } from 'bun:test';

type Act = (callback: () => Promise<void>) => Promise<unknown>;
type Scope = { test: number; ended: boolean; acts: Set<Promise<unknown>> };

// Bun keeps running a timed-out body; once its test ends, settle/until reject instead of acting.
// Call before the file's afterEach hooks: Bun runs them in registration order, so this ends first.
export function testScope(act: Act) {
  let started = 0;
  let current: Scope | null = null;

  const end = async () => {
    if (!current) return;
    current.ended = true;
    await Promise.allSettled(current.acts);
  };
  beforeEach(async () => {
    await end();
    current = { test: ++started, ended: false, acts: new Set() };
  });
  afterEach(end);

  const live = (scope: Scope | null, call: string): Scope => {
    if (!scope) throw new Error(`${call}() needs a running test`);
    if (scope.ended) {
      throw new Error(
        `${call}() from test #${scope.test} of this file ran after that test ended: its body outlived the test, likely past its timeout`
      );
    }
    return scope;
  };
  const poll = async (scope: Scope, ms: number, call: string) => {
    // React's act thenable must be subscribed once; end() waits on this promise instead.
    const pending = Promise.resolve(act(async () => {
      await new Promise((done) => setTimeout(done, ms));
    }));
    scope.acts.add(pending);
    try {
      await pending;
    } finally {
      scope.acts.delete(pending);
    }
    live(scope, call);
  };

  return {
    async settle(ms = 20) {
      await poll(live(current, 'settle'), ms, 'settle');
    },
    async until(done: () => boolean) {
      const scope = live(current, 'until');
      for (let attempt = 0; attempt < 300 && !done(); attempt += 1) await poll(scope, 10, 'until');
      expect(done()).toBe(true);
    },
  };
}
