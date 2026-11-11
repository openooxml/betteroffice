import { test } from 'bun:test';

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); a wait inside one gets about 2.2x that. */
export const EDITOR_WAIT_BUDGET_MS = 15_000;

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); one test gets about 3x that. */
export const EDITOR_TEST_BUDGET_MS = 20_000;

/** @internal A teardown waits one test budget for a test that ran out of its own, plus one for the file's cleanup. */
export const EDITOR_TEARDOWN_BUDGET_MS = 2 * EDITOR_TEST_BUDGET_MS;

/** @internal Thrown when a budgeted wait or a timed-out test's late work outlasts its budget. */
export class BudgetExceededError extends Error {
  override name = 'BudgetExceededError';
}

type Run = { promise: Promise<void>; budgetMs: number; settled: boolean };
type Register = {
  (name: string, body: () => Promise<void>, timeoutMs: number): void;
  each<T>(cases: T[]): (name: string, body: (value: T) => Promise<void>, timeoutMs: number) => void;
};

/**
 * @internal Tests that run under one budget, and a teardown step that waits for the test it follows.
 * A test that runs out of budget keeps running; the teardown waits for it, so its late work cannot
 * land in the next test, and a late failure is thrown there instead of being dropped.
 */
export function budgetedTests(budgetMs: number = EDITOR_TEST_BUDGET_MS, register: Register = test as unknown as Register) {
  let current: Run | undefined;
  const track = (body: () => Promise<void> | void, runBudgetMs: number) => {
    const run: Run = { promise: Promise.resolve(body()), budgetMs: runBudgetMs, settled: false };
    run.promise.then(() => { run.settled = true; }, () => { run.settled = true; });
    current = run;
    return run.promise;
  };
  const budgeted = (name: string, body: () => Promise<void> | void, timeoutMs: number = budgetMs) => {
    register(name, () => track(body, timeoutMs), timeoutMs);
  };
  budgeted.each = <T>(cases: readonly T[]) =>
    (name: string, body: (value: T) => Promise<void> | void, timeoutMs: number = budgetMs) => {
      register.each(cases as T[])(name, (value: T) => track(() => body(value), timeoutMs), timeoutMs);
    };
  return {
    test: budgeted,
    async previousTestSettled() {
      const run = current;
      current = undefined;
      if (!run || run.settled) return;
      let timer: ReturnType<typeof setTimeout> | undefined;
      const late = await Promise.race([
        run.promise.then(() => ({ done: true as const }), (error: unknown) => ({ done: true as const, error })),
        new Promise<{ done: false }>((resolve) => { timer = setTimeout(() => resolve({ done: false }), run.budgetMs); }),
      ]);
      clearTimeout(timer);
      if (!late.done) {
        throw new BudgetExceededError(`the test before this teardown was still running ${run.budgetMs} ms after its ${run.budgetMs} ms budget ran out`);
      }
      if ('error' in late) throw late.error;
    },
  };
}

type WaitOptions = { timeout?: number };
type WaitFor<O extends WaitOptions> = <T>(callback: () => T | Promise<T>, options?: O) => Promise<T>;

/**
 * @internal The given waitFor with the wait budget as its default timeout; an explicit timeout is kept.
 * A wait that runs out rethrows the same error with the budget at the front of its message.
 */
export function budgetedWaitFor<O extends WaitOptions>(waitFor: WaitFor<O>, budgetMs: number = EDITOR_WAIT_BUDGET_MS): WaitFor<O> {
  return async (callback, options) => {
    const timeout = options?.timeout ?? budgetMs;
    try {
      return await waitFor(callback, { ...options, timeout } as O);
    } catch (error) {
      if (error instanceof Error) error.message = `wait still unmet after its ${timeout} ms budget: ${error.message}`;
      throw error;
    }
  };
}

/** @internal Polls done every stepMs until it holds; throws naming the budget when budgetMs passes first. */
export async function pollUntil(
  done: () => boolean,
  step: (ms: number) => Promise<void>,
  budgetMs: number = EDITOR_WAIT_BUDGET_MS,
  stepMs = 10,
): Promise<void> {
  const deadline = performance.now() + budgetMs;
  while (!done()) {
    if (performance.now() >= deadline) {
      throw new BudgetExceededError(`condition still false after its ${budgetMs} ms wait budget: ${done.toString()}`);
    }
    await step(stepMs);
  }
}
