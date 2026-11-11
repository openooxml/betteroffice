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

type Outcome = { late: boolean; failed: boolean; error?: unknown };
type Run = { name: string; promise: Promise<void>; budgetMs: number; outcome?: Outcome };
type Body = (...args: any[]) => Promise<void> | void;
type Register = {
  (name: string, body: () => Promise<void>, timeoutMs: number): void;
  each(cases: readonly unknown[]): (name: string, body: (...args: any[]) => Promise<void>, timeoutMs: number) => void;
};

/**
 * @internal Tests that run under one budget, and a teardown step that waits for the test it follows.
 * A test that runs out of budget keeps running; the teardown waits for it, so its late work cannot
 * land in the next test, and a late failure is thrown there instead of being dropped. A test the
 * teardown gave up on fails every later test of this set until it settles.
 */
export function budgetedTests(budgetMs: number = EDITOR_TEST_BUDGET_MS, register: Register = test as unknown as Register) {
  let current: Run | undefined;
  let abandoned: Run | undefined;
  const track = (name: string, start: () => Promise<void> | void, runBudgetMs: number) => {
    if (abandoned) {
      const earlier = abandoned;
      if (!earlier.outcome) {
        throw new BudgetExceededError(`"${earlier.name}" is still running after its teardown budget ran out`);
      }
      abandoned = undefined;
      if (earlier.outcome.failed) {
        const error = new BudgetExceededError(`"${earlier.name}" failed after its teardown budget ran out: ${String(earlier.outcome.error)}`);
        throw Object.assign(error, { cause: earlier.outcome.error });
      }
    }
    const startedAt = performance.now();
    const run: Run = { name, promise: Promise.resolve(start()), budgetMs: runBudgetMs };
    const settle = (failed: boolean, error?: unknown) => {
      run.outcome = { late: performance.now() - startedAt >= runBudgetMs, failed, error };
    };
    run.promise.then(() => settle(false), (error: unknown) => settle(true, error));
    current = run;
    return run.promise;
  };
  const budgeted = (name: string, body: () => Promise<void> | void, timeoutMs: number = budgetMs) => {
    register(name, () => track(name, body, timeoutMs), timeoutMs);
  };
  budgeted.each = (cases: readonly unknown[]) =>
    (name: string, body: Body, timeoutMs: number = budgetMs) => {
      register.each(cases)(name, (...args: unknown[]) => track(name, () => body(...args), timeoutMs), timeoutMs);
    };
  return {
    test: budgeted,
    async previousTestSettled() {
      const run = current;
      current = undefined;
      if (!run) return;
      if (!run.outcome) {
        let timer: ReturnType<typeof setTimeout> | undefined;
        const settled = await Promise.race([
          run.promise.then(() => true, () => true),
          new Promise<false>((resolve) => { timer = setTimeout(() => resolve(false), run.budgetMs); }),
        ]);
        clearTimeout(timer);
        if (!settled) {
          abandoned = run;
          throw new BudgetExceededError(`the test before this teardown was still running ${run.budgetMs} ms after its ${run.budgetMs} ms budget ran out`);
        }
      }
      if (run.outcome?.late && run.outcome.failed) throw run.outcome.error;
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
