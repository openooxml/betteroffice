import { test } from 'bun:test';

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); a wait inside one gets about 2.2x that. */
export const EDITOR_WAIT_BUDGET_MS = 15_000;

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); one test gets about 3x that. */
export const EDITOR_TEST_BUDGET_MS = 20_000;

/** @internal Thrown when a budgeted wait runs out before its condition holds. */
export class BudgetExceededError extends Error {
  override name = 'BudgetExceededError';
}

type Body = (...args: any[]) => Promise<void> | void;
/** @internal The runner function budgetedTest registers through. */
export type TestRegister = {
  (name: string, body: Body, timeoutMs: number): void;
  each(cases: readonly unknown[]): (name: string, body: Body, timeoutMs: number) => void;
};

/** @internal The runner's test with the test budget as its default timeout; an explicit timeout is kept. */
export function budgetedTest(budgetMs: number = EDITOR_TEST_BUDGET_MS, register: TestRegister = test as unknown as TestRegister) {
  const budgeted = (name: string, body: Body, timeoutMs: number = budgetMs) => register(name, body, timeoutMs);
  budgeted.each = (cases: readonly unknown[]) =>
    (name: string, body: Body, timeoutMs: number = budgetMs) => register.each(cases)(name, body, timeoutMs);
  return budgeted;
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
