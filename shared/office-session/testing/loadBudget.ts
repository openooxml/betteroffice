import { afterAll, beforeEach, test } from 'bun:test';

const unmockedSetTimeout = globalThis.setTimeout;
const unmockedClearTimeout = globalThis.clearTimeout;

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); a wait inside one gets about 2.2x that. */
export const EDITOR_WAIT_BUDGET_MS = 15_000;

/** @internal Under 24 busy loops, outside a stalled round, a whole editor test took at most 6682 ms (p90 at most 4249 ms); one test gets about 3x that. */
export const EDITOR_TEST_BUDGET_MS = 20_000;

/** @internal A teardown waits one test budget for a test that ran out of its own, plus one for the file's cleanup. */
export const EDITOR_TEARDOWN_BUDGET_MS = 2 * EDITOR_TEST_BUDGET_MS;

/** @internal Bun's own test timeout sits this far past a budget, so the budget's timer, which keeps the late work tracked, fires first. */
export const RUNNER_TIMEOUT_MARGIN_MS = 5_000;

/** @internal Thrown when a budgeted wait or a timed-out test's late work outlasts its budget. */
export class BudgetExceededError extends Error {
  override name = 'BudgetExceededError';
}

type Run = {
  name: string;
  promise: Promise<void>;
  timedOut: boolean;
  settled: boolean;
  abandoned: boolean;
  failed?: boolean;
  error?: unknown;
};
type Body = (...args: any[]) => Promise<void> | void;
/** @internal The runner functions budgetedTests registers through. */
export type BudgetHooks = {
  test: {
    (name: string, body: () => Promise<void>, timeoutMs: number): void;
    each(cases: readonly unknown[]): (name: string, body: (...args: any[]) => Promise<void>, timeoutMs: number) => void;
  };
  beforeEach(body: () => void): void;
  afterAll(body: () => void): void;
};

/**
 * @internal Tests that each fail once their budget runs out, and a teardown step that waits for such a test.
 * The timed-out body keeps running; the teardown waits up to one budget for it and throws its late
 * failure there. If it is still running then, every later test of this set fails before its setup
 * until it settles, and a failure it ends with is thrown once after the last test.
 */
export function budgetedTests(
  budgetMs: number = EDITOR_TEST_BUDGET_MS,
  hooks: BudgetHooks = { test: test as unknown as BudgetHooks['test'], beforeEach, afterAll },
) {
  let current: Run | undefined;
  let abandoned: Run | undefined;
  const lateFailures: Run[] = [];
  hooks.beforeEach(() => {
    if (abandoned && !abandoned.settled) {
      throw new BudgetExceededError(`"${abandoned.name}" is still running after its teardown wait of ${budgetMs} ms ran out`);
    }
    abandoned = undefined;
  });
  hooks.afterAll(() => {
    if (lateFailures.length === 0) return;
    const names = lateFailures.map((run) => `"${run.name}": ${String(run.error)}`).join('; ');
    throw Object.assign(new BudgetExceededError(`failed after the teardown wait ran out: ${names}`), { cause: lateFailures[0].error });
  });
  const track = (name: string, start: () => Promise<void> | void, timeoutMs: number) => {
    const run: Run = { name, promise: Promise.resolve(start()), timedOut: false, settled: false, abandoned: false };
    run.promise.then(
      () => { run.settled = true; },
      (error: unknown) => {
        run.settled = true;
        run.failed = true;
        run.error = error;
        if (run.abandoned) lateFailures.push(run);
      },
    );
    current = run;
    let timer: ReturnType<typeof setTimeout> | undefined;
    return new Promise<void>((resolve, reject) => {
      timer = unmockedSetTimeout(() => {
        if (run.settled) return;
        run.timedOut = true;
        reject(new BudgetExceededError(`"${name}" ran past its ${timeoutMs} ms budget`));
      }, timeoutMs);
      run.promise.then(resolve, reject);
    }).finally(() => unmockedClearTimeout(timer));
  };
  const budgeted = (name: string, body: () => Promise<void> | void, timeoutMs: number = budgetMs) => {
    hooks.test(name, () => track(name, body, timeoutMs), timeoutMs + RUNNER_TIMEOUT_MARGIN_MS);
  };
  budgeted.each = (cases: readonly unknown[]) =>
    (name: string, body: Body, timeoutMs: number = budgetMs) => {
      hooks.test.each(cases)(name, (...args: unknown[]) => track(`${name} ${JSON.stringify(args)}`, () => body(...args), timeoutMs), timeoutMs + RUNNER_TIMEOUT_MARGIN_MS);
    };
  return {
    test: budgeted,
    async previousTestSettled() {
      const run = current;
      current = undefined;
      if (!run?.timedOut) return;
      if (!run.settled) {
        let timer: ReturnType<typeof setTimeout> | undefined;
        const settled = await Promise.race([
          run.promise.then(() => true, () => true),
          new Promise<false>((resolve) => { timer = unmockedSetTimeout(() => resolve(false), budgetMs); }),
        ]);
        unmockedClearTimeout(timer);
        if (!settled) {
          run.abandoned = true;
          abandoned = run;
          throw new BudgetExceededError(`"${run.name}" was still running ${budgetMs} ms after its budget ran out`);
        }
      }
      if (run.failed) throw run.error;
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
