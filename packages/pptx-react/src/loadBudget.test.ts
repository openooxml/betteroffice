import { expect, test } from 'bun:test';
import { BudgetExceededError, budgetedTests, budgetedWaitFor, pollUntil } from '../../../shared/office-session/testing/loadBudget';

const pause = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

function capturedTests(budgetMs: number) {
  const bodies: Array<() => Promise<void>> = [];
  const register = Object.assign(
    (_name: string, body: () => Promise<void>) => { bodies.push(body); },
    { each: () => () => {} },
  );
  return { bodies, tests: budgetedTests(budgetMs, register) };
}

test('pollUntil fails naming its budget when the condition never holds', async () => {
  const waited = pollUntil(() => false, pause, 40, 5);
  await expect(waited).rejects.toBeInstanceOf(BudgetExceededError);
  await expect(pollUntil(() => false, pause, 40, 5)).rejects.toThrow('after its 40 ms wait budget');
});

test('pollUntil returns once the condition holds', async () => {
  let polls = 0;
  await pollUntil(() => ++polls > 3, pause, 1_000, 1);
  expect(polls).toBe(4);
});

test('budgetedWaitFor fails naming its budget and keeps the waited error', async () => {
  const seen: Array<number | undefined> = [];
  const waitFor = budgetedWaitFor(async <T>(_callback: () => T | Promise<T>, options?: { timeout?: number }): Promise<T> => {
    seen.push(options?.timeout);
    throw new Error('expected selection focus 4');
  }, 40);
  await expect(waitFor(() => false)).rejects.toThrow('wait still unmet after its 40 ms budget: expected selection focus 4');
  await expect(waitFor(() => false, { timeout: 7 })).rejects.toThrow('after its 7 ms budget');
  expect(seen).toEqual([40, 7]);
});

test('previousTestSettled fails naming the budget when the test before it never settles', async () => {
  const { bodies, tests } = capturedTests(40);
  tests.test('never settles', () => new Promise<void>(() => {}));
  void bodies[0]();
  await expect(tests.previousTestSettled()).rejects.toThrow('still running 40 ms after its 40 ms budget ran out');
});

test('previousTestSettled waits for late work and throws its late failure', async () => {
  const { bodies, tests } = capturedTests(1_000);
  let finished = false;
  tests.test('late failure', async () => {
    await pause(30);
    finished = true;
    throw new Error('late failure');
  });
  void bodies[0]().catch(() => {});
  await expect(tests.previousTestSettled()).rejects.toThrow('late failure');
  expect(finished).toBe(true);
});

test('previousTestSettled passes after a test that settled in time', async () => {
  const { bodies, tests } = capturedTests(1_000);
  tests.test('settled', async () => {});
  await bodies[0]();
  await tests.previousTestSettled();
});
