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

function gate() {
  let fail!: (error: Error) => void;
  const opened = new Promise<void>((_resolve, reject) => { fail = reject; });
  return { opened, fail };
}

test('previousTestSettled fails naming the budget when the test before it never settles, and later tests fail until it does', async () => {
  const { bodies, tests } = capturedTests(20);
  const stuck = gate();
  tests.test('never settles', () => stuck.opened);
  tests.test('next', async () => {});
  tests.test('after', async () => {});
  void bodies[0]().catch(() => {});
  await expect(tests.previousTestSettled()).rejects.toThrow('still running 20 ms after its 20 ms budget ran out');
  expect(() => bodies[1]()).toThrow('"never settles" is still running after its teardown budget ran out');
  stuck.fail(new Error('late failure'));
  await pause(0);
  expect(() => bodies[2]()).toThrow('"never settles" failed after its teardown budget ran out: Error: late failure');
});

test('previousTestSettled waits for late work and throws its late failure', async () => {
  const { bodies, tests } = capturedTests(20);
  const late = gate();
  tests.test('late failure', () => late.opened);
  void bodies[0]().catch(() => {});
  await pause(30);
  const settled = tests.previousTestSettled();
  late.fail(new Error('late failure'));
  await expect(settled).rejects.toThrow('late failure');
});

test('previousTestSettled throws a late failure that settled before the teardown began', async () => {
  const { bodies, tests } = capturedTests(20);
  const late = gate();
  tests.test('late failure', () => late.opened);
  void bodies[0]().catch(() => {});
  await pause(30);
  late.fail(new Error('late failure'));
  await pause(0);
  await expect(tests.previousTestSettled()).rejects.toThrow('late failure');
});

test('previousTestSettled leaves a failure inside the budget to the test itself', async () => {
  const { bodies, tests } = capturedTests(1_000);
  tests.test('fails in time', async () => { throw new Error('in time'); });
  await expect(bodies[0]()).rejects.toThrow('in time');
  await tests.previousTestSettled();
});

test('previousTestSettled passes after a test that settled in time', async () => {
  const { bodies, tests } = capturedTests(1_000);
  tests.test('settled', async () => {});
  await bodies[0]();
  await tests.previousTestSettled();
});
