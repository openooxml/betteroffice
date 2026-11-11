import { expect, test } from 'bun:test';
import { BudgetExceededError, budgetedTests, budgetedWaitFor, pollUntil } from '../../../shared/office-session/testing/loadBudget';

const pause = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

function capturedTests(budgetMs: number) {
  const bodies: Array<() => Promise<void>> = [];
  const setups: Array<() => void> = [];
  const finals: Array<() => void> = [];
  const register = Object.assign(
    (_name: string, body: () => Promise<void>) => { bodies.push(body); },
    { each: () => () => {} },
  );
  const tests = budgetedTests(budgetMs, {
    test: register,
    beforeEach: (body) => { setups.push(body); },
    afterAll: (body) => { finals.push(body); },
  });
  return { bodies, tests, setup: () => setups[0](), finish: () => finals[0]() };
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
  let open!: () => void;
  let fail!: (error: Error) => void;
  const opened = new Promise<void>((resolve, reject) => { open = resolve; fail = reject; });
  return { opened, open, fail };
}

test('a budgeted test fails naming its budget when its body never settles', async () => {
  const { bodies, tests } = capturedTests(20);
  tests.test('never settles', () => gate().opened);
  await expect(bodies[0]()).rejects.toThrow('"never settles" ran past its 20 ms budget');
});

test('a teardown that gives up fails later setups until the test settles, and its late failure is thrown once after the last test', async () => {
  const { bodies, tests, setup, finish } = capturedTests(20);
  const stuck = gate();
  tests.test('stuck', () => stuck.opened);
  await expect(bodies[0]()).rejects.toThrow('ran past its 20 ms budget');
  await expect(tests.previousTestSettled()).rejects.toThrow('"stuck" was still running 20 ms after its budget ran out');
  expect(setup).toThrow('"stuck" is still running after its teardown wait of 20 ms ran out');
  stuck.fail(new Error('late failure'));
  await pause(0);
  expect(setup).not.toThrow();
  expect(finish).toThrow('failed after the teardown wait ran out: "stuck": Error: late failure');
});

test('a test that finishes during the teardown wait blocks nothing', async () => {
  const { bodies, tests, setup, finish } = capturedTests(20);
  const slow = gate();
  tests.test('slow', () => slow.opened);
  await expect(bodies[0]()).rejects.toThrow('ran past its 20 ms budget');
  const settled = tests.previousTestSettled();
  slow.open();
  await settled;
  expect(setup).not.toThrow();
  expect(finish).not.toThrow();
});

test('previousTestSettled waits for late work and throws its late failure', async () => {
  const { bodies, tests, finish } = capturedTests(20);
  const late = gate();
  tests.test('late failure', () => late.opened);
  await expect(bodies[0]()).rejects.toThrow('ran past its 20 ms budget');
  const settled = tests.previousTestSettled();
  late.fail(new Error('late failure'));
  await expect(settled).rejects.toThrow('late failure');
  expect(finish).not.toThrow();
});

test('previousTestSettled throws a late failure that settled before the teardown began', async () => {
  const { bodies, tests } = capturedTests(20);
  const late = gate();
  tests.test('late failure', () => late.opened);
  await expect(bodies[0]()).rejects.toThrow('ran past its 20 ms budget');
  late.fail(new Error('late failure'));
  await pause(0);
  await expect(tests.previousTestSettled()).rejects.toThrow('late failure');
});

test('previousTestSettled leaves a failure inside the budget to the test itself', async () => {
  const { bodies, tests, finish } = capturedTests(1_000);
  tests.test('fails in time', async () => { throw new Error('in time'); });
  await expect(bodies[0]()).rejects.toThrow('in time');
  await tests.previousTestSettled();
  expect(finish).not.toThrow();
});

test('previousTestSettled passes after a test that settled in time', async () => {
  const { bodies, tests, setup } = capturedTests(1_000);
  tests.test('settled', async () => {});
  await bodies[0]();
  await tests.previousTestSettled();
  expect(setup).not.toThrow();
});
