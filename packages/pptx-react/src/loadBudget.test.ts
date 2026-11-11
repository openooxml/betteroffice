import { expect, test } from 'bun:test';
import { BudgetExceededError, budgetedTest, budgetedWaitFor, pollUntil, type TestRegister } from '../../../shared/office-session/testing/loadBudget';

const pause = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

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

test('budgetedTest registers the budget as the timeout, keeps an explicit one, and passes rows through', () => {
  const registered: Array<[string, number]> = [];
  const rows: unknown[][] = [];
  const register: TestRegister = Object.assign(
    (name: string, _body: unknown, timeoutMs: number) => { registered.push([name, timeoutMs]); },
    {
      each: (cases: readonly unknown[]) => (name: string, _body: unknown, timeoutMs: number) => {
        rows.push([...cases]);
        registered.push([name, timeoutMs]);
      },
    },
  );
  const budgeted = budgetedTest(40, register);
  const body = async () => {};
  budgeted('default', body);
  budgeted('explicit', body, 30_000);
  budgeted.each([false, true])('row %s', body);
  expect(registered).toEqual([['default', 40], ['explicit', 30_000], ['row %s', 40]]);
  expect(rows).toEqual([[false, true]]);
});
