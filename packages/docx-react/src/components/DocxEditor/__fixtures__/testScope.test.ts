import { afterEach, expect, mock, test } from 'bun:test';
import { testScope } from './testScope';

let runCallbacks = true;
const act = mock(async (callback: () => Promise<void>) => {
  if (runCallbacks) await callback();
});
const { settle, until } = testScope(act);

const outcome = (call: Promise<void>) =>
  call.then(
    () => null,
    (error: unknown) => (error instanceof Error ? Bun.stripANSI(error.message) : String(error))
  );

let ended: { acts: number; outcomes: Array<string | null> } | null = null;
let actsAtEnd = 0;
afterEach(async () => {
  const acts = act.mock.calls.length;
  const outcomes = await Promise.all([outcome(settle(1)), outcome(until(() => true))]);
  actsAtEnd = act.mock.calls.length;
  ended = { acts: actsAtEnd - acts, outcomes };
});

test('settle and until act while their test is live', async () => {
  act.mockClear();
  await settle();
  expect(act).toHaveBeenCalledTimes(1);
  let checks = 0;
  await until(() => ++checks > 3);
  expect(act).toHaveBeenCalledTimes(4);
});

test('the scope ends before later afterEach hooks, which then cannot act through it', () => {
  expect(ended).toEqual({
    acts: 0,
    outcomes: [
      expect.stringMatching(/^settle\(\) from test #1 of this file ran after that test ended/),
      expect.stringMatching(/^until\(\) from test #1 of this file ran after that test ended/),
    ],
  });
});

let leftover: Promise<string | null> = Promise.resolve(null);
test('a body can leave an until loop polling when its test ends', () => {
  const acts = act.mock.calls.length;
  leftover = outcome(until(() => false));
  expect(act.mock.calls.length).toBe(acts + 1);
});

test('that loop rejects at its next poll instead of acting in the next test', async () => {
  expect(await leftover).toStartWith('until() from test #3 of this file ran after that test ended');
  expect(act.mock.calls.length).toBe(actsAtEnd);
});

test('until still fails on its expect when done() stays false', async () => {
  runCallbacks = false;
  try {
    const acts = act.mock.calls.length;
    const failure = await outcome(until(() => false));
    expect(act.mock.calls.length - acts).toBe(300);
    expect(failure).toContain('Received: false');
    expect(failure).not.toContain('ran after that test ended');
  } finally {
    runCallbacks = true;
  }
});
