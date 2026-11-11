import { describe, expect, test } from 'bun:test';
import { withEngineCleanup } from './trap';

function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected a throw');
}

describe('withEngineCleanup', () => {
  test('first error wins, a cleanup error alone still surfaces', () => {
    const runError = new Error('run');
    const cleanupError = new Error('cleanup');
    expect(thrown(() => withEngineCleanup(() => { throw runError; }, () => { throw cleanupError; })))
      .toBe(runError);
    expect(thrown(() => withEngineCleanup(() => 1, () => { throw cleanupError; }))).toBe(cleanupError);
  });

  test('a failing run surfaces its own error after the cleanup ran', () => {
    const runError = new Error('run');
    let cleanups = 0;
    expect(thrown(() => withEngineCleanup(() => { throw runError; }, () => { cleanups += 1; })))
      .toBe(runError);
    expect(cleanups).toBe(1);
  });

  test('a successful run returns its result after the cleanup ran once', () => {
    const value = { result: true };
    let cleanups = 0;
    expect(withEngineCleanup(() => value, () => { cleanups += 1; })).toBe(value);
    expect(cleanups).toBe(1);
  });

  test('a returned promise is passed through as the result, with the cleanup run at once', async () => {
    const order: string[] = [];
    const rejection = new Error('later');
    const promise = withEngineCleanup(
      () => Promise.resolve().then(() => { order.push('settled'); throw rejection; }),
      () => { order.push('cleanup'); },
    );
    expect(order).toEqual(['cleanup']);
    await expect(promise).rejects.toBe(rejection);
    expect(order).toEqual(['cleanup', 'settled']);
  });
});
