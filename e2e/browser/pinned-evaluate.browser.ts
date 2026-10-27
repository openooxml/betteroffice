import { expect, test } from 'playwright/test';
import { pinnedEvaluate } from './pinned-evaluate';

test.use({ launchOptions: { args: ['--js-flags=--expose-gc'] } });

test('pins a settled promise until its value is read', async ({ page }) => {
  expect(await pinnedEvaluate(page, (value) => {
    const box: { resolve?: (value: number) => void } = {};
    const promise = new Promise<number>((resolve) => { box.resolve = resolve; });
    const settle = () => {
      const resolve = box.resolve!;
      box.resolve = undefined;
      resolve(value);
    };
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.onmessage = null;
      channel.port1.close();
      channel.port2.close();
      settle();
      const collect = () => (globalThis as unknown as { gc(): void }).gc();
      collect();
      queueMicrotask(collect);
    };
    channel.port2.postMessage(null);
    return promise;
  }, 7)).toBe(7);
});

test('serializes synchronous and asynchronous functions with unboxed arguments', async ({ page }) => {
  const arg = await page.evaluateHandle(() => ({ value: 7 }));
  try {
    expect(await pinnedEvaluate(page, ({ value }) => value, arg)).toBe(7);
    expect(await pinnedEvaluate(page, async ({ value }) => ({ value, missing: undefined }), arg))
      .toEqual({ value: 7, missing: undefined });
    expect(await pinnedEvaluate(page, function ({ value }) { return Promise.resolve(value); }, arg)).toBe(7);
    expect(await pinnedEvaluate(page, () => ({ date: new Date(0), nan: NaN, negativeZero: -0, infinity: Infinity, bigint: 7n })))
      .toEqual({ date: new Date(0), nan: NaN, negativeZero: -0, infinity: Infinity, bigint: 7n });
  } finally {
    await arg.dispose();
  }
});

test('propagates synchronous throws and rejected promises', async ({ page }) => {
  await expect(pinnedEvaluate(page, () => { throw new Error('synchronous failure'); })).rejects.toThrow('synchronous failure');
  await expect(pinnedEvaluate(page, async () => { throw new Error('asynchronous failure'); })).rejects.toThrow('asynchronous failure');
  expect(await pinnedEvaluate(page, () => undefined)).toBeUndefined();
});
