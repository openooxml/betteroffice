import { expect, test } from 'playwright/test';
import type {} from './pptx-relay-join-probe';

const EDITS = 200;

test.afterEach(async ({ page }) => {
  expect(await page.evaluate(() => window.__pptxRelayJoin.errors)).toEqual([]);
});

for (const legacy of [false, true]) {
  test(`a ${legacy ? 'main-thread' : 'worker'} pptx joiner replays a relay room`, async ({ page }) => {
    await page.goto('/pptx-relay-join.html');
    await page.waitForFunction(() => '__pptxRelayJoin' in window);
    await page.evaluate(() => window.__pptxRelayJoin.ready);
    const expected = await page.evaluate((count) => window.__pptxRelayJoin.edit(count), EDITS);
    await expect.poll(() => page.evaluate(() => window.__pptxRelayJoin.retained())).toBeGreaterThan(EDITS / 2);
    const joined = await page.evaluate(([legacy, expected]) => window.__pptxRelayJoin.join(legacy, expected), [legacy, expected] as const);
    console.log(`JOIN ${legacy ? 'legacy' : 'worker'} ${JSON.stringify(joined)}`);
    expect(joined.ms).toBeGreaterThan(0);
    if (!legacy) {
      expect(joined.applies).toBe(0);
      expect(joined.pipePosts).toBeGreaterThan(0);
      expect(joined.pipePosts).toBeLessThan(20);
      expect(joined.paints).toBeLessThan(20);
    } else {
      expect(joined.applies).toBeGreaterThan(0);
      expect(joined.pipePosts).toBe(0);
    }
  });
}
