import { expect, test, type Page } from 'playwright/test';
import type {} from './xlsx-relay-join-probe';

const EDITS = 200;
const ROWS = 20;
const STEADY = 60;
const writer = (page: Page) => page.locator('[data-editor="0"]');
const value = (row: number, pass: number) => `r${row + 1}p${pass}`;

async function type(page: Page, from: number, count: number) {
  const scroll = writer(page).getByTestId('xlsx-scroll');
  const formula = writer(page).getByTestId('xlsx-formula-input');
  for (let edit = from; edit < from + count; edit++) {
    if (edit % ROWS === 0) { await scroll.press('ControlOrMeta+Home'); await expect(writer(page).getByTestId('xlsx-name-box')).toHaveValue('A1'); }
    await formula.fill(value(edit % ROWS, Math.floor(edit / ROWS)));
    await formula.press('Enter');
  }
  const last = from + count - 1;
  return Array.from({ length: ROWS }, (_, row) => {
    const pass = Math.floor(last / ROWS) - (row > last % ROWS ? 1 : 0);
    return value(row, pass);
  });
}

test.afterEach(async ({ page }) => {
  expect(await page.evaluate(() => window.__xlsxRelayJoin.errors)).toEqual([]);
});

for (const legacy of [false, true]) {
  test(`a ${legacy ? 'main-thread' : 'worker'} xlsx joiner replays a relay room`, async ({ page }) => {
    await page.goto('/xlsx-relay-join.html');
    await page.waitForFunction(() => '__xlsxRelayJoin' in window);
    await page.evaluate(() => window.__xlsxRelayJoin.ready);
    const expected = await type(page, 0, EDITS);
    await expect.poll(() => page.evaluate(() => window.__xlsxRelayJoin.retained())).toBeGreaterThan(EDITS / 2);
    const joined = await page.evaluate(([legacy, expected]) => window.__xlsxRelayJoin.join(legacy, expected), [legacy, expected] as const);
    console.log(`JOIN ${legacy ? 'legacy' : 'worker'} ${JSON.stringify(joined)}`);
    await page.evaluate(() => window.__xlsxRelayJoin.mark());
    const steady = await type(page, EDITS, STEADY);
    const typed = await page.evaluate((expected) => window.__xlsxRelayJoin.watch(expected), steady);
    console.log(`STEADY ${legacy ? 'legacy' : 'worker'} ${JSON.stringify(typed)}`);
    expect(joined.ms).toBeGreaterThan(0);
    expect(typed.ms).toBeGreaterThan(-1);
    if (!legacy) {
      expect(joined.applies).toBe(0);
      expect(joined.pipePosts).toBeGreaterThan(0);
      expect(joined.pipePosts).toBeLessThanOrEqual(4);
      expect(typed.applies).toBe(0);
      expect(typed.pipePosts).toBeGreaterThan(0);
    } else {
      expect(joined.applies).toBeGreaterThan(0);
      expect(joined.pipePosts).toBe(0);
    }
  });
}
