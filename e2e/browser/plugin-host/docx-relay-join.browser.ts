import { expect, test, type Page } from 'playwright/test';
import type {} from './docx-relay-join-probe';

const EDITS = 200;

async function room(page: Page) {
  await page.goto('/docx-relay-join.html');
  await page.waitForFunction(() => '__docxRelayJoin' in window);
  await page.evaluate(() => window.__docxRelayJoin.ready);
  const input = page.locator('[data-editor="0"] textarea.paged-editor__input');
  await input.focus(); await input.press('ControlOrMeta+End');
  for (let edit = 0; edit < EDITS; edit++) {
    await input.evaluate((element, data) => element.dispatchEvent(new InputEvent('beforeinput', {
      bubbles: true, cancelable: true, inputType: 'insertText', data,
    })), String.fromCharCode(97 + (edit % 26)));
    if (edit % 10 === 9) await expect.poll(async () => (await page.evaluate(() => window.__docxRelayJoin.text(0))).length).toBe(5 + edit);
  }
  const expected = await page.evaluate(() => window.__docxRelayJoin.text(0));
  await expect.poll(() => page.evaluate(() => window.__docxRelayJoin.retained())).toBeGreaterThan(EDITS / 2);
  return expected;
}

test.afterEach(async ({ page }) => {
  if (await page.evaluate(() => '__docxRelayJoin' in window)) expect(await page.evaluate(() => window.__docxRelayJoin.errors)).toEqual([]);
});

for (const legacy of [false, true]) {
  test(`a ${legacy ? 'main-thread' : 'worker'} joiner merges a relay room's history in a few applies`, async ({ page }) => {
    const expected = await room(page);
    const result = await page.evaluate(([legacy, expected]) => window.__docxRelayJoin.join(legacy, expected), [legacy, expected] as const);
    expect(result.text).toContain(expected);
    expect(await page.evaluate(() => window.__docxRelayJoin.text(1))).toBe(expected);
    if (!legacy) {
      expect(result.merges).toBeGreaterThan(0);
      expect(result.merges).toBeLessThanOrEqual(4);
    }
  });
}
