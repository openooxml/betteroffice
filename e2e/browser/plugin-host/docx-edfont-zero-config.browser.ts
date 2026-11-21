import { test, expect } from 'playwright/test';
import { documentBytes, expectOneGlyphLinePerParagraph, fontRegistryWarnings } from '../edfont-document';

const standIns = ['Aptos', 'Calibri Light', 'Symbol'];

for (const worker of [true, false]) {
  test(`${worker ? 'worker' : 'in-thread'} editor with no font setup stands in for families it does not carry, one line each`, async ({ page }) => {
    const warnings = fontRegistryWarnings(page);
    await page.goto(`/docx-edfont-zero-config.html?mode=${worker ? 'worker' : 'legacy'}`);
    const bytes = [...await documentBytes()];
    await page.evaluate(data => (window as unknown as { __edfontZeroConfig: { load(bytes: number[]): void } }).__edfontZeroConfig.load(data), bytes);
    await expectOneGlyphLinePerParagraph(page, worker);
    const standingIn = warnings.filter(text => text.includes('stands in'));
    expect(Object.fromEntries(standIns.map(family => [family, standingIn.filter(text => text.includes(`"${family}"`)).length])))
      .toEqual(Object.fromEntries(standIns.map(family => [family, 1])));
    expect(standingIn).toHaveLength(standIns.length);
    expect(warnings.some(text => text.includes('"Calibri"'))).toBe(false);
    expect(await page.evaluate(() => (window as unknown as { __edfontZeroConfig: { errors: string[] } }).__edfontZeroConfig.errors)).toEqual([]);
  });
}
