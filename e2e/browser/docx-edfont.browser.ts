import { test, expect } from 'playwright/test';
import { documentBytes, expectOneGlyphLinePerParagraph, fontRegistryWarnings } from './edfont-document';

for (const worker of [true, false]) {
  test(`${worker ? 'worker' : 'in-thread'} editor draws families it does not carry with a stand-in face on one line each`, async ({ page }) => {
    const warnings = fontRegistryWarnings(page);
    await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
    await page.locator('input[type=file]').last().setInputFiles({ name: 'edfont.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
    await expectOneGlyphLinePerParagraph(page, worker);
    expect(warnings.filter(text => text.includes('stands in'))).toEqual([]);
    expect(warnings.some(text => text.includes('"Calibri"'))).toBe(false);
    await expect(page.locator('.notice')).toHaveCount(0);
  });
}
