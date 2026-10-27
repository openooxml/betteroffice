import { test, expect } from 'playwright/test';

test('the default editor opens a File passed as file and shows fileName', async ({ page }) => {
  await page.goto('/?peerless=1&prop=file');
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  await expect(page.getByLabel('Document name')).toHaveValue('Aliased report');
  const paragraphs = await page.evaluate(async () => {
    const editor = (window as unknown as {
      __probe: { editor: { readParagraphs(request: { view: 'accepted' }): Promise<{ paragraphs: { text: string }[] }> } };
    }).__probe.editor;
    return (await editor.readParagraphs({ view: 'accepted' })).paragraphs.map(({ text }) => text);
  });
  expect(paragraphs.some((text) => text.trim().length > 0)).toBe(true);
});
