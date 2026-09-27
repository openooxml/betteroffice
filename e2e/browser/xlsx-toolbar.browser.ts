import { test, expect, type Page } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');

async function open(page: Page) {
  await page.goto('/?format=xlsx');
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(resolve(root, 'apps/demo/public/showcase.xlsx'));
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
}

test('xlsx: a narrow toolbar keeps every control in a keyboard menu inside the viewport', async ({
  page,
}) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);

  const rail = page.getByTestId('xlsx-formatting-toolbar');
  const more = rail.getByTestId('xlsx-toolbar-more');
  await expect(more).toBeVisible();
  const pageWidth = await page.evaluate(() => document.documentElement.scrollWidth);
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More toolbar items' });
  await expect(menu).toBeVisible();
  const box = (await menu.boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(360);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(pageWidth);
  for (const label of ['Bold', 'Borders', 'Horizontal alignment', 'Text wrapping']) {
    await expect(menu.locator(`[data-label="${label}"]`)).toHaveCount(1);
  }
  const wasBold = await menu.locator('[data-label="Bold"]').getAttribute('aria-checked');

  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Text wrapping');
  await page.keyboard.press('ArrowRight');
  const wrapping = page.getByRole('menu', { name: 'Text wrapping' });
  await expect(wrapping).toBeVisible();
  await page.keyboard.press('ArrowLeft');
  await expect(wrapping).toBeHidden();
  await page.keyboard.press('Home');
  await page.keyboard.type('b');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Bold');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();

  await page.keyboard.press('ArrowDown');
  await expect(menu.locator('[data-label="Bold"]')).toHaveAttribute(
    'aria-checked',
    wasBold === 'true' ? 'false' : 'true'
  );
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();
});

test('xlsx: a toolbar command pressed during IME composition lands after the composed text', async ({
  page,
}) => {
  await open(page);
  const bold = page.getByTestId('xlsx-formatting-toolbar').getByRole('button', { name: 'Bold' });
  const wasBold = await bold.getAttribute('aria-pressed');
  await page.getByTestId('xlsx-scroll').press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await input.press('ControlOrMeta+A');
  await input.pressSequentially('ab');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', {
    text: 'かな',
    selectionStart: 2,
    selectionEnd: 2,
  });
  await bold.click();

  await expect(input).toHaveCount(0);
  await expect(page.getByRole('gridcell').filter({ hasText: 'abかな' })).toHaveCount(1);
  await expect(bold).toHaveAttribute('aria-pressed', wasBold === 'true' ? 'false' : 'true');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('abかな');
});
