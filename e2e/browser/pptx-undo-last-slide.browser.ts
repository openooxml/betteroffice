import { test, expect, type Page } from 'playwright/test';
import { resolve } from 'node:path';
import { painted, slideCanvas } from './pptx-async-helpers';

const root = resolve(import.meta.dirname, '../..');
const thumbnails = (page: Page) => page.locator('aside[aria-label="Slides"] button');

async function openDeck(page: Page) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.pptx'));
  await expect(slideCanvas(page)).toBeVisible();
  await expect(thumbnails(page).first()).toHaveAttribute('aria-current', 'page');
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
}

async function deleteActiveSlide(page: Page, remaining: number) {
  const before = await painted(page);
  await page.getByRole('button', { name: 'Delete slide', exact: true }).click();
  await expect(thumbnails(page)).toHaveCount(remaining);
  if (remaining > 0) await expect.poll(() => painted(page)).not.toBe(before);
}

async function undo(page: Page) {
  await page.getByTestId('pptx-undo').focus();
  await page.keyboard.press('Enter');
}

test('pptx: undoing the deletion of the only slide paints it again', async ({ page }) => {
  test.setTimeout(60_000);
  await openDeck(page);
  for (let count = await thumbnails(page).count(); count > 1; count--) await deleteActiveSlide(page, count - 1);
  const last = await painted(page);
  await deleteActiveSlide(page, 0);
  await expect(page.getByTestId('pptx-worker-status')).toHaveText('No slides');
  await undo(page);
  await expect(thumbnails(page)).toHaveCount(1);
  await expect(slideCanvas(page)).toBeVisible();
  await expect.poll(() => painted(page)).toBe(last);
});
