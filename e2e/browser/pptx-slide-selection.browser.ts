import { test, expect, type Page } from 'playwright/test';
import { resolve } from 'node:path';
import { painted, expectRepaint, slideCanvas } from './pptx-async-helpers';

const root = resolve(import.meta.dirname, '../..');
const thumbnails = (page: Page) => page.locator('aside[aria-label="Slides"] button');
const selectionBox = (page: Page) => page.getByTestId('pptx-shape-selection');

async function openDeck(page: Page) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.pptx'));
  await expect(slideCanvas(page)).toBeVisible();
  await expect(thumbnails(page).first()).toHaveAttribute('aria-current', 'page');
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
}

async function selectOnFirstSlide(page: Page, group: boolean) {
  const box = (await slideCanvas(page).boundingBox())!;
  await page.mouse.click(box.x + box.width * (group ? 0.59 : 0.009375), box.y + box.height * 0.5);
  await expect(selectionBox(page)).toBeVisible();
  await expect(page.getByTestId('pptx-resize-se')).toBeVisible();
  if (group) await expect.poll(async () => (await selectionBox(page).boundingBox())?.width ?? 0)
    .toBeGreaterThan(box.width * 0.3);
}

async function expectNoSelection(page: Page) {
  await expect(selectionBox(page)).toHaveCount(0);
  await expect(page.locator('[data-testid^="pptx-resize-"]')).toHaveCount(0);
  await expect(page.getByTestId('pptx-text-edit-outline')).toHaveCount(0);
  await expect(page.getByTestId('pptx-peerless-caret')).toBeHidden();
  await expect(page.getByTestId('pptx-shape-fill')).toBeDisabled();
}

for (const group of [false, true]) {
  test(`pptx: New slide drops the previous slide's ${group ? 'group' : 'shape'} selection`, async ({ page }) => {
    test.setTimeout(30_000);
    await openDeck(page);
    await selectOnFirstSlide(page, group);
    const before = await painted(page);
    await page.getByRole('button', { name: 'New slide', exact: true }).click();
    await expectRepaint(page, before);
    await expect(thumbnails(page).nth(1)).toHaveAttribute('aria-current', 'page');
    await expectNoSelection(page);
  });
}

test('pptx: Delete after New slide leaves the previous slide untouched', async ({ page }) => {
  test.setTimeout(30_000);
  await openDeck(page);
  const original = await painted(page);
  const count = await thumbnails(page).count();
  await selectOnFirstSlide(page, false);
  await page.getByRole('button', { name: 'New slide', exact: true }).click();
  await expectRepaint(page, original);
  await expect(thumbnails(page)).toHaveCount(count + 1);
  await page.keyboard.press('Delete');
  await page.keyboard.press('Backspace');
  await thumbnails(page).first().click();
  await expect(thumbnails(page)).toHaveCount(count + 1);
  await expect.poll(async () => await painted(page) === original).toBe(true);
  await expectNoSelection(page);
});

test('pptx: thumbnail navigation drops the selection', async ({ page }) => {
  test.setTimeout(30_000);
  await openDeck(page);
  await selectOnFirstSlide(page, false);
  const before = await painted(page);
  await thumbnails(page).nth(1).click();
  await expectRepaint(page, before);
  await expect(thumbnails(page).nth(1)).toHaveAttribute('aria-current', 'page');
  await expectNoSelection(page);
  await thumbnails(page).first().click();
  await expect.poll(async () => await painted(page) === before).toBe(true);
  await expectNoSelection(page);
});
