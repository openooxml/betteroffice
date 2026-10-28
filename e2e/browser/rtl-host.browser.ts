import { test, expect, type Page } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');

async function openRtl(page: Page, format: 'docx' | 'xlsx' | 'pptx', file: string) {
  await page.goto('/?format=' + format);
  await page.evaluate(() => document.documentElement.setAttribute('dir', 'rtl'));
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(resolve(root, 'apps/demo/public', file));
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.locator('[data-testid$="-error"]')).toHaveCount(0);
  await expect
    .poll(() =>
      page.evaluate(() => getComputedStyle(document.documentElement).direction)
    )
    .toBe('rtl');
}

async function expectRootLtr(page: Page, selector: string) {
  const rootElement = page.locator(selector).first();
  await expect(rootElement).toHaveAttribute('dir', 'ltr');
  await expect
    .poll(() => rootElement.evaluate((element) => getComputedStyle(element).direction))
    .toBe('ltr');
}

test('xlsx: grid paints on an rtl host page', async ({ page }) => {
  await openRtl(page, 'xlsx', 'showcase.xlsx');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-sheet-tabs')).toBeVisible();
  const scroll = page.getByTestId('xlsx-scroll');
  const canvas = scroll.locator('canvas');
  await expect(canvas).toBeVisible();
  const view = (await scroll.boundingBox())!;
  const surface = (await canvas.boundingBox())!;
  const overlap = Math.min(view.x + view.width, surface.x + surface.width) - Math.max(view.x, surface.x);
  expect(overlap).toBeGreaterThan(view.width / 2);
  await expect
    .poll(async () =>
      canvas.evaluate((element) => {
        const surface = element as HTMLCanvasElement;
        const context = surface.getContext('2d');
        if (!context || surface.width === 0 || surface.height === 0) return 0;
        const width = Math.min(surface.width, 400);
        const height = Math.min(surface.height, 200);
        const data = context.getImageData(0, 0, width, height).data;
        let nonWhite = 0;
        for (let index = 0; index < data.length; index += 4) {
          if (Math.min(data[index], data[index + 1], data[index + 2]) < 240) nonWhite++;
        }
        return nonWhite;
      })
    )
    .toBeGreaterThan(20);
  await expect.poll(async () => page.getByRole('gridcell').count()).toBeGreaterThan(5);
  await expectRootLtr(page, 'div[role="application"]');
});

test('docx: File menu stays inside the viewport on an rtl host page', async ({ page }) => {
  await openRtl(page, 'docx', 'betteroffice-demo.docx');
  await page.getByRole('menuitem', { name: 'File', exact: true }).click();
  const menu = page.getByRole('menu', { name: 'File' });
  await expect(menu).toBeVisible();
  const box = (await menu.boundingBox())!;
  const viewport = page.viewportSize()!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width);
  await expectRootLtr(page, '[data-testid="docx-editor"]');
});

test('pptx: layout dropdown stays inside the viewport on an rtl host page', async ({ page }) => {
  await openRtl(page, 'pptx', 'betteroffice-demo.pptx');
  await expect(page.getByTestId('pptx-slide-canvas')).toBeVisible({ timeout: 60_000 });
  await page.getByTestId('pptx-new-slide-layout').click();
  const menu = page.locator('[role="menu"]').last();
  await expect(menu).toBeVisible();
  const box = (await menu.boundingBox())!;
  const viewport = page.viewportSize()!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width);
  await expectRootLtr(page, '[data-testid="pptx-worker-editor"]');
});
