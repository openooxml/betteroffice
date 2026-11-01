import { expect, test, type Page } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');

async function open(page: Page, arm: string) {
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.docx'));
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
}

for (const arm of ['&peerless=0', '']) {
  test(`ruler markers show their name and value on hover, focus and drag (${arm ? 'hybrid' : 'default'})`, async ({ page }) => {
    await open(page, arm);
    const ruler = page.locator('.docx-horizontal-ruler');
    const tooltip = ruler.getByRole('tooltip');
    const indent = ruler.getByRole('slider', { name: 'Left indent', exact: true });
    await expect(indent).toBeVisible();
    await indent.hover();
    await expect(tooltip).toHaveText(/^Left indent: \d+\.\d\d"$/);
    await expect(tooltip).toBeVisible();
    const [tip, bar] = [(await tooltip.boundingBox())!, (await ruler.boundingBox())!];
    expect(tip.y).toBeGreaterThanOrEqual(bar.y + bar.height);
    await page.mouse.move(0, 0);
    await expect(tooltip).toHaveCount(0);
    await indent.focus();
    await expect(tooltip).toHaveText(/^Left indent: /);
    await expect(indent).toHaveAttribute('aria-valuetext', /^Left indent: /);
    await indent.blur();
    const box = (await indent.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 48, box.y + box.height / 2, { steps: 4 });
    await expect(tooltip).toHaveText('Left indent: 0.50"');
    await page.mouse.up();
    const vertical = page.locator('.docx-vertical-ruler');
    await vertical.getByRole('slider', { name: 'Top margin', exact: true }).hover();
    await expect(vertical.getByRole('tooltip')).toHaveText(/^Top margin: \d+\.\d\d"$/);
  });
}
