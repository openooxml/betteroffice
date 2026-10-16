import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const text = 'Toolbar sample';

async function savedProperties(page: Page, info: TestInfo, arm: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${arm}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ xml, text }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const paragraphs = Array.from(document.getElementsByTagNameNS('*', 'p'));
    const paragraph = paragraphs.find(node => Array.from(node.getElementsByTagNameNS('*', 't'))
      .map(node => node.textContent ?? '').join('').includes(text));
    const value = (name: string) => paragraph?.getElementsByTagNameNS('*', name)[0]?.getAttributeNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? null;
    const table = document.getElementsByTagNameNS('*', 'tbl')[0];
    return {
      strike: !!paragraph?.getElementsByTagNameNS('*', 'strike').length,
      alignment: value('jc'),
      list: !!paragraph?.getElementsByTagNameNS('*', 'numPr').length,
      style: value('pStyle'),
      table: table ? Array.from(table.getElementsByTagNameNS('*', 'tr')).map(row => row.getElementsByTagNameNS('*', 'tc').length) : null,
    };
  }, { xml, text });
}

test('docx toolbar saves equal formatting and table properties in both arms', async ({ page }, info) => {
  await page.setViewportSize({ width: 2560, height: 1200 });
  const results: Awaited<ReturnType<typeof savedProperties>>[] = [];
  for (const arm of ['&peerless=0', '']) {
    const label = arm ? 'hybrid' : 'default';
    await test.step(label, async () => {
      await page.goto('/?format=docx' + arm);
      await page.locator('input[type=file]').last().setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.docx'));
      await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
      const rail = page.getByTestId('formatting-bar');
      await expect(rail).toBeVisible();
      await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
      const input = page.locator('[data-testid=yrs-input], .paged-editor__peerless-input');
      await expect(input).toBeEditable();
      await input.press('ControlOrMeta+Home');
      await input.pressSequentially(text);
      await input.press('Shift+Home');
      const strike = rail.getByRole('button', { name: 'Strikethrough', exact: true });
      await expect(strike).not.toHaveAttribute('aria-disabled', 'true');
      await strike.click();
      await expect(strike).toHaveAttribute('aria-pressed', 'true');
      await rail.getByTestId('toolbar-alignment').click();
      await page.getByTestId('alignment-center').click();
      const bullet = rail.getByRole('button', { name: 'Bullet List', exact: true });
      await bullet.click();
      await expect(bullet).toHaveAttribute('aria-pressed', 'true');
      await rail.getByRole('combobox', { name: 'Select paragraph style', exact: true }).click();
      await page.getByRole('option', { name: 'Heading 1', exact: true }).click();
      await input.press('ArrowRight');
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Insert', exact: true }).click();
      await page.getByRole('menu', { name: 'Insert' }).getByRole('menuitem', { name: 'Table', exact: true }).hover();
      const grid = page.getByRole('grid', { name: 'Table size selector' });
      await expect(grid).toBeVisible();
      await grid.getByRole('gridcell').nth(7).click();
      const properties = await savedProperties(page, info, label);
      expect(properties).toEqual({ strike: true, alignment: 'center', list: true, style: 'Heading1', table: [2, 2] });
      results.push(properties);
    });
  }
  expect(results[1]).toEqual(results[0]);
});
