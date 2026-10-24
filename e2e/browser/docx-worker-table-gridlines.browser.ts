import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

type PrintScope = { __printPopups?: Window[] };

function documentXml(outer: string): string {
  const edge = (side: string, value: string) => `<w:${side} w:val="${value}" w:sz="4" w:space="0" w:color="000000"/>`;
  const borders = ['top', 'left', 'bottom', 'right'].map((side) => edge(side, outer)).join('') +
    ['insideH', 'insideV'].map((side) => edge(side, 'nil')).join('');
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const rows = [0, 1].map((row) => `<w:tr>${[0, 1].map((column) => cell(row === 0 && column === 0 ? 'Label' : '')).join('')}</w:tr>`).join('');
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${'<w:gridCol w:w="2340"/>'.repeat(2)}</w:tblGrid>${rows}</w:tbl><w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>';
}

async function open(page: Page, outer = 'nil') {
  await page.addInitScript(() => {
    const scope = window as unknown as PrintScope;
    scope.__printPopups = [];
    const open = window.open.bind(window);
    window.open = (...args: Parameters<typeof window.open>) => {
      const popup = open(...args);
      if (popup) {
        popup.print = () => {};
        popup.close = () => {};
        scope.__printPopups!.push(popup);
      }
      return popup;
    };
  });
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', documentXml(outer));
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-gridlines.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await input.press('ArrowDown');
}

const gridlines = (page: Page) => page.locator('svg[data-table-gridlines] line');

async function toggleFromBorders(page: Page, pressed: boolean) {
  await page.getByTestId('toolbar-table-borders').click();
  const toggle = page.getByTestId('toolbar-table-gridlines');
  await expect(toggle).toContainText('View gridlines');
  await expect(toggle).toHaveAttribute('aria-pressed', String(pressed));
  await toggle.click();
  await expect(toggle).toBeHidden();
}

async function savedParts(page: Page, info: TestInfo, step: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`table-gridlines-${step}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const entries = Object.values(zip.files).filter((entry) => !entry.dir);
  return Object.fromEntries(await Promise.all(entries.map(async (entry) => [entry.name, await entry.async('base64')] as const)));
}

async function printedPages(page: Page) {
  const index = await page.evaluate(() => (window as unknown as PrintScope).__printPopups!.length);
  await page.getByRole('menubar').getByRole('menuitem', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Print', exact: true }).click();
  await page.waitForFunction((at) => {
    const images = (window as unknown as PrintScope).__printPopups![at]?.document.images;
    return !!images?.length && [...images].every((image) => image.complete && image.naturalWidth > 0);
  }, index, { timeout: 60000 });
  return page.evaluate((at) => {
    const popup = (window as unknown as PrintScope).__printPopups![at]!;
    return [...popup.document.images].map((image) => {
      const canvas = popup.document.createElement('canvas');
      canvas.width = image.naturalWidth;
      canvas.height = image.naturalHeight;
      canvas.getContext('2d')!.drawImage(image, 0, 0);
      return canvas.toDataURL();
    });
  }, index);
}

test('worker table gridlines show borderless cell edges on screen only', async ({ page }, info) => {
  await open(page);
  await expect(page.getByTestId('toolbar-table-borders')).toBeEnabled();
  await expect(page.locator('svg[data-table-gridlines]')).toHaveCount(0);
  const savedOff = await savedParts(page, info, 'off');
  const printedOff = await printedPages(page);
  expect(printedOff.length).toBe(1);

  await toggleFromBorders(page, false);
  await expect(gridlines(page)).toHaveCount(12);
  await expect(page.locator('svg[data-table-gridlines]')).toHaveAttribute('aria-hidden', 'true');
  expect(await savedParts(page, info, 'on')).toEqual(savedOff);
  expect(await printedPages(page)).toEqual(printedOff);

  await toggleFromBorders(page, true);
  await expect(page.locator('svg[data-table-gridlines]')).toHaveCount(0);
});

test('worker table gridlines mark only the edges a partial border leaves open', async ({ page }) => {
  await open(page, 'single');
  await toggleFromBorders(page, false);
  await expect(gridlines(page)).toHaveCount(4);
});

test('worker table gridlines toggle from the overflow Table menu after Borders', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);
  await expect(page.getByTestId('toolbar-table-borders')).toBeHidden();
  const tableMenu = async () => {
    await page.getByTestId('toolbar-more').click();
    await page.getByRole('menu', { name: 'More actions', exact: true }).locator('[data-label="Table"]').focus();
    await page.keyboard.press('ArrowRight');
    const menu = page.getByRole('menu', { name: 'Table', exact: true });
    await expect(menu).toBeVisible();
    return menu;
  };
  let menu = await tableMenu();
  const labels = await menu.locator('[data-label]').evaluateAll((items) => items.map((item) => item.getAttribute('data-label')));
  expect(labels.indexOf('View gridlines')).toBe(labels.indexOf('Borders') + 1);
  const item = menu.getByRole('menuitemcheckbox', { name: 'View gridlines', exact: true });
  await expect(item).toHaveAttribute('aria-checked', 'false');
  await item.click();
  await expect(gridlines(page)).toHaveCount(12);

  menu = await tableMenu();
  await expect(menu.getByRole('menuitemcheckbox', { name: 'View gridlines', exact: true })).toHaveAttribute('aria-checked', 'true');
  await menu.getByRole('menuitemcheckbox', { name: 'View gridlines', exact: true }).click();
  await expect(page.locator('svg[data-table-gridlines]')).toHaveCount(0);
});
