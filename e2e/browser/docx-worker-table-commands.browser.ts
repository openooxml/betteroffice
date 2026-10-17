import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

function documentXml(): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const rows = [0, 1, 2].map((row) => `<w:tr>${[0, 1, 2, 3].map((column) => cell(`R${row}C${column}`)).join('')}</w:tr>`).join('');
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${'<w:gridCol w:w="2340"/>'.repeat(4)}</w:tblGrid>${rows}</w:tbl><w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>';
}

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', documentXml());
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-commands.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function expectHighlights(page: Page, count: number) {
  await page.waitForFunction((expected) => [...document.querySelectorAll('div[aria-hidden="true"]')]
    .filter((element) => (element as HTMLElement).style.background.includes('59, 130, 246')).length === expected,
    count, { timeout: 30000 });
}

async function highlightGrid(page: Page): Promise<{ x: number; y: number }[][]> {
  const centers = await page.evaluate(() => [...document.querySelectorAll('div[aria-hidden="true"]')]
    .filter((element) => (element as HTMLElement).style.background.includes('59, 130, 246'))
    .map((element) => {
      const box = element.getBoundingClientRect();
      return { x: box.x + box.width / 2, y: box.y + box.height / 2, height: box.height };
    }));
  const rows: typeof centers[] = [];
  for (const center of centers.sort((a, b) => a.y - b.y)) {
    const row = rows.find((candidate) => Math.abs(candidate[0]!.y - center.y) < center.height / 2);
    if (row) row.push(center);
    else rows.push([center]);
  }
  return rows.map((row) => row.sort((a, b) => a.x - b.x).map(({ x, y }) => ({ x, y })));
}

async function caretIntoTable(page: Page, input: ReturnType<Page['locator']>) {
  await input.press('ControlOrMeta+Home');
  await input.press('ArrowDown');
  await expect(async () => {
    if (await page.locator('[role="menu"]').count() > 0) await page.keyboard.press('Escape');
    await page.locator('.paged-editor__page').first()
      .click({ button: 'right', position: { x: 100, y: 100 }, force: true });
    await page.getByRole('menuitem', { name: 'Select entire table', exact: true }).click({ timeout: 2000 });
    await page.waitForFunction(() => [...document.querySelectorAll('div[aria-hidden="true"]')]
      .filter((element) => (element as HTMLElement).style.background.includes('59, 130, 246')).length >= 1,
      undefined, { timeout: 3000 });
  }).toPass({ timeout: 30000 });
}

async function tableMenuAction(page: Page, label: string) {
  if (await page.locator('[role="menu"]').count() > 0) await page.keyboard.press('Escape');
  await page.locator('.paged-editor__page').first()
    .click({ button: 'right', position: { x: 100, y: 100 }, force: true });
  const item = page.getByRole('menuitem', { name: label, exact: true });
  await expect(item).not.toHaveAttribute('aria-disabled', 'true');
  await item.click();
}

async function setCellFillRed(page: Page) {
  const fillButton = page.getByRole('button', { name: 'Cell Fill Color', exact: true }).last();
  if (await fillButton.isVisible()) {
    await fillButton.click();
    const picker = page.getByRole('dialog', { name: 'Text Highlight Color picker', exact: true });
    await expect(picker).toBeVisible();
    await picker.getByLabel('Custom hex color', { exact: true }).fill('FF0000');
    await picker.getByRole('button', { name: 'Apply', exact: true }).click();
    return;
  }
  await page.getByRole('button', { name: 'More actions', exact: true }).click();
  const overflow = page.getByRole('menu', { name: 'More actions', exact: true });
  await overflow.getByRole('menuitem', { name: 'Table', exact: true }).hover();
  const table = page.getByRole('menu', { name: 'Table', exact: true });
  await table.getByRole('menuitem', { name: 'Cell Fill Color', exact: true }).hover();
  const fill = page.getByRole('menu', { name: 'Cell Fill Color', exact: true });
  await fill.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
}

async function savedGrid(page: Page, info: TestInfo, step: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`table-commands-${step}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const doc = new DOMParser().parseFromString(source, 'application/xml');
    const table = doc.getElementsByTagNameNS('*', 'tbl')[0];
    if (!table) return { hasTable: false, rows: [] as number[], gridCols: 0, spans: [] as string[], fills: [] as string[] };
    const rows = Array.from(table.getElementsByTagNameNS('*', 'tr'))
      .map((row) => row.getElementsByTagNameNS('*', 'tc').length);
    const grid = table.getElementsByTagNameNS('*', 'tblGrid')[0];
    const spans = Array.from(table.getElementsByTagNameNS('*', 'gridSpan'))
      .map((node) => node.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? '');
    const fills = Array.from(table.getElementsByTagNameNS('*', 'shd'))
      .map((node) => node.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'fill') ?? '');
    return { hasTable: true, rows, gridCols: grid ? grid.getElementsByTagNameNS('*', 'gridCol').length : 0, spans, fills };
  }, xml);
}

test('worker table commands reshape the saved grid', async ({ page }, info) => {
  const input = await open(page);
  await caretIntoTable(page, input);
  await expectHighlights(page, 12);
  const initial = await highlightGrid(page);
  await page.mouse.click(initial[0]![0]!.x, initial[0]![0]!.y);
  await expectHighlights(page, 0);

  await tableMenuAction(page, 'Insert row below');
  expect(await savedGrid(page, info, 'row-below')).toMatchObject({ hasTable: true, rows: [4, 4, 4, 4], gridCols: 4 });

  await tableMenuAction(page, 'Insert row above');
  expect(await savedGrid(page, info, 'row-above')).toMatchObject({ hasTable: true, rows: [4, 4, 4, 4, 4], gridCols: 4 });

  await tableMenuAction(page, 'Insert column right');
  expect(await savedGrid(page, info, 'column-right')).toMatchObject({ hasTable: true, rows: [5, 5, 5, 5, 5], gridCols: 5 });

  await tableMenuAction(page, 'Insert column left');
  expect(await savedGrid(page, info, 'column-left')).toMatchObject({ hasTable: true, rows: [6, 6, 6, 6, 6], gridCols: 6 });

  await tableMenuAction(page, 'Delete column');
  expect(await savedGrid(page, info, 'delete-column')).toMatchObject({ hasTable: true, rows: [5, 5, 5, 5, 5], gridCols: 5 });

  await tableMenuAction(page, 'Delete row');
  expect(await savedGrid(page, info, 'delete-row')).toMatchObject({ hasTable: true, rows: [5, 5, 5, 5], gridCols: 5 });

  await input.focus();
  await input.press('ControlOrMeta+z');
  expect(await savedGrid(page, info, 'undo')).toMatchObject({ hasTable: true, rows: [5, 5, 5, 5, 5], gridCols: 5 });

  await caretIntoTable(page, input);
  const grid = await highlightGrid(page);
  expect(grid.length).toBe(5);
  await page.mouse.click(grid[0]![0]!.x, grid[0]![0]!.y);
  await expectHighlights(page, 0);
  await input.focus();
  await input.press('Shift+ArrowRight');
  await expectHighlights(page, 2);
  await tableMenuAction(page, 'Merge cells');
  const merged = await savedGrid(page, info, 'merge');
  expect(merged.hasTable).toBe(true);
  expect(merged.spans).toContain('2');
  expect(merged.rows[0]).toBe(4);

  await tableMenuAction(page, 'Split cell');
  const dialog = page.getByRole('dialog', { name: 'Split Cell', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(await savedGrid(page, info, 'split')).toMatchObject({ hasTable: true, rows: [5, 5, 5, 5, 5], gridCols: 5 });

  await setCellFillRed(page);
  const filled = await savedGrid(page, info, 'fill');
  expect(filled.fills.map((fill) => fill.toUpperCase())).toContain('FF0000');

  await tableMenuAction(page, 'Delete table');
  expect(await savedGrid(page, info, 'delete-table')).toMatchObject({ hasTable: false });
});
