import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

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
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-selection.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function highlightCount(page: Page): Promise<number> {
  return page.evaluate(() => [...document.querySelectorAll('div[aria-hidden="true"]')]
    .filter((element) => (element as HTMLElement).style.background.includes('59, 130, 246')).length);
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
  for (let attempt = 0; attempt < 4; attempt += 1) {
    await input.press('ArrowDown');
    const surface = page.locator('.paged-editor__page').first();
    const box = (await surface.boundingBox())!;
    await surface.click({ button: 'right', position: { x: 100, y: 100 }, force: true });
    const item = page.getByRole('menuitem', { name: 'Select entire table', exact: true });
    if (await item.count() > 0) {
      await item.click();
      await expectHighlights(page, 12);
      return;
    }
    await page.keyboard.press('Escape');
    expect(box.width).toBeGreaterThan(0);
  }
  throw new Error('Caret did not enter the table');
}

test('worker select entire table highlights every cell and a click clears it', async ({ page }) => {
  const input = await open(page);
  await caretIntoTable(page, input);
  expect(await highlightCount(page)).toBe(12);
  const grid = await highlightGrid(page);
  expect(grid.length).toBe(3);
  expect(grid[0]!.length).toBe(4);
  await page.mouse.click(grid[0]![0]!.x, grid[0]![0]!.y);
  await expectHighlights(page, 0);
});

test('worker drag shift-click and shift-arrow extend the cell range', async ({ page }) => {
  const input = await open(page);
  await caretIntoTable(page, input);
  const grid = await highlightGrid(page);
  await page.mouse.click(grid[0]![0]!.x, grid[0]![0]!.y);
  await expectHighlights(page, 0);
  await page.mouse.move(grid[0]![0]!.x, grid[0]![0]!.y);
  await page.mouse.down();
  await page.mouse.move(grid[1]![1]!.x, grid[1]![1]!.y, { steps: 8 });
  await page.mouse.up();
  await expectHighlights(page, 4);
  await page.keyboard.down('Shift');
  await page.mouse.click(grid[2]![2]!.x, grid[2]![2]!.y);
  await page.keyboard.up('Shift');
  await expectHighlights(page, 9);
  await input.focus();
  await input.press('Shift+ArrowRight');
  await expectHighlights(page, 12);
  await page.mouse.click(grid[0]![0]!.x, grid[0]![0]!.y);
  await expectHighlights(page, 0);
});
