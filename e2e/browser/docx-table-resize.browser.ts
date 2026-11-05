import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const ARMS = [{ name: 'worker', query: '&peerless=1' }, { name: 'classic', query: '&peerless=0' }] as const;
type Arm = (typeof ARMS)[number];
const COLUMN = 3120;
const COLUMN_HANDLE = '.layout-table-resize-handle';
const RIGHT_HANDLE = '.layout-table-edge-handle-right';
const ROW_HANDLE = '.layout-table-row-resize-handle';

function documentXml(borders: boolean, rowCount: number): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  const tableBorders = borders ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` : '';
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="${COLUMN}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const header = rowCount > 3 ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
  const rows = Array.from({ length: rowCount }, (_, row) => `<w:tr>${row === 0 ? header : ''}${[0, 1, 2].map((column) => cell(`R${row}C${column}`)).join('')}</w:tr>`).join('');
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="${COLUMN * 3}" w:type="dxa"/>${tableBorders}</w:tblPr>` +
    `<w:tblGrid>${`<w:gridCol w:w="${COLUMN}"/>`.repeat(3)}</w:tblGrid>${rows}</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>';
}

async function open(page: Page, arm: Arm, borders = true, rowCount = 3): Promise<Locator> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', documentXml(borders, rowCount));
  await page.goto('/?format=docx' + arm.query);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-resize.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.canvas-page canvas, .paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

interface SavedTable {
  grid: number[];
  cellWidths: number[][];
  layout: string | null;
  width: number | null;
  texts: string[][];
  bodyText: string;
  rowHeights: (number | null)[];
  rowRules: (string | null)[];
}

async function savedTable(page: Page, info: TestInfo, name: string): Promise<SavedTable> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const table = document.getElementsByTagNameNS('*', 'tbl')[0]!;
    const children = (node: Element, name: string) => Array.from(node.children).filter((child) => child.localName === name);
    const attr = (node: Element | undefined, name: string) => node?.getAttributeNS(node.namespaceURI, name) ?? null;
    const grid = children(children(table, 'tblGrid')[0]!, 'gridCol').map((column) => Number(attr(column, 'w')));
    const rows = children(table, 'tr');
    const cellWidths = rows.map((row) => children(row, 'tc').map((cell) =>
      Number(attr(cell.getElementsByTagNameNS('*', 'tcW')[0], 'w'))));
    const properties = children(table, 'tblPr')[0]!;
    const width = attr(children(properties, 'tblW')[0], 'w');
    const texts = rows.map((row) => children(row, 'tc').map((cell) => cell.textContent ?? ''));
    const heights = rows.map((row) => row.getElementsByTagNameNS('*', 'trHeight')[0]);
    const rowHeights = heights.map((height) => height ? Number(attr(height, 'val')) : null);
    const rowRules = heights.map((height) => attr(height, 'hRule'));
    const bodyText = document.getElementsByTagNameNS('*', 'body')[0]!.textContent ?? '';
    return { grid, cellWidths, texts, bodyText, rowHeights, rowRules, layout: attr(children(properties, 'tblLayout')[0], 'type'), width: width === null ? null : Number(width) };
  }, xml);
}

async function centers(page: Page, selector: string) {
  return page.locator(selector).evaluateAll((elements) => elements.map((element) => {
    const box = element.getBoundingClientRect();
    return { x: box.x + box.width / 2, y: box.y + box.height / 2, top: box.y };
  }).sort((a, b) => a.x - b.x || a.y - b.y));
}

async function rowCenters(page: Page) {
  return (await centers(page, ROW_HANDLE)).sort((a, b) => a.y - b.y);
}

async function pageBox(page: Page) {
  const box = await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().boundingBox();
  expect(box).not.toBeNull();
  return box!;
}

async function drag(page: Page, from: { x: number; y: number }, dx: number, dy: number) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(from.x + dx / 2, from.y + dy / 2, { steps: 4 });
  await page.mouse.move(from.x + dx, from.y + dy, { steps: 4 });
  await page.mouse.up();
}

async function setZoom(page: Page, value: string) {
  const control = page.getByRole('combobox', { name: /^Zoom:/ });
  await control.click();
  await page.getByRole('option', { name: value, exact: true }).click();
  await expect(control).toContainText(value);
}

const near = (actual: number, expected: number, tolerance: number) =>
  expect(Math.abs(actual - expected), `${actual} within ${tolerance} of ${expected}`).toBeLessThanOrEqual(tolerance);

for (const arm of ARMS) {
  test.describe(arm.name, () => {
    test('table resize handles mount on a visible table with resize cursors', async ({ page }) => {
      await open(page, arm);
      await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
      await expect(page.locator(RIGHT_HANDLE)).toHaveCount(1);
      await expect(page.locator(ROW_HANDLE)).toHaveCount(3);
      const cursorAt = async ({ x, y }: { x: number; y: number }) => {
        await page.mouse.move(x, y);
        return page.evaluate(({ x, y }) => getComputedStyle(document.elementFromPoint(x, y)!).cursor, { x, y });
      };

      const columns = await centers(page, COLUMN_HANDLE);
      const [right] = await centers(page, RIGHT_HANDLE);
      near(columns[1]!.x - columns[0]!.x, COLUMN / 15, 2);
      near(right!.x - columns[1]!.x, COLUMN / 15, 2);
      await expect.poll(() => cursorAt(columns[0]!)).toBe('col-resize');
      await expect.poll(() => cursorAt(right!)).toBe('col-resize');
      await expect.poll(async () => cursorAt((await rowCenters(page))[0]!)).toBe('row-resize');
      await expect.poll(() => cursorAt({ x: columns[0]!.x + 40, y: columns[0]!.y })).toBe('text');
    });

    test('a table without borders keeps its handles on the grid, off its text', async ({ page }) => {
      await open(page, arm, false);
      await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
      const columns = await centers(page, COLUMN_HANDLE);
      const [right] = await centers(page, RIGHT_HANDLE);
      near(columns[1]!.x - columns[0]!.x, COLUMN / 15, 2);
      near(right!.x - columns[1]!.x, COLUMN / 15, 2);
      near(columns[0]!.x - (await pageBox(page)).x, 96 + COLUMN / 15, 8);
      await expect(page.locator(ROW_HANDLE)).toHaveCount(0);
    });

    test('dragging a column border saves both column widths and undoes in one step', async ({ page }, info) => {
      const input = await open(page, arm);
      await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
      const before = await centers(page, COLUMN_HANDLE);
      await drag(page, before[0]!, 60, 0);
      await expect.poll(async () => (await centers(page, COLUMN_HANDLE))[0]?.x ?? 0).toBeGreaterThan(before[0]!.x + 55);
      const after = await centers(page, COLUMN_HANDLE);
      near(after[0]!.x - before[0]!.x, 60, 2);
      near(after[1]!.x, before[1]!.x, 1);

      const saved = await savedTable(page, info, 'column');
      expect(saved.grid).toEqual([COLUMN + 900, COLUMN - 900, COLUMN]);
      for (const row of saved.cellWidths) expect(row).toEqual([COLUMN + 900, COLUMN - 900, COLUMN]);
      expect(saved.layout).toBe('fixed');
      expect(saved.width).toBe(COLUMN * 3);

      await input.focus();
      await input.press('ControlOrMeta+z');
      await expect.poll(async () => (await centers(page, COLUMN_HANDLE))[0]?.x ?? 0).toBeLessThan(before[0]!.x + 2);
      expect((await savedTable(page, info, 'column-undo')).grid).toEqual([COLUMN, COLUMN, COLUMN]);
    });

    if (arm.name === 'worker') {
      test('a plain click on a column border reaches the page and places the caret', async ({ page }, info) => {
        await open(page, arm);
        await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
        const [border] = await centers(page, COLUMN_HANDLE);
        const box = await pageBox(page);
        await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: border!.x - box.x, y: border!.y - box.y } });
        await page.keyboard.type('Q');
        await expect.poll(async () => (await savedTable(page, info, 'click')).texts.flat().filter((text) => text.includes('Q')).length).toBe(1);
        const saved = await savedTable(page, info, 'click-grid');
        expect(saved.grid).toEqual([COLUMN, COLUMN, COLUMN]);
        expect(saved.layout).toBeNull();
      });

      test('a plain click on a row border reaches the page and places the caret', async ({ page }, info) => {
        await open(page, arm);
        await expect(page.locator(ROW_HANDLE)).toHaveCount(3);
        const [border] = await rowCenters(page);
        const box = await pageBox(page);
        await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: border!.x - box.x, y: border!.y - box.y } });
        await page.keyboard.type('Q');
        await expect.poll(async () => (await savedTable(page, info, 'row-click')).texts.flat().filter((text) => text.includes('Q')).length).toBe(1);
        expect((await savedTable(page, info, 'row-click-height')).rowHeights).toEqual([null, null, null]);
      });

      test('a press that moves less than the slop stays a click', async ({ page }, info) => {
        await open(page, arm);
        await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
        const [border] = await centers(page, COLUMN_HANDLE);
        await drag(page, border!, 2, 0);
        expect((await savedTable(page, info, 'slop')).grid).toEqual([COLUMN, COLUMN, COLUMN]);
      });
    }

    test('dragging the right edge resizes the last column', async ({ page }, info) => {
      await open(page, arm);
      await expect(page.locator(RIGHT_HANDLE)).toHaveCount(1);
      const [right] = await centers(page, RIGHT_HANDLE);
      await drag(page, right!, -60, 0);
      await expect.poll(async () => (await centers(page, RIGHT_HANDLE))[0]?.x ?? 0).toBeLessThan(right!.x - 55);
      const saved = await savedTable(page, info, 'right');
      expect(saved.grid).toEqual([COLUMN, COLUMN, COLUMN - 900]);
      for (const row of saved.cellWidths) expect(row).toEqual([COLUMN, COLUMN, COLUMN - 900]);
      expect(saved.layout).toBe('fixed');
      if (arm.name === 'worker') expect(saved.width).toBe(COLUMN * 3 - 900);
    });

    test('dragging a row border saves an at-least height, repeats and undoes in one step', async ({ page }, info) => {
      const input = await open(page, arm);
      await expect(page.locator(ROW_HANDLE)).toHaveCount(3);
      const before = await rowCenters(page);
      const rowHeight = before[1]!.y - before[0]!.y;
      await drag(page, before[0]!, 0, 30);
      await expect.poll(async () => (await rowCenters(page))[1]?.y ?? 0).toBeGreaterThan(before[1]!.y + 25);
      const after = await rowCenters(page);
      near(after[0]!.y - before[0]!.y, 30, 2);
      near(after[1]!.y - before[1]!.y, 30, 2);

      const saved = await savedTable(page, info, 'row');
      near(saved.rowHeights[0]!, Math.round((rowHeight + 30) * 15), 30);
      expect(saved.rowRules[0]).toBe('atLeast');
      expect(saved.rowHeights.slice(1)).toEqual([null, null]);
      expect(saved.grid).toEqual([COLUMN, COLUMN, COLUMN]);

      await drag(page, after[0]!, 0, 20);
      await expect.poll(async () => (await rowCenters(page))[1]?.y ?? 0).toBeGreaterThan(after[1]!.y + 15);
      near((await rowCenters(page))[0]!.y - after[0]!.y, 20, 2);

      await input.focus();
      await input.press('ControlOrMeta+z');
      await expect.poll(async () => (await rowCenters(page))[0]?.y ?? 0).toBeLessThan(after[0]!.y + 2);
      near((await savedTable(page, info, 'row-undo-1')).rowHeights[0]!, saved.rowHeights[0]!, 0);
      await input.press('ControlOrMeta+z');
      await expect.poll(async () => (await rowCenters(page))[0]?.y ?? 0).toBeLessThan(before[0]!.y + 2);
      expect((await savedTable(page, info, 'row-undo-2')).rowHeights).toEqual([null, null, null]);
    });

    test('typing straight after a row drag undoes on its own', async ({ page }, info) => {
      const input = await open(page, arm);
      await expect(page.locator(ROW_HANDLE)).toHaveCount(3);
      const before = await rowCenters(page);
      await drag(page, before[0]!, 0, 30);
      await input.focus();
      await page.keyboard.type('Z');
      await expect.poll(async () => (await savedTable(page, info, 'row-type')).bodyText).toContain('Z');
      await input.press('ControlOrMeta+z');
      await expect.poll(async () => (await savedTable(page, info, 'row-type-undo')).bodyText).not.toContain('Z');
      expect((await savedTable(page, info, 'row-type-kept')).rowHeights[0]).not.toBeNull();
    });

    test('a repeated header row saves from its own height, not its copies', async ({ page }, info) => {
      await open(page, arm, true, 120);
      await expect.poll(async () => (await rowCenters(page)).length).toBeGreaterThan(3);
      const [first, second] = await rowCenters(page);
      const rowHeight = second!.y - first!.y;
      await drag(page, first!, 0, 12);
      await expect.poll(async () => (await savedTable(page, info, 'header')).rowHeights[0]).not.toBeNull();
      near((await savedTable(page, info, 'header-height')).rowHeights[0]!, Math.round((rowHeight + 12) * 15), 30);
    });

    test('handles follow the borders and drags keep twips under zoom', async ({ page }, info) => {
      await open(page, arm);
      await expect(page.locator(COLUMN_HANDLE)).toHaveCount(2);
      const ratio = async () => {
        const box = await pageBox(page);
        const [first] = await centers(page, COLUMN_HANDLE);
        return { x: (first!.x - box.x) / box.width, y: (first!.top - box.y) / box.height };
      };
      const offset = async () => (await centers(page, COLUMN_HANDLE))[0]!.x - (await pageBox(page)).x;
      await setZoom(page, '100%');
      const width = (await pageBox(page)).width;
      const start = await offset();
      let expected = [COLUMN, COLUMN, COLUMN];
      for (const value of ['50%', '200%']) {
        await setZoom(page, '100%');
        await expect.poll(async () => (await pageBox(page)).width).toBeCloseTo(width, 0);
        await expect.poll(offset).toBeCloseTo(start + (expected[0]! - COLUMN) / 15, 0);
        const base = await ratio();
        await setZoom(page, value);
        const scale = parseFloat(value) / 100;
        await expect.poll(async () => (await pageBox(page)).width).toBeCloseTo(width * scale, 0);
        await expect.poll(async () => (await ratio()).x).toBeCloseTo(base.x, 2);
        near((await ratio()).y, base.y, 0.005);
        const [first] = await centers(page, COLUMN_HANDLE);
        await drag(page, { x: first!.x, y: first!.top + 4 }, 20 * scale, 0);
        expected = [expected[0]! + 300, expected[1]! - 300, COLUMN];
        await expect.poll(async () => (await savedTable(page, info, `zoom-${parseFloat(value)}`)).grid).toEqual(expected);
        const [row] = await rowCenters(page);
        const [, next] = await rowCenters(page);
        await drag(page, row!, 0, 10 * scale);
        await expect.poll(async () => (await rowCenters(page))[1]?.y ?? 0).toBeGreaterThan(next!.y + 8 * scale);
      }
    });
  });
}
