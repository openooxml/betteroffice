import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

type Rect = { x: number; y: number; width: number; height: number };
type Grid = { x: number; y: number; width: number; height: number };
type XmlSnapshot = { name: string; attributes: Record<string, string>; text: string; children: XmlSnapshot[] };
const namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function column(index: number): string {
  let name = '';
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + (value - 1) % 26) + name;
  }
  return name;
}

async function workbook(frozen: boolean) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <bookViews><workbookView activeTab="0"/></bookViews>
    <sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets>
  </workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${namespace}">
    <fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
    <borders count="1"><border/></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  </styleSheet>`);
  for (const sheet of [1, 2]) {
    const rows: string[] = [];
    for (let row = 1; row <= 180; row++) {
      const cells: string[] = [];
      for (let col = 1; col <= 40; col++) {
        if ((row <= 6 && col <= 4) || (row === 3 && col >= 8 && col <= 10) ||
          (col === 3 && row >= 10 && row <= 12) || (row === 180 && col === 40)) {
          cells.push(`<c r="${column(col)}${row}"><v>${(sheet - 1) * 100000 + row * 100 + col}</v></c>`);
        }
      }
      if (cells.length) rows.push(`<row r="${row}" ht="18" customHeight="1">${cells.join('')}</row>`);
    }
    zip.file(`xl/worksheets/sheet${sheet}.xml`, `<worksheet xmlns="${namespace}">
      <dimension ref="A1:AN180"/>
      <sheetViews><sheetView workbookViewId="0" zoomScale="100" zoomScaleNormal="100">${frozen ? `
        <pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>
        <selection pane="topRight" activeCell="C1" sqref="C1"/>
        <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
        <selection pane="bottomRight" activeCell="C3" sqref="C3:D4"/>
      ` : '<selection activeCell="A1" sqref="A1"/>'}</sheetView></sheetViews>
      <sheetFormatPr defaultRowHeight="18" customHeight="1"/>
      <cols><col min="1" max="40" width="12" customWidth="1"/></cols>
      <sheetData>${rows.join('')}</sheetData>
    </worksheet>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, frozen = false) {
  const buffer = await workbook(frozen);
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'selection.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
  await expect(page.getByTestId('xlsx-scroll').locator('canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-scroll').locator('canvas').first()).toHaveAttribute('data-worker-zoom', '1');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('101');
  await expect(page.getByTestId('xlsx-selection')).toBeVisible();
  await expect(page.locator('.notice, [data-testid="xlsx-error"]')).toHaveCount(0);
  const canvas = (await page.getByTestId('xlsx-scroll').locator('canvas').first().boundingBox())!;
  const selection = (await page.getByTestId('xlsx-selection').boundingBox())!;
  const grid = {
    x: selection.x - canvas.x,
    y: selection.y - canvas.y,
    width: selection.width,
    height: selection.height,
  };
  expect(grid.width).toBeGreaterThan(10);
  expect(grid.height).toBeCloseTo(24, 0);
  return { buffer, grid };
}

async function point(page: Page, grid: Grid, row: number, col: number, scale = 1) {
  const canvas = (await page.getByTestId('xlsx-scroll').locator('canvas').first().boundingBox())!;
  return {
    x: canvas.x + (grid.x + (col - 0.5) * grid.width) * scale,
    y: canvas.y + (grid.y + (row - 0.5) * grid.height) * scale,
  };
}

async function clickCell(page: Page, grid: Grid, row: number, col: number, scale = 1) {
  const at = await point(page, grid, row, col, scale);
  await page.mouse.click(at.x, at.y);
}

async function expectRect(page: Page, expected: Rect) {
  await expect.poll(async () => {
    const canvas = await page.getByTestId('xlsx-scroll').locator('canvas').first().boundingBox();
    const selection = await page.getByTestId('xlsx-selection').boundingBox();
    if (!canvas || !selection) return Infinity;
    return Math.max(
      Math.abs(selection.x - canvas.x - expected.x),
      Math.abs(selection.y - canvas.y - expected.y),
      Math.abs(selection.width - expected.width),
      Math.abs(selection.height - expected.height)
    );
  }).toBeLessThan(1);
}

async function expectRange(page: Page, grid: Grid, top: number, left: number, bottom = top, right = left, scale = 1) {
  await expectRect(page, {
    x: (grid.x + (left - 1) * grid.width) * scale,
    y: (grid.y + (top - 1) * grid.height) * scale,
    width: (right - left + 1) * grid.width * scale,
    height: (bottom - top + 1) * grid.height * scale,
  });
}

async function expectActiveCell(page: Page, cell: string, value?: string) {
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(cell);
  if (value !== undefined) await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(value);
}

async function workbookSnapshot(page: Page, buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const paths = ['xl/workbook.xml', 'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'];
  const xml = await Promise.all(paths.map((path) => {
    const part = zip.file(path);
    if (!part) throw new Error(`Missing workbook part: ${path}`);
    return part.async('string');
  }));
  return page.evaluate((parts) => {
    const documents = parts.map((part) => {
      const document = new DOMParser().parseFromString(part, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error('Invalid workbook XML');
      return document;
    });
    const attributes = (node: Element) => Object.fromEntries(
      Array.from(node.attributes).map((attribute) => [attribute.name, attribute.value]).sort(([a], [b]) => a.localeCompare(b))
    );
    const element = (node: Element): XmlSnapshot => ({
      name: node.localName,
      attributes: attributes(node),
      text: Array.from(node.childNodes).filter((child) => child.nodeType === Node.TEXT_NODE)
        .map((child) => child.textContent).join('').trim(),
      children: Array.from(node.children).map(element),
    });
    return {
      workbook: element(documents[0].documentElement),
      styles: element(documents[1].documentElement),
      sheets: documents.slice(2).map((document) => ({
        cells: Array.from(document.getElementsByTagNameNS('*', 'c')).map((cell) => ({
          ref: cell.getAttribute('r'),
          type: cell.getAttribute('t') ?? 'n',
          value: cell.getElementsByTagNameNS('*', 'v')[0]?.textContent ??
            Array.from(cell.getElementsByTagNameNS('*', 't')).map((node) => node.textContent).join(''),
          formula: cell.getElementsByTagNameNS('*', 'f')[0]?.textContent ?? null,
          style: cell.getAttribute('s') ?? '0',
        })),
        rows: Array.from(document.getElementsByTagNameNS('*', 'row')).map(attributes),
        cols: Array.from(document.getElementsByTagNameNS('*', 'col')).map(attributes),
        format: Array.from(document.getElementsByTagNameNS('*', 'sheetFormatPr')).map(attributes),
        dimensions: Array.from(document.getElementsByTagNameNS('*', 'dimension')).map(attributes),
        views: Array.from(document.getElementsByTagNameNS('*', 'sheetViews')).map(element),
      })),
    };
  }, xml);
}

async function saveAndCheck(page: Page, info: TestInfo, original: Buffer, label: string, edited = false, reopen = false) {
  const expected = await workbookSnapshot(page, original);
  if (edited) expected.sheets[0].cells.find((cell) => cell.ref === 'C3')!.value = '4242';
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.xlsx');
  await (await downloading).saveAs(file);
  expect(await workbookSnapshot(page, await readFile(file))).toEqual(expected);
  if (reopen) {
    await page.goto('/?format=xlsx');
    await page.locator('input[type=file]').last().setInputFiles(file);
    await expect(page.getByTestId('xlsx-selection')).toBeVisible();
    await expect(page.getByTestId('xlsx-scroll').locator('canvas').first()).toHaveAttribute('data-worker-zoom', '1');
    await expect(page.getByRole('tab', { name: 'First', exact: true })).toBeVisible();
    await expect(page.getByRole('tab', { name: 'Second', exact: true })).toBeVisible();
    if (edited) {
      const scroll = page.getByTestId('xlsx-scroll');
      await scroll.press('ControlOrMeta+Home');
      await scroll.press('ArrowRight');
      await scroll.press('ArrowRight');
      await scroll.press('ArrowDown');
      await scroll.press('ArrowDown');
      await expectActiveCell(page, 'C3', '4242');
    }
    await saveAndCheck(page, info, original, label + '-reopened', edited);
  }
}

test('xlsx selection: Ctrl or Meta horizontal arrows stop at the next contiguous data edge', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 3, 3);
  await expectActiveCell(page, 'C3', '303');
  await page.keyboard.press('ControlOrMeta+ArrowRight');
  await expectActiveCell(page, 'D3', '304');
  await expectRange(page, grid, 3, 4);
  await saveAndCheck(page, info, buffer, 'horizontal-edge');
});

test('xlsx selection: Ctrl or Meta vertical arrows stop at the next contiguous data edge', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 3, 3);
  await expectActiveCell(page, 'C3', '303');
  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await expectActiveCell(page, 'C6', '603');
  await expectRange(page, grid, 6, 3);
  await saveAndCheck(page, info, buffer, 'vertical-edge');
});

test('xlsx selection: Ctrl or Meta A first selects the current data region', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 2, 2);
  await expectActiveCell(page, 'B2', '202');
  await page.keyboard.press('ControlOrMeta+a');
  await expectRange(page, grid, 1, 1, 6, 4);
  await saveAndCheck(page, info, buffer, 'select-region');
});

test('xlsx selection: Home returns to column A and Ctrl or Meta Home returns to A1', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 5, 4);
  await page.keyboard.press('Home');
  await expectActiveCell(page, 'A5', '501');
  await expectRange(page, grid, 5, 1);
  await page.keyboard.press('ControlOrMeta+Home');
  await expectActiveCell(page, 'A1', '101');
  await expectRange(page, grid, 1, 1);
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('ArrowUp');
  await expectActiveCell(page, 'A1', '101');
  await expectRange(page, grid, 1, 1);
  await saveAndCheck(page, info, buffer, 'home');
});

test('xlsx selection: PageDown and PageUp move one visible page and reveal the active cell', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 2, 3);
  const height = await page.getByTestId('xlsx-scroll').evaluate((element) => element.clientHeight);
  const row = 2 + Math.floor(height / grid.height);
  await page.keyboard.press('PageDown');
  await expectActiveCell(page, `C${row}`, '');
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await expect(page.getByTestId('xlsx-selection')).toBeInViewport({ ratio: 1 });
  await page.keyboard.press('PageUp');
  await expectActiveCell(page, 'C2', '203');
  await expect(page.getByTestId('xlsx-selection')).toBeInViewport({ ratio: 1 });
  await page.keyboard.press('ControlOrMeta+Home');
  await expectActiveCell(page, 'A1', '101');
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBe(0);
  await expectRange(page, grid, 1, 1);
  await saveAndCheck(page, info, buffer, 'page-navigation');
});

test('xlsx selection: Tab and Shift-Tab navigate cells while retaining grid focus', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 2, 2);
  for (const [key, cell, col] of [['Tab', 'C2', 3], ['Tab', 'D2', 4], ['Shift+Tab', 'C2', 3]] as const) {
    await page.keyboard.press(key);
    await expectActiveCell(page, cell, String(200 + col));
    await expectRange(page, grid, 2, col);
    await expect(page.getByTestId('xlsx-scroll')).toBeFocused();
  }
  await saveAndCheck(page, info, buffer, 'tab-navigation', false, true);
});

test('xlsx selection: wheel scrolling moves the viewport while keeping the selection aligned', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 12, 3);
  await expectActiveCell(page, 'C12', '1203');
  const at = await point(page, grid, 12, 3);
  await page.mouse.move(at.x, at.y);
  await page.mouse.wheel(0, grid.height * 4);
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await expect.poll(async () => {
    const scroll = await page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop);
    const canvas = (await page.getByTestId('xlsx-scroll').locator('canvas').first().boundingBox())!;
    const selection = await page.getByTestId('xlsx-selection').boundingBox();
    return selection ? Math.abs(selection.y - canvas.y - (grid.y + 11 * grid.height - scroll)) : Infinity;
  }).toBeLessThan(1);
  await expectActiveCell(page, 'C12', '1203');
  await expect(page.getByTestId('xlsx-selection')).toBeInViewport({ ratio: 1 });
  await page.mouse.wheel(0, -grid.height * 4);
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBe(0);
  await expectRange(page, grid, 12, 3);
  await saveAndCheck(page, info, buffer, 'wheel');
});

test('xlsx selection: arrow navigation beyond the viewport scrolls the active cell into view', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 2, 2);
  const size = await page.getByTestId('xlsx-scroll').evaluate((element) => ({
    width: element.clientWidth, height: element.clientHeight,
  }));
  const row = Math.floor(size.height / grid.height) + 3;
  const col = Math.floor(size.width / grid.width) + 3;
  for (let index = 2; index < row; index++) await page.keyboard.press('ArrowDown');
  await expectActiveCell(page, `B${row}`, '');
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await expect(page.getByTestId('xlsx-selection')).toBeInViewport({ ratio: 1 });
  for (let index = 2; index < col; index++) await page.keyboard.press('ArrowRight');
  await expectActiveCell(page, `${column(col)}${row}`, '');
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  await expect(page.getByTestId('xlsx-selection')).toBeInViewport({ ratio: 1 });
  const scrolled = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  await expectRect(page, {
    x: grid.x + (col - 1) * grid.width - scrolled.x,
    y: grid.y + (row - 1) * grid.height - scrolled.y,
    width: grid.width,
    height: grid.height,
  });
  await page.keyboard.press('ControlOrMeta+Home');
  await expectActiveCell(page, 'A1', '101');
  await expectRange(page, grid, 1, 1);
  await saveAndCheck(page, info, buffer, 'offscreen-arrows');
});

test('xlsx selection: zoom changes painted scale and preserves cell hit testing', async ({ page }, info) => {
  const { buffer, grid } = await open(page);
  await clickCell(page, grid, 3, 3);
  for (const [scale, row, col] of [[0.5, 2, 2], [1.5, 4, 4], [2, 3, 3], [1, 5, 2]] as const) {
    const zoom = page.getByTestId('xlsx-zoom');
    await zoom.fill(`${scale * 100}%`);
    await zoom.press('Enter');
    await expect(zoom).toHaveValue(`${scale * 100}%`);
    await expect(page.getByTestId('xlsx-scroll').locator('canvas').first()).toHaveAttribute('data-worker-zoom', String(scale));
    await clickCell(page, grid, row, col, scale);
    await expectActiveCell(page, `${column(col)}${row}`, String(row * 100 + col));
    await expectRange(page, grid, row, col, row, col, scale);
    await saveAndCheck(page, info, buffer, `zoom-${scale}`);
  }
  await saveAndCheck(page, info, buffer, 'zoom-restored', false, true);
});

test('xlsx selection: frozen rows and columns retain their positions and hit testing while scrolling', async ({ page }, info) => {
  const { buffer, grid } = await open(page, true);
  await clickCell(page, grid, 2, 2);
  await expectActiveCell(page, 'B2', '202');
  const at = await point(page, grid, 8, 8);
  await page.mouse.move(at.x, at.y);
  await page.mouse.wheel(grid.width * 4, grid.height * 6);
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await expect(page.getByRole('rowheader', { name: 'Row 3', exact: true })).toHaveCount(0);
  await expect(page.getByRole('columnheader', { name: 'Column C', exact: true })).toHaveCount(0);
  for (const name of ['Row 1', 'Row 2']) {
    await expect(page.getByRole('rowheader', { name, exact: true })).toHaveCount(1);
  }
  for (const name of ['Column A', 'Column B']) {
    await expect(page.getByRole('columnheader', { name, exact: true })).toHaveCount(1);
  }
  await expectRange(page, grid, 2, 2);
  await clickCell(page, grid, 1, 1);
  await expectActiveCell(page, 'A1', '101');
  await expectRange(page, grid, 1, 1);
  const scroll = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  const visibleCol = 3 + Math.ceil(scroll.x / grid.width);
  const visibleRow = 3 + Math.ceil(scroll.y / grid.height);
  const canvas = (await page.getByTestId('xlsx-scroll').locator('canvas').first().boundingBox())!;
  const x = grid.x + (visibleCol - 0.5) * grid.width - scroll.x;
  const y = grid.y + (visibleRow - 0.5) * grid.height - scroll.y;
  await page.mouse.click(canvas.x + x, canvas.y + grid.y + grid.height / 2);
  await expectActiveCell(page, `${column(visibleCol)}1`, '');
  await expectRect(page, { x: x - grid.width / 2, y: grid.y, width: grid.width, height: grid.height });
  await page.mouse.click(canvas.x + grid.x + grid.width / 2, canvas.y + y);
  await expectActiveCell(page, `A${visibleRow}`, '');
  await expectRect(page, { x: grid.x, y: y - grid.height / 2, width: grid.width, height: grid.height });
  await page.mouse.click(canvas.x + x, canvas.y + y);
  await expectActiveCell(page, `${column(visibleCol)}${visibleRow}`, '');
  await expectRect(page, { x: x - grid.width / 2, y: y - grid.height / 2, width: grid.width, height: grid.height });
  expect(await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }))).toEqual(scroll);
  await saveAndCheck(page, info, buffer, 'frozen', false, true);
});
