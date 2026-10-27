import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const columnWidth = 145;
const rowHeight = 40;
const canvas = (page: Page) => page.getByTestId('xlsx-scroll').locator('canvas');
type Rect = { x: number; y: number; width: number; height: number };
type Sheet = { name: string; data: string; before?: string; after?: string; rels?: string };

function xml(value: string) {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
}

function textCell(address: string, value: string) {
  return `<c r="${address}" t="inlineStr"><is><t>${xml(value)}</t></is></c>`;
}

function styles(formats: string[] = []) {
  return `<styleSheet xmlns="${main}">
    <numFmts count="${formats.length}">${formats.map((format, index) =>
      `<numFmt numFmtId="${164 + index}" formatCode="${xml(format)}"/>`).join('')}</numFmts>
    <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="${formats.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    ${formats.map((_, index) => `<xf numFmtId="${164 + index}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')}</cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  </styleSheet>`;
}

async function workbook(sheets: Sheet[], styleXml = styles()) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    ${sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
  </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${packageRelationships}"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${relationships}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>
    ${sheets.map((sheet, index) => `<sheet name="${xml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}
  </sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${packageRelationships}">
    ${sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${relationships}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('')}
    <Relationship Id="styles" Type="${relationships}/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file('xl/styles.xml', styleXml);
  sheets.forEach((sheet, index) => {
    zip.file(`xl/worksheets/sheet${index + 1}.xml`, `<worksheet xmlns="${main}" xmlns:r="${relationships}">
      ${sheet.before ?? '<sheetFormatPr defaultRowHeight="30" customHeight="1"/><cols><col min="1" max="50" width="20" customWidth="1"/></cols>'}
      <sheetData>${sheet.data}</sheetData>${sheet.after ?? ''}</worksheet>`);
    if (sheet.rels) zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, sheet.rels);
  });
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, file: string | Buffer) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(typeof file === 'string' ? file : {
    name: 'view.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: file,
  });
  await expect(canvas(page)).toBeVisible();
  await expect(canvas(page)).toHaveAttribute('data-paint-source', 'worker');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
}

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(`${label}.xlsx`);
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts = await Promise.all(Object.values(zip.files)
    .filter((part) => !part.dir && /\.(xml|rels)$/.test(part.name))
    .map((part) => part.async('string')));
  expect(await page.evaluate((parts) => parts.every((part) =>
    !new DOMParser().parseFromString(part, 'application/xml').querySelector('parsererror')), parts)).toBe(true);
  return { file, zip };
}

async function assertParts(actual: JSZip, expected: JSZip, parts: string[]) {
  for (const part of parts) {
    expect(actual.file(part), part).not.toBeNull();
    expect(await actual.file(part)!.async('string'), part).toBe(await expected.file(part)!.async('string'));
  }
}

async function roundTrip(page: Page, info: TestInfo, source: Buffer, parts?: string[]) {
  const expected = await JSZip.loadAsync(source);
  const originalSheet = await page.evaluate((source) => {
    const workbook = new DOMParser().parseFromString(source, 'application/xml');
    const active = Number(workbook.querySelector('workbookView')?.getAttribute('activeTab') ?? 0);
    return workbook.querySelectorAll('sheets > sheet')[active].getAttribute('name')!;
  }, await expected.file('xl/workbook.xml')!.async('string'));
  const tab = page.getByRole('tab', { name: originalSheet, exact: true });
  if (await tab.getAttribute('aria-selected') !== 'true') await tab.click();
  await expect(tab).toHaveAttribute('aria-selected', 'true');
  const preserved = parts ?? Object.keys(expected.files).filter((part) =>
    /^xl\/(worksheets\/|styles\.xml)/.test(part) && !expected.files[part].dir);
  const saved = await save(page, info, 'viewed');
  await assertParts(saved.zip, expected, preserved);
  await open(page, saved.file);
  await assertParts((await save(page, info, 'reopened')).zip, expected, preserved);
}

async function selectionBox(page: Page) {
  const selection = page.getByTestId('xlsx-selection');
  await expect(selection).toBeVisible();
  return (await selection.boundingBox())!;
}

async function zoom(page: Page, percent: number) {
  await page.getByTestId('xlsx-zoom').fill(`${percent}%`);
  await page.getByTestId('xlsx-zoom').press('Enter');
  await expect(canvas(page)).toHaveAttribute('data-worker-zoom', String(percent / 100));
  await expect(page.getByTestId('xlsx-zoom')).toHaveValue(`${percent}%`);
}

async function pixels(page: Page, rect: Rect) {
  return canvas(page).evaluate((element, rect) => {
    const surface = element as HTMLCanvasElement;
    const bounds = surface.getBoundingClientRect();
    const scale = surface.width / bounds.width;
    const x = Math.max(0, Math.round((rect.x - bounds.x) * scale));
    const y = Math.max(0, Math.round((rect.y - bounds.y) * scale));
    const width = Math.min(surface.width - x, Math.round(rect.width * scale));
    const height = Math.min(surface.height - y, Math.round(rect.height * scale));
    if (width <= 0 || height <= 0) throw new Error('Pixel region is outside the painted canvas');
    const data = surface.getContext('2d')!.getImageData(x, y, width, height).data;
    let dark = 0, red = 0, blue = 0, colored = 0, nonWhite = 0, hash = 2166136261;
    for (let index = 0; index < data.length; index += 4) {
      const r = data[index], g = data[index + 1], b = data[index + 2];
      if (Math.max(r, g, b) < 150) dark++;
      if (r > 150 && g < 100 && b < 100) red++;
      if (b > 120 && b > r + 40 && b > g + 20) blue++;
      if (Math.max(r, g, b) - Math.min(r, g, b) > 50) colored++;
      if (Math.min(r, g, b) < 240) nonWhite++;
      hash = Math.imul(hash ^ r, 16777619);
      hash = Math.imul(hash ^ g, 16777619);
      hash = Math.imul(hash ^ b, 16777619);
    }
    return { dark, red, blue, colored, nonWhite, hash };
  }, rect);
}

async function wheel(page: Page, dx: number, dy: number) {
  const box = (await canvas(page).boundingBox())!;
  const body = { x: box.x + box.width / 4, y: box.y + box.height / 4, width: box.width / 2, height: box.height / 2 };
  const before = (await pixels(page, body)).hash;
  const offset = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  await page.mouse.move(box.x + box.width - 50, box.y + box.height - 50);
  await page.mouse.wheel(dx, dy);
  await expect.poll(() => page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }))).not.toEqual(offset);
  await expect.poll(async () => (await pixels(page, body)).hash).not.toBe(before);
}

async function selectShowcaseChart(page: Page) {
  const box = (await canvas(page).boundingBox())!;
  const left = [16, 10, 14, 14, 10, 12, 8.43]
    .reduce((sum, width) => sum + Math.round((width * 7 + 5) / 7 * 256) / 256 * 7, 0);
  await page.mouse.click(box.x + left + 40, box.y + 80);
  await expect(page.getByTestId('xlsx-chart-selection')).toBeVisible();
  return (await page.getByTestId('xlsx-chart-selection').boundingBox())!;
}

async function linkWorkbook() {
  return workbook([
    { name: 'Sheet1', data: `<row r="1">${textCell('A1', 'External report')}</row><row r="2">${textCell('A2', 'Go to Sheet2')}</row>`,
      after: '<hyperlinks><hyperlink ref="A1" r:id="external" display="External report" tooltip="Open external report"/><hyperlink ref="A2" location="#Sheet2!A1" display="Go to Sheet2" tooltip="Jump to Sheet2"/></hyperlinks>',
      rels: `<Relationships xmlns="${packageRelationships}"><Relationship Id="external" Type="${relationships}/hyperlink" Target="https://example.invalid/report" TargetMode="External"/></Relationships>` },
    { name: 'Sheet2', data: `<row r="1">${textCell('A1', 'Internal destination')}</row>` },
  ]);
}

async function hiddenWorkbook() {
  return workbook([{ name: 'Hidden tracks',
    before: '<sheetFormatPr defaultRowHeight="30" customHeight="1"/><cols><col min="1" max="1" width="20"/><col min="2" max="2" width="20" hidden="1"/><col min="3" max="20" width="20"/></cols>',
    data: `<row r="1">${textCell('A1', 'Visible origin')}${textCell('B1', 'Hidden column')}${textCell('C1', 'Visible column')}</row>
      <row r="2" hidden="1">${textCell('A2', 'Hidden row')}</row><row r="3">${textCell('A3', 'Visible row')}</row>`,
  }]);
}

test.beforeEach(async ({ page }) => {
  await page.route('**/*', async (route) => {
    const url = new URL(route.request().url());
    if (['http:', 'https:'].includes(url.protocol) && !['127.0.0.1', 'localhost'].includes(url.hostname)) {
      await route.abort();
    } else await route.continue();
  });
});

test('xlsx view: showcase chart paints, selects, and deselects with Escape and a cell click', async ({ page }, info) => {
  const source = await readFile(resolve(root, 'apps/demo/public/showcase.xlsx'));
  await open(page, source);
  await expect(page.getByRole('img', { name: /Revenue by region/ })).toHaveCount(1);
  const chart = await selectShowcaseChart(page);
  await expect(page.getByTestId('xlsx-selection')).toHaveCount(0);
  expect((await pixels(page, { x: chart.x + 10, y: chart.y + 10, width: chart.width - 20, height: chart.height - 20 })).colored).toBeGreaterThan(100);
  await page.getByTestId('xlsx-scroll').press('Escape');
  await expect(page.getByTestId('xlsx-chart-selection')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-selection')).toBeVisible();
  await selectShowcaseChart(page);
  await canvas(page).click({ position: { x: 20, y: 110 } });
  await expect(page.getByTestId('xlsx-chart-selection')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A5');
  const zip = await JSZip.loadAsync(source);
  await roundTrip(page, info, source, Object.keys(zip.files).filter((part) =>
    /^xl\/(charts|drawings)\//.test(part) && !zip.files[part].dir));
  await selectShowcaseChart(page);
});

test('xlsx view: selected chart stays selected and aligned through scroll and zoom', async ({ page }, info) => {
  const archive = await JSZip.loadAsync(await readFile(resolve(root, 'apps/demo/public/showcase.xlsx')));
  const sheet = await archive.file('xl/worksheets/sheet1.xml')!.async('string');
  archive.file('xl/worksheets/sheet1.xml', sheet.replace('</sheetData>',
    '<row r="80"><c r="Z80"><v>1</v></c></row></sheetData>'));
  const source = await archive.generateAsync({ type: 'nodebuffer' });
  await open(page, source);
  const surface = (await canvas(page).boundingBox())!;
  const before = await selectShowcaseChart(page);
  const chart = page.getByTestId('xlsx-chart-selection');
  const id = await chart.getAttribute('data-chart-id');
  await wheel(page, 100, 60);
  await expect(chart).toHaveAttribute('data-chart-id', id!);
  const scroll = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  expect(scroll.x + scroll.y).toBeGreaterThan(0);
  await expect.poll(async () => Math.round((await chart.boundingBox())!.x)).toBe(Math.round(before.x - scroll.x));
  await expect.poll(async () => Math.round((await chart.boundingBox())!.y)).toBe(Math.round(before.y - scroll.y));
  await zoom(page, 150);
  await expect(chart).toHaveAttribute('data-chart-id', id!);
  await expect.poll(async () => (await chart.boundingBox())!.width).toBeCloseTo(before.width * 1.5, 0);
  await expect.poll(async () => (await chart.boundingBox())!.height).toBeCloseTo(before.height * 1.5, 0);
  const zoomedScroll = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  await expect.poll(async () => (await chart.boundingBox())!.x).toBeCloseTo(surface.x + (before.x - surface.x) * 1.5 - zoomedScroll.x, 0);
  await expect.poll(async () => (await chart.boundingBox())!.y).toBeCloseTo(surface.y + (before.y - surface.y) * 1.5 - zoomedScroll.y, 0);
  await zoom(page, 100);
  await expect(chart).toHaveAttribute('data-chart-id', id!);
  await expect.poll(async () => (await chart.boundingBox())!.width).toBeCloseTo(before.width, 0);
  await roundTrip(page, info, source, ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml']);
});

for (const modifier of [false, true]) {
  test(`xlsx view: ${modifier ? 'Ctrl-click' : 'click'} follows a painted internal hyperlink to Sheet2 A1`, async ({ page }, info) => {
    const source = await linkWorkbook();
    await open(page, source);
    const origin = await selectionBox(page);
    const link = { x: origin.x + 30, y: origin.y + rowHeight + 20 };
    await page.mouse.move(link.x, link.y);
    await expect(page.getByTestId('xlsx-scroll')).toHaveCSS('cursor', 'pointer');
    await expect(page.getByTestId('xlsx-scroll')).toHaveAttribute('title', 'Jump to Sheet2');
    expect((await pixels(page, { x: origin.x + 2, y: origin.y + rowHeight + 2, width: columnWidth - 4, height: rowHeight - 4 })).blue).toBeGreaterThan(20);
    if (modifier) await page.keyboard.down('Control');
    await page.mouse.click(link.x, link.y);
    if (modifier) await page.keyboard.up('Control');
    await expect(page.getByRole('tab', { name: 'Sheet2', exact: true })).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Internal destination');
    await expect(page.getByRole('gridcell', { name: 'A1, Internal destination, selected', exact: true })).toHaveCount(1);
    await expect(page).toHaveURL(/\?format=xlsx$/);
    const followed = await save(page, info, 'followed-link');
    const savedView = await page.evaluate(([book, sheet]) => {
      const parser = new DOMParser();
      return [
        parser.parseFromString(book, 'application/xml').querySelector('workbookView')?.getAttribute('activeTab'),
        parser.parseFromString(sheet, 'application/xml').querySelector('sheetView')?.getAttribute('tabSelected'),
      ];
    }, await Promise.all(['xl/workbook.xml', 'xl/worksheets/sheet2.xml']
      .map((part) => followed.zip.file(part)!.async('string'))));
    expect(savedView).toEqual(['1', '1']);
    await roundTrip(page, info, source);
    await open(page, followed.file);
    await expect(page.getByRole('tab', { name: 'Sheet2', exact: true })).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Internal destination');
  });
}

test('xlsx view: external hyperlink opens a separate window on click and Ctrl-click', async ({ page }, info) => {
  await page.addInitScript(() => {
    const calls: { url: string; target: string | undefined; features: string | undefined }[] = [];
    Object.assign(window, { __viewOpenCalls: calls });
    window.open = (url, target, features) => {
      calls.push({ url: String(url), target, features });
      return null;
    };
  });
  const source = await linkWorkbook();
  await open(page, source);
  const origin = await selectionBox(page);
  await page.mouse.move(origin.x + 25, origin.y + 20);
  await expect(page.getByTestId('xlsx-scroll')).toHaveCSS('cursor', 'pointer');
  await expect(page.getByTestId('xlsx-scroll')).toHaveAttribute('title', 'Open external report');
  expect((await pixels(page, { x: origin.x + 2, y: origin.y + 2, width: columnWidth - 4, height: rowHeight - 4 })).blue).toBeGreaterThan(20);
  for (const modifier of [false, true]) {
    if (modifier) await page.keyboard.down('Control');
    await page.mouse.click(origin.x + 25, origin.y + 20);
    if (modifier) await page.keyboard.up('Control');
    await expect.poll(() => page.evaluate(() => Reflect.get(window, '__viewOpenCalls'))).toEqual(
      Array.from({ length: modifier ? 2 : 1 }, () => ({ url: 'https://example.invalid/report', target: '_blank', features: 'noopener,noreferrer' }))
    );
    await expect(page).toHaveURL(/\?format=xlsx$/);
    await expect(page.getByRole('tab', { name: 'Sheet1', exact: true })).toHaveAttribute('aria-selected', 'true');
  }
  await roundTrip(page, info, source);
});

test('xlsx view: hidden rows and columns occupy no painted space and survive reopening', async ({ page }, info) => {
  const source = await hiddenWorkbook();
  await open(page, source);
  const origin = await selectionBox(page);
  await expect(page.getByRole('gridcell').filter({ hasText: 'Hidden column' })).toHaveCount(0);
  await expect(page.getByRole('gridcell').filter({ hasText: 'Hidden row' })).toHaveCount(0);
  await expect(page.getByRole('columnheader', { name: 'Column B', exact: true })).toHaveCount(0);
  await expect(page.getByRole('rowheader', { name: 'Row 2', exact: true })).toHaveCount(0);
  await page.mouse.click(origin.x + columnWidth + 20, origin.y + 20);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('C1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Visible column');
  await page.mouse.click(origin.x + 20, origin.y + rowHeight + 20);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A3');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Visible row');
  await roundTrip(page, info, source);
  await expect(page.getByRole('gridcell').filter({ hasText: 'Hidden row' })).toHaveCount(0);
});

for (const [key, address, reverse, track] of [
  ['ArrowRight', 'C1', 'ArrowLeft', 'columns'], ['Tab', 'C1', 'Shift+Tab', 'columns'],
  ['ArrowDown', 'A3', 'ArrowUp', 'rows'], ['Enter', 'A3', 'Shift+Enter', 'rows'],
]) {
  test(`xlsx view: ${key} and ${reverse} navigation skip hidden ${track}`, async ({ page }, info) => {
    const source = await hiddenWorkbook();
    await open(page, source);
    const scroll = page.getByTestId('xlsx-scroll');
    await scroll.press(key);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue(address);
    await scroll.press(reverse);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await roundTrip(page, info, source);
  });
}

test('xlsx view: file column widths and row heights scale accurately with zoom', async ({ page }, info) => {
  const source = await workbook([{ name: 'Dimensions',
    before: '<sheetFormatPr defaultRowHeight="30" customHeight="1"/><cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/></cols>',
    data: `<row r="1" ht="30" customHeight="1">${textCell('A1', 'Narrow')}${textCell('B1', 'Wide')}</row><row r="2" ht="60" customHeight="1">${textCell('A2', 'Tall')}</row>`,
  }]);
  await open(page, source);
  const a1 = await selectionBox(page);
  expect(a1.width).toBeCloseTo(75, 0);
  expect(a1.height).toBeCloseTo(40, 0);
  await page.getByTestId('xlsx-scroll').press('ArrowRight');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  const b1 = await selectionBox(page);
  expect(b1.x - a1.x).toBeCloseTo(75, 0);
  expect(b1.width).toBeCloseTo(145, 0);
  await page.mouse.click(a1.x + 20, a1.y + 60);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A2');
  const a2 = await selectionBox(page);
  expect(a2.y - a1.y).toBeCloseTo(40, 0);
  expect(a2.height).toBeCloseTo(80, 0);
  await zoom(page, 150);
  await expect.poll(async () => (await selectionBox(page)).width).toBeCloseTo(112.5, 0);
  await expect.poll(async () => (await selectionBox(page)).height).toBeCloseTo(120, 0);
  await zoom(page, 100);
  await roundTrip(page, info, source);
  expect((await selectionBox(page)).width).toBeCloseTo(75, 0);
});

test('xlsx view: merged cells have no interior grid lines and select the whole merge', async ({ page }, info) => {
  const source = await workbook([{ name: 'Merged',
    data: `<row r="1">${textCell('A1', 'Origin')}</row><row r="2">${textCell('B2', 'Merged heading')}</row>`,
    after: '<mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>',
  }]);
  await open(page, source);
  const origin = await selectionBox(page);
  const seam = { x: origin.x + 2 * columnWidth - 1, y: origin.y + 2 * rowHeight + 8, width: 3, height: 12 };
  expect((await pixels(page, seam)).nonWhite).toBe(0);
  await expect(page.getByRole('gridcell').filter({ hasText: 'Merged heading' })).toHaveCount(1);
  await page.mouse.click(origin.x + 2.5 * columnWidth, origin.y + 2.5 * rowHeight);
  let selected = await selectionBox(page);
  expect(selected.x - origin.x).toBeCloseTo(columnWidth, 0);
  expect(selected.y - origin.y).toBeCloseTo(rowHeight, 0);
  expect(selected.width).toBeCloseTo(columnWidth * 2, 0);
  expect(selected.height).toBeCloseTo(rowHeight * 2, 0);
  await page.mouse.click(origin.x + columnWidth / 2, origin.y + rowHeight / 2);
  await page.getByTestId('xlsx-scroll').press('ArrowRight');
  await page.getByTestId('xlsx-scroll').press('ArrowDown');
  selected = await selectionBox(page);
  expect(selected.width).toBeCloseTo(columnWidth * 2, 0);
  expect(selected.height).toBeCloseTo(rowHeight * 2, 0);
  await roundTrip(page, info, source);
  expect((await pixels(page, seam)).nonWhite).toBe(0);
});

test('xlsx view: frozen row and column stay painted and clickable during diagonal scrolling', async ({ page }, info) => {
  const source = await workbook([{ name: 'Frozen',
    before: '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="30" customHeight="1"/><cols><col min="1" max="20" width="20"/></cols>',
    data: Array.from({ length: 70 }, (_, row) => `<row r="${row + 1}">${
      Array.from({ length: 20 }, (_, col) => textCell(`${String.fromCharCode(65 + col)}${row + 1}`, row === 0 && col === 0 ? 'Frozen corner' : `R${row + 1}C${col + 1}`)).join('')
    }</row>`).join(''),
  }]);
  await open(page, source);
  const origin = await selectionBox(page);
  const corner = { x: origin.x + 2, y: origin.y + 2, width: columnWidth - 4, height: rowHeight - 4 };
  const top = { x: origin.x + columnWidth + 2, y: origin.y + 2, width: columnWidth - 4, height: rowHeight - 4 };
  const left = { x: origin.x + 2, y: origin.y + rowHeight + 2, width: columnWidth - 4, height: rowHeight - 4 };
  const before = await pixels(page, corner);
  await wheel(page, 0, 400);
  await expect(page.getByRole('rowheader', { name: 'Row 2', exact: true })).toHaveCount(0);
  expect((await pixels(page, corner)).hash).toBe(before.hash);
  await expect(page.getByRole('gridcell', { name: 'B1, R1C2', exact: true })).toHaveCount(1);
  const topBefore = await pixels(page, top);
  await wheel(page, 0, 200);
  expect((await pixels(page, top)).hash).toBe(topBefore.hash);
  const leftBefore = await pixels(page, left);
  await wheel(page, 300, 0);
  await expect(page.getByRole('columnheader', { name: 'Column B', exact: true })).toHaveCount(0);
  expect((await pixels(page, left)).hash).toBe(leftBefore.hash);
  expect((await pixels(page, corner)).hash).toBe(before.hash);
  const offset = await page.getByTestId('xlsx-scroll').evaluate((element) => ({ x: element.scrollLeft, y: element.scrollTop }));
  expect(offset.x).toBeGreaterThan(0);
  expect(offset.y).toBeGreaterThan(0);
  await page.mouse.click(origin.x + 20, origin.y + 20);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Frozen corner');
  await roundTrip(page, info, source);
});

test('xlsx view: date, percent, currency, red negative, and custom number formats display Excel text', async ({ page }, info) => {
  const formats = ['yyyy-mm-dd', '0.0%', '$#,##0.00', '#,##0.00;[Red](#,##0.00)', '"Units: "0.00'];
  const values = [45292, 0.125, 1234.5, -1234.5, 12.5];
  const displays = ['2024-01-01', '12.5%', '$1,234.50', '(1,234.50)', 'Units: 12.50'];
  const source = await workbook([{ name: 'Formats', data: values.map((value, index) =>
    `<row r="${index + 1}"><c r="A${index + 1}" s="${index + 1}"><v>${value}</v></c></row>`).join(''),
  }], styles(formats));
  await open(page, source);
  const origin = await selectionBox(page);
  for (const [index, display] of displays.entries()) {
    await expect(page.getByRole('gridcell', { name: new RegExp(`^A${index + 1}, `) })).toContainText(display);
    await canvas(page).click({ position: { x: 30, y: index * rowHeight + 20 } });
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue(`A${index + 1}`);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(String(values[index]));
    await expect(page.getByRole('gridcell', { name: new RegExp(`^A${index + 1}, `) })).toHaveText(`A${index + 1}, ${display}, selected`);
    expect((await pixels(page, { x: origin.x + 2, y: origin.y + index * rowHeight + 2, width: columnWidth - 4, height: rowHeight - 4 })).nonWhite).toBeGreaterThan(20);
  }
  expect((await pixels(page, { x: origin.x + 2, y: origin.y + 3 * rowHeight + 2, width: columnWidth - 4, height: rowHeight - 4 })).red).toBeGreaterThan(20);
  expect((await pixels(page, { x: origin.x + 2, y: origin.y + 2 * rowHeight + 2, width: columnWidth - 4, height: rowHeight - 4 })).red).toBe(0);
  await roundTrip(page, info, source);
  for (const [index, display] of displays.entries()) {
    await expect(page.getByRole('gridcell', { name: new RegExp(`^A${index + 1}, `) })).toContainText(display);
  }
});

test('xlsx view: long text overflows empty neighbours and clips after filling a neighbour with undo and redo', async ({ page }, info) => {
  const longText = 'MMMMMMMMMMMMMMMMMMMMMMMM';
  const source = await workbook([{ name: 'Overflow', data: `<row r="1">${textCell('A1', longText)}</row>` }]);
  await open(page, source);
  const origin = await selectionBox(page);
  const spill = { x: origin.x + columnWidth + 10, y: origin.y + 2, width: 50, height: rowHeight - 4 };
  expect((await pixels(page, spill)).dark).toBeGreaterThan(30);
  await canvas(page).click({ position: { x: columnWidth + 25, y: 20 } });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  await page.getByTestId('xlsx-scroll').press('9');
  await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue('9');
  await page.getByTestId('xlsx-cell-editor').press('Enter');
  await expect(page.getByRole('gridcell', { name: /^B1, 9(?:, selected)?$/ })).toHaveCount(1);
  await expect.poll(async () => (await pixels(page, spill)).dark).toBe(0);
  const edited = await save(page, info, 'clipped');
  expect(await edited.zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(/<c\b[^>]*r="B1"[^>]*>[\s\S]*?<v>9<\/v>/);
  await page.getByTestId('xlsx-undo').click();
  await expect.poll(async () => (await pixels(page, spill)).dark).toBeGreaterThan(30);
  const undone = await save(page, info, 'overflow-restored');
  expect(await undone.zip.file('xl/worksheets/sheet1.xml')!.async('string')).not.toMatch(/<v>9<\/v>/);
  await page.getByTestId('xlsx-redo').click();
  await expect.poll(async () => (await pixels(page, spill)).dark).toBe(0);
  const redone = await save(page, info, 'clipped-again');
  expect(await redone.zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(/<c\b[^>]*r="B1"[^>]*>[\s\S]*?<v>9<\/v>/);
  await open(page, redone.file);
  await expect(page.getByRole('gridcell', { name: /^A1, / })).toContainText(longText);
  expect((await pixels(page, spill)).dark).toBe(0);
  await assertParts((await save(page, info, 'reopened')).zip, redone.zip, ['xl/worksheets/sheet1.xml', 'xl/styles.xml']);
});

test('xlsx view: thirty sheet tabs scroll into view and select every sheet', async ({ page }, info) => {
  await page.setViewportSize({ width: 800, height: 900 });
  const source = await workbook(Array.from({ length: 30 }, (_, index) => ({
    name: `Worksheet ${String(index + 1).padStart(2, '0')}`,
    data: `<row r="1">${textCell('A1', `Content of sheet ${index + 1}`)}</row>`,
  })));
  await open(page, source);
  const tabs = page.getByTestId('xlsx-sheet-tabs');
  await expect(tabs.getByRole('tab')).toHaveCount(30);
  expect(await tabs.evaluate((element) => element.scrollWidth > element.clientWidth)).toBe(true);
  for (let index = 1; index <= 30; index++) {
    const tab = tabs.getByRole('tab', { name: `Worksheet ${String(index).padStart(2, '0')}`, exact: true });
    await tab.click();
    await expect(tab).toHaveAttribute('aria-selected', 'true');
    await expect(tabs.locator('[aria-selected="true"]')).toHaveCount(1);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(`Content of sheet ${index}`);
    await expect(page.getByRole('gridcell', { name: `A1, Content of sheet ${index}, selected`, exact: true })).toHaveCount(1);
  }
  expect(await tabs.evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  await tabs.getByRole('tab', { name: 'Worksheet 01', exact: true }).click();
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Content of sheet 1');
  await roundTrip(page, info, source);
  await expect(page.getByTestId('xlsx-sheet-tabs').getByRole('tab')).toHaveCount(30);
  await page.getByRole('tab', { name: 'Worksheet 30', exact: true }).click();
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Content of sheet 30');
});
