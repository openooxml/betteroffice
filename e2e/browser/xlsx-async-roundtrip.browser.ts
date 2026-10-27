import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const styles = `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
const firstRow = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>12.5</v></c><c r="C1" s="1"><v>45292</v></c><c r="D1"><f>B1*2</f><v>25</v></c><c r="E1" t="s"><v>1</v></c><c r="F1" t="s"><v>2</v></c></row>';
type Sheet = { name: string; rows: string; beforeData?: string; afterData?: string };
type Upload = { name: string; mimeType: string; buffer: Buffer };
type Parts = Record<string, string>;

async function fixture(options: { name?: string; sheets?: Sheet[]; definedNames?: string } = {}): Promise<Upload> {
  const sheets = options.sheets ?? [
    { name: 'Entries', rows: firstRow },
    { name: 'Other', rows: '<row r="1"><c r="A1" t="s"><v>3</v></c><c r="B1"><v>7</v></c></row>' },
  ];
  const zip = new JSZip();
  zip.file('[Content_Types].xml', declaration + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' + sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>');
  zip.file('_rels/.rels', declaration + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', declaration + `<workbook xmlns="${main}" xmlns:r="${relationships}"><workbookPr date1904="0"/><bookViews><workbookView activeTab="0"/></bookViews><sheets>` + sheets.map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('') + `</sheets>${options.definedNames ?? ''}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', declaration + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${relationships}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('') + `<Relationship Id="rIdStyles" Type="${relationships}/styles" Target="styles.xml"/><Relationship Id="rIdStrings" Type="${relationships}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
  zip.file('xl/styles.xml', declaration + styles);
  const indexes = sheets.flatMap((sheet) => Array.from(sheet.rows.matchAll(/<c\b[^>]*\bt="s"[^>]*>\s*<v>(\d+)<\/v>/g), (match) => Number(match[1])));
  const strings = ['Original text', 'Format me', 'Clear me', 'Other original'].slice(0, Math.max(-1, ...indexes) + 1);
  zip.file('xl/sharedStrings.xml', declaration + `<sst xmlns="${main}" count="${indexes.length}" uniqueCount="${strings.length}">${strings.map((text) => `<si><t>${text}</t></si>`).join('')}</sst>`);
  sheets.forEach((sheet, index) => zip.file(`xl/worksheets/sheet${index + 1}.xml`, declaration + `<worksheet xmlns="${main}" xmlns:r="${relationships}">${sheet.beforeData ?? ''}<sheetData>${sheet.rows}</sheetData>${sheet.afterData ?? ''}</worksheet>`));
  return { name: options.name ?? 'roundtrip.xlsx', mimeType, buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function xmlParts(bytes: Buffer): Promise<Parts> {
  const zip = await JSZip.loadAsync(bytes);
  const entries = await Promise.all(Object.values(zip.files)
    .filter((part) => !part.dir && /\.(xml|rels)$/.test(part.name))
    .map(async (part) => [part.name, await part.async('string')] as const));
  return Object.fromEntries(entries);
}

async function open(page: Page, file?: string | Upload) {
  await page.goto('/?format=xlsx');
  await replace(page, file ?? await fixture());
}

async function replace(page: Page, file: string | Upload) {
  await page.locator('input[type=file]').last().setInputFiles(file);
  await expect(page.locator('.document-name')).toHaveText(typeof file === 'string' ? file.split('/').pop()! : file.name);
  await expect(page.locator('.editor-stage canvas[data-paint-source="worker"]').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
}

async function select(page: Page, address: string) {
  const match = /^([A-Z]+)(\d+)$/.exec(address)!;
  const column = [...match[1]].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  for (let col = 1; col < column; col++) await scroll.press('ArrowRight');
  for (let row = 1; row < Number(match[2]); row++) await scroll.press('ArrowDown');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(address);
}

async function inputAt(page: Page, address: string, value: string) {
  await select(page, address);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(value);
}

async function lastRow(page: Page, row: number) {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  await scroll.press('ControlOrMeta+ArrowDown');
  const name = page.getByTestId('xlsx-name-box');
  await expect(name).not.toHaveValue('A1');
  const bottom = Number((await name.inputValue()).slice(1));
  expect(bottom).toBeGreaterThanOrEqual(row);
  expect(bottom - row).toBeLessThan(10);
  // The scroll extent includes blank rows beyond the last populated row.
  for (let current = bottom; current > row; current--) await scroll.press('ArrowUp');
  await expect(name).toHaveValue(`A${row}`);
}

async function cellText(page: Page, address: string, value: string) {
  const text = `${address}, ${value || 'empty'}`.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  await expect(page.getByRole('gridcell').filter({ hasText: new RegExp(`^${text}(?:, selected)?$`) })).toHaveCount(1);
}

async function edit(page: Page, address: string, value: string) {
  await select(page, address);
  await page.getByTestId('xlsx-scroll').press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+A');
  await input.pressSequentially(value);
  await expect(input).toHaveValue(value);
  await input.press('Enter');
  await expect(input).toHaveCount(0);
  await inputAt(page, address, value);
}

async function history(page: Page, address: string, before: string, after: string) {
  await page.getByTestId('xlsx-undo').click();
  await inputAt(page, address, before);
  await page.getByTestId('xlsx-redo').click();
  await inputAt(page, address, after);
}

async function save(page: Page, info: TestInfo, label: string, shortcut = false) {
  const downloading = page.waitForEvent('download');
  if (shortcut) await page.getByTestId('xlsx-scroll').press('ControlOrMeta+s');
  else await page.locator('.title-actions').getByRole('button', { name: 'Save', exact: true }).click();
  const download = await downloading;
  expect(await download.failure()).toBeNull();
  const file = info.outputPath(`${label}.xlsx`);
  await download.saveAs(file);
  await expect(page.locator('.document-name')).not.toContainText(' · Edited');
  await expect(page.locator('.notice')).toHaveCount(0);
  return { file, parts: await xmlParts(await readFile(file)) };
}

async function inspect(page: Page, parts: Parts) {
  return page.evaluate((parts) => {
    const parse = (xml: string) => {
      const document = new DOMParser().parseFromString(xml, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error('Invalid workbook XML');
      return document;
    };
    const documents = Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, parse(xml)]));
    const strings = Array.from(documents['xl/sharedStrings.xml']?.getElementsByTagNameNS('*', 'si') ?? [])
      .map((item) => Array.from(item.getElementsByTagNameNS('*', 't')).map((text) => text.textContent ?? '').join(''));
    const styles = documents['xl/styles.xml'];
    const formats = Array.from(styles.getElementsByTagNameNS('*', 'cellXfs')[0].children);
    const fonts = Array.from(styles.getElementsByTagNameNS('*', 'fonts')[0].children);
    const fills = Array.from(styles.getElementsByTagNameNS('*', 'fills')[0].children);
    const borders = Array.from(styles.getElementsByTagNameNS('*', 'borders')[0].children);
    const serialize = (node: Element | undefined) => node ? new XMLSerializer().serializeToString(node) : '';
    return Object.fromEntries(Object.entries(documents)
      .filter(([name]) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))
      .map(([name, document]) => [name, Object.fromEntries(Array.from(document.getElementsByTagNameNS('*', 'c')).map((cell) => {
        const type = cell.getAttribute('t') ?? 'n';
        const raw = cell.getElementsByTagNameNS('*', 'v')[0]?.textContent ?? '';
        const format = formats[Number(cell.getAttribute('s') ?? 0)];
        const numFmtId = format.getAttribute('numFmtId') ?? '0';
        const font = fonts[Number(format.getAttribute('fontId') ?? 0)];
        return [cell.getAttribute('r')!, {
          value: type === 's' ? strings[Number(raw)] : type === 'inlineStr'
            ? Array.from(cell.getElementsByTagNameNS('*', 't')).map((text) => text.textContent ?? '').join('') : raw,
          type: type === 's' || type === 'inlineStr' ? 'text' : type,
          formula: cell.getElementsByTagNameNS('*', 'f')[0]?.textContent ?? null,
          style: {
            numFmtId,
            formatCode: Array.from(styles.getElementsByTagNameNS('*', 'numFmt')).find((item) => item.getAttribute('numFmtId') === numFmtId)?.getAttribute('formatCode') ?? null,
            bold: font.getElementsByTagNameNS('*', 'b').length > 0,
            font: serialize(font),
            fill: serialize(fills[Number(format.getAttribute('fillId') ?? 0)]),
            border: serialize(borders[Number(format.getAttribute('borderId') ?? 0)]),
            alignment: serialize(format.getElementsByTagNameNS('*', 'alignment')[0]),
          },
        }];
      }))]));
  }, parts);
}

function stableParts(parts: Parts) {
  return Object.fromEntries(Object.entries(parts).filter(([name]) => /^xl\/(worksheets\/sheet\d+\.xml|styles\.xml|sharedStrings\.xml)$/.test(name)));
}

function elementXml(xml: string, name: string) {
  return xml.match(new RegExp(`<${name}(?:\\s[^>]*?)?(?:\\/>|>[\\s\\S]*?<\\/${name}>)`))?.[0] ?? null;
}

test('xlsx roundtrip: mixed edits on two sheets preserve values, formulas, dates, and formatting', async ({ page }, info) => {
  await open(page);
  await edit(page, 'A1', 'Roundtrip text & <xml> café');
  await cellText(page, 'A1', 'Roundtrip text & <xml> café');
  await history(page, 'A1', 'Original text', 'Roundtrip text & <xml> café');
  await edit(page, 'B1', '42.5');
  await cellText(page, 'B1', '42.5');
  await history(page, 'B1', '12.5', '42.5');
  await edit(page, 'C1', '45306');
  await cellText(page, 'C1', '1/15/24');
  await history(page, 'C1', '45292', '45306');
  await select(page, 'D1');
  await page.getByTestId('xlsx-formula-input').fill('=SUM(B1,7.5)');
  await page.getByTestId('xlsx-formula-input').press('Enter');
  await inputAt(page, 'D1', '=SUM(B1,7.5)');
  await cellText(page, 'D1', '50');
  await history(page, 'D1', '=B1*2', '=SUM(B1,7.5)');
  await select(page, 'E1');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+b');
  const bold = page.getByTestId('xlsx-formatting-toolbar').getByRole('button', { name: 'Bold', exact: true });
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.getByTestId('xlsx-undo').click();
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await page.getByTestId('xlsx-redo').click();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await select(page, 'F1');
  await page.getByTestId('xlsx-scroll').press('Delete');
  await inputAt(page, 'F1', '');
  await cellText(page, 'F1', '');
  await history(page, 'F1', 'Clear me', '');
  await page.getByRole('tab', { name: 'Other', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true');
  await edit(page, 'A1', 'Second sheet edit');
  await cellText(page, 'A1', 'Second sheet edit');
  await history(page, 'A1', 'Other original', 'Second sheet edit');
  const first = await save(page, info, 'mixed');
  const model = await inspect(page, first.parts);
  expect(model['xl/worksheets/sheet1.xml']).toMatchObject({
    A1: { value: 'Roundtrip text & <xml> café', type: 'text' },
    B1: { value: '42.5', type: 'n' },
    C1: { value: '45306', type: 'n', style: { numFmtId: '14' } },
    D1: { formula: 'SUM(B1,7.5)', value: '50' },
    E1: { value: 'Format me', style: { bold: true } },
  });
  expect(model['xl/worksheets/sheet1.xml'].F1?.value ?? '').toBe('');
  expect(model['xl/worksheets/sheet2.xml'].A1.value).toBe('Second sheet edit');
  await open(page, first.file);
  await page.getByRole('tab', { name: 'Entries', exact: true }).click();
  await inputAt(page, 'A1', 'Roundtrip text & <xml> café');
  await cellText(page, 'A1', 'Roundtrip text & <xml> café');
  await inputAt(page, 'B1', '42.5');
  await inputAt(page, 'C1', '45306');
  await cellText(page, 'C1', '1/15/24');
  await inputAt(page, 'D1', '=SUM(B1,7.5)');
  await cellText(page, 'D1', '50');
  await select(page, 'E1');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await inputAt(page, 'F1', '');
  await page.getByRole('tab', { name: 'Other', exact: true }).click();
  await inputAt(page, 'A1', 'Second sheet edit');
  const second = await save(page, info, 'mixed-reopened');
  expect(await inspect(page, second.parts)).toEqual(model);
});

test('xlsx roundtrip: unchanged reopen keeps sheet, styles, and shared strings XML byte-identical', async ({ page }, info) => {
  await open(page);
  await edit(page, 'A1', 'Stable shared text');
  await cellText(page, 'A1', 'Stable shared text');
  await select(page, 'A1');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+b');
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  const first = await save(page, info, 'stable-first');
  expect(first.parts['xl/sharedStrings.xml']).toBeDefined();
  await open(page, first.file);
  await inputAt(page, 'A1', 'Stable shared text');
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  const second = await save(page, info, 'stable-second');
  expect(stableParts(second.parts)).toEqual(stableParts(first.parts));
});

test('xlsx roundtrip: cleared content stays empty while its date style survives undo, redo, and reopen', async ({ page }, info) => {
  await open(page);
  await inputAt(page, 'C1', '45292');
  await page.getByTestId('xlsx-scroll').press('Backspace');
  await inputAt(page, 'C1', '');
  await cellText(page, 'C1', '');
  await history(page, 'C1', '45292', '');
  const cleared = await save(page, info, 'cleared-date');
  const model = await inspect(page, cleared.parts);
  expect(model['xl/worksheets/sheet1.xml'].C1).toMatchObject({ value: '', formula: null, style: { numFmtId: '14' } });
  await open(page, cleared.file);
  await inputAt(page, 'C1', '');
  await cellText(page, 'C1', '');
  const reopened = await save(page, info, 'cleared-date-reopened');
  expect(await inspect(page, reopened.parts)).toEqual(model);
});

test('xlsx roundtrip: currency formatting survives menu selection, undo, redo, and reopen', async ({ page }, info) => {
  await open(page);
  await select(page, 'B1');
  const formats = page.getByRole('button', { name: 'More number formats', exact: true });
  await formats.click();
  const menu = page.getByRole('menu', { name: 'More number formats', exact: true });
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitemradio', { name: 'Currency', exact: true }).click();
  await expect(menu).toBeHidden();
  await cellText(page, 'B1', '$12.50');
  await page.getByTestId('xlsx-undo').click();
  await cellText(page, 'B1', '12.5');
  await page.getByTestId('xlsx-redo').click();
  await cellText(page, 'B1', '$12.50');
  const first = await save(page, info, 'currency');
  const model = await inspect(page, first.parts);
  expect(model['xl/worksheets/sheet1.xml'].B1.value).toBe('12.5');
  expect(model['xl/worksheets/sheet1.xml'].B1.style.numFmtId).not.toBe('0');
  await open(page, first.file);
  await inputAt(page, 'B1', '12.5');
  await cellText(page, 'B1', '$12.50');
  await formats.click();
  await expect(menu.getByRole('menuitemradio', { name: 'Currency', exact: true })).toHaveAttribute('aria-checked', 'true');
  await menu.press('Escape');
  await expect(menu).toBeHidden();
  const second = await save(page, info, 'currency-reopened');
  expect(await inspect(page, second.parts)).toEqual(model);
});

test('xlsx roundtrip: edits mark the document dirty and saving resets it without disabling Save', async ({ page }, info) => {
  await open(page);
  const title = page.locator('.document-name');
  const button = page.locator('.title-actions').getByRole('button', { name: 'Save', exact: true });
  await expect(title).toHaveText('roundtrip.xlsx');
  await expect(button).toBeEnabled();
  await edit(page, 'A1', 'Dirty text');
  await cellText(page, 'A1', 'Dirty text');
  await expect(title).toHaveText('roundtrip.xlsx · Edited');
  await expect(button).toBeEnabled();
  await history(page, 'A1', 'Original text', 'Dirty text');
  const first = await save(page, info, 'clean');
  expect((await inspect(page, first.parts))['xl/worksheets/sheet1.xml'].A1.value).toBe('Dirty text');
  await expect(title).toHaveText('roundtrip.xlsx');
  await expect(button).toBeEnabled();
  await edit(page, 'B1', '99');
  await cellText(page, 'B1', '99');
  await expect(title).toHaveText('roundtrip.xlsx · Edited');
  const second = await save(page, info, 'clean-again', true);
  expect((await inspect(page, second.parts))['xl/worksheets/sheet1.xml'].B1.value).toBe('99');
  await expect(title).toHaveText('roundtrip.xlsx');
  await expect(button).toBeEnabled();
});

test('xlsx roundtrip: saving twice without edits downloads identical workbook XML', async ({ page }, info) => {
  await open(page);
  await edit(page, 'A1', 'Save twice');
  await cellText(page, 'A1', 'Save twice');
  const first = await save(page, info, 'twice-first');
  const second = await save(page, info, 'twice-second');
  expect(second.parts).toEqual(first.parts);
  await inputAt(page, 'A1', 'Save twice');
  await page.getByTestId('xlsx-undo').click();
  await inputAt(page, 'A1', 'Original text');
  await page.getByTestId('xlsx-redo').click();
  await inputAt(page, 'A1', 'Save twice');
  const redone = await save(page, info, 'twice-redone');
  expect(await inspect(page, redone.parts)).toEqual(await inspect(page, first.parts));
});

for (const kind of ['corrupt', 'empty'] as const) {
  test(`xlsx roundtrip: refusing a ${kind} file keeps the current workbook and its undo history`, async ({ page }, info) => {
    await open(page);
    await edit(page, 'A1', 'Keep this workbook');
    await cellText(page, 'A1', 'Keep this workbook');
    const before = await save(page, info, `${kind}-before`);
    await page.locator('input[type=file]').last().setInputFiles({
      name: `${kind}.xlsx`, mimeType,
      buffer: kind === 'empty' ? Buffer.alloc(0) : Buffer.from('not an OOXML archive'),
    });
    await expect(page.locator('.notice')).toBeVisible();
    await expect(page.locator('.document-name')).toHaveText('roundtrip.xlsx');
    await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
    await inputAt(page, 'A1', 'Keep this workbook');
    await cellText(page, 'A1', 'Keep this workbook');
    await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
    const after = await save(page, info, `${kind}-after`);
    expect(after.parts).toEqual(before.parts);
    await page.getByTestId('xlsx-undo').click();
    await inputAt(page, 'A1', 'Original text');
    await page.getByTestId('xlsx-redo').click();
    await inputAt(page, 'A1', 'Keep this workbook');
    const restored = await save(page, info, `${kind}-restored`);
    expect(await inspect(page, restored.parts)).toEqual(await inspect(page, before.parts));
  });
}

test('xlsx roundtrip: opening a second workbook replaces cells, sheets, styles, and undo history', async ({ page }, info) => {
  await open(page);
  await edit(page, 'A1', 'First workbook only');
  await cellText(page, 'A1', 'First workbook only');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+b');
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await save(page, info, 'replacement-first');
  const next = await fixture({ name: 'replacement.xlsx', sheets: [
    { name: 'Replacement', rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>Fresh workbook</t></is></c><c r="B1"><v>100</v></c></row>' },
  ] });
  await replace(page, next);
  await expect(page.getByRole('tab')).toHaveText(['Replacement']);
  await expect(page.getByTestId('xlsx-undo')).toBeDisabled();
  await expect(page.getByTestId('xlsx-redo')).toBeDisabled();
  await inputAt(page, 'A1', 'Fresh workbook');
  await cellText(page, 'A1', 'Fresh workbook');
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'false');
  await inputAt(page, 'C1', '');
  await cellText(page, 'C1', '');
  await expect(page.getByRole('gridcell').filter({ hasText: 'First workbook only' })).toHaveCount(0);
  await edit(page, 'B1', '101');
  await cellText(page, 'B1', '101');
  await history(page, 'B1', '100', '101');
  const saved = await save(page, info, 'replacement-second');
  const model = await inspect(page, saved.parts);
  expect(Object.keys(model)).toEqual(['xl/worksheets/sheet1.xml']);
  expect(Object.keys(model['xl/worksheets/sheet1.xml']).sort()).toEqual(['A1', 'B1']);
  expect(model['xl/worksheets/sheet1.xml']).toMatchObject({ A1: { value: 'Fresh workbook', style: { bold: false, numFmtId: '0' } }, B1: { value: '101' } });
  expect(saved.parts['xl/workbook.xml']).not.toContain('name="Other"');
  await open(page, saved.file);
  await inputAt(page, 'A1', 'Fresh workbook');
  await inputAt(page, 'B1', '101');
  const reopened = await save(page, info, 'replacement-reopened');
  expect(await inspect(page, reopened.parts)).toEqual(model);
});

test('xlsx roundtrip: showcase charts, drawings, relationships, widths, and merges survive an edit', async ({ page }, info) => {
  const file = resolve(root, 'apps/demo/public/showcase.xlsx');
  const source = await xmlParts(await readFile(file));
  await open(page, file);
  await expect(page.getByTestId('xlsx-formula-input')).not.toHaveValue('');
  const original = await page.getByTestId('xlsx-formula-input').inputValue();
  await edit(page, 'A1', 'Preserved showcase');
  await cellText(page, 'A1', 'Preserved showcase');
  await history(page, 'A1', original, 'Preserved showcase');
  const saved = await save(page, info, 'showcase-preserved');
  const untouched = Object.keys(source).filter((name) => /^(xl\/(charts|drawings)\/|xl\/worksheets\/_rels\/)/.test(name));
  expect(untouched).toContain('xl/charts/chart1.xml');
  expect(untouched).toContain('xl/drawings/drawing1.xml');
  for (const name of untouched) expect(saved.parts[name], name).toBe(source[name]);
  expect(saved.parts['xl/workbook.xml']).toBe(source['xl/workbook.xml']
    .replace(' date1904="0"', '').replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'));
  expect(saved.parts['xl/styles.xml']).toBe(source['xl/styles.xml']);
  for (const name of Object.keys(source).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
    for (const element of ['cols', 'mergeCells', 'sheetViews', 'drawing']) {
      expect(elementXml(saved.parts[name], element), `${name}: ${element}`).toBe(elementXml(source[name], element));
    }
  }
  await open(page, saved.file);
  await inputAt(page, 'A1', 'Preserved showcase');
  await cellText(page, 'A1', 'Preserved showcase');
  const reopened = await save(page, info, 'showcase-reopened');
  expect(reopened.parts).toEqual(saved.parts);
});

test('xlsx roundtrip: defined names, freeze panes, widths, and merges retain their original XML', async ({ page }, info) => {
  const input = await fixture({ definedNames: '<definedNames><definedName name="InputAmount">Entries!$B$1</definedName><definedName name="LocalDate" localSheetId="0">Entries!$C$1</definedName></definedNames>', sheets: [
    { name: 'Entries', rows: firstRow,
      beforeData: '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="6" width="15" customWidth="1"/></cols>',
      afterData: '<mergeCells count="1"><mergeCell ref="A3:C3"/></mergeCells>' },
  ] });
  const source = await xmlParts(input.buffer);
  await open(page, input);
  await edit(page, 'B1', '80');
  await cellText(page, 'B1', '80');
  await history(page, 'B1', '12.5', '80');
  const first = await save(page, info, 'metadata-preserved');
  for (const element of ['cols', 'mergeCells', 'sheetViews']) {
    const original = elementXml(source['xl/worksheets/sheet1.xml'], element);
    expect(original).not.toBeNull();
    expect(elementXml(first.parts['xl/worksheets/sheet1.xml'], element)).toBe(original);
  }
  expect(elementXml(first.parts['xl/workbook.xml'], 'definedNames')).toBe(elementXml(source['xl/workbook.xml'], 'definedNames'));
  expect((await inspect(page, first.parts))['xl/worksheets/sheet1.xml'].B1.value).toBe('80');
  await open(page, first.file);
  await inputAt(page, 'B1', '80');
  await cellText(page, 'B1', '80');
  const second = await save(page, info, 'metadata-reopened');
  expect(second.parts).toEqual(first.parts);
});

test('xlsx roundtrip: a 2000-row workbook opens, edits its last row, and saves within 30 seconds', async ({ page }, info) => {
  test.setTimeout(30_000);
  const rows = Array.from({ length: 2000 }, (_, index) => {
    const row = index + 1;
    return `<row r="${row}"><c r="A${row}"><v>${row}</v></c><c r="B${row}" t="inlineStr"><is><t>Row ${row}</t></is></c></row>`;
  }).join('');
  await open(page, await fixture({ name: 'large.xlsx', sheets: [{ name: 'Large', rows }] }));
  await inputAt(page, 'A1', '1');
  await cellText(page, 'B1', 'Row 1');
  const scroll = page.getByTestId('xlsx-scroll');
  await lastRow(page, 2000);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('2000');
  await cellText(page, 'A2000', '2000');
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue('2000');
  await input.press('ControlOrMeta+A');
  await input.pressSequentially('2001');
  await input.press('Enter');
  await expect(input).toHaveCount(0);
  await cellText(page, 'A2000', '2001');
  await page.getByTestId('xlsx-undo').click();
  await cellText(page, 'A2000', '2000');
  await page.getByTestId('xlsx-redo').click();
  await cellText(page, 'A2000', '2001');
  const first = await save(page, info, 'large-edited');
  const model = await inspect(page, first.parts);
  expect(Object.keys(model['xl/worksheets/sheet1.xml'])).toHaveLength(4000);
  expect(model['xl/worksheets/sheet1.xml']).toMatchObject({ A1: { value: '1' }, A1999: { value: '1999' }, A2000: { value: '2001' }, B2000: { value: 'Row 2000' } });
  await open(page, first.file);
  await lastRow(page, 2000);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('2001');
  await cellText(page, 'B2000', 'Row 2000');
  const second = await save(page, info, 'large-reopened');
  expect(stableParts(second.parts)).toEqual(stableParts(first.parts));
});

test('xlsx roundtrip: typing a date in a date cell saves a numeric date rather than text', async ({ page }, info) => {
  await open(page, await fixture({ sheets: [
    { name: 'Dates', rows: '<row r="1"><c r="C1" s="1"><v>45292</v></c></row>' },
  ] }));
  await select(page, 'C1');
  await cellText(page, 'C1', '1/1/24');
  const input = page.getByTestId('xlsx-formula-input');
  await input.fill('1/15/2024');
  await input.press('Enter');
  await select(page, 'C1');
  await cellText(page, 'C1', '1/15/24');
  await page.getByTestId('xlsx-undo').click();
  await cellText(page, 'C1', '1/1/24');
  await page.getByTestId('xlsx-redo').click();
  await cellText(page, 'C1', '1/15/24');
  const first = await save(page, info, 'typed-date');
  expect((await inspect(page, first.parts))['xl/worksheets/sheet1.xml'].C1).toMatchObject({ value: '45306', type: 'n', style: { numFmtId: '14' } });
  await open(page, first.file);
  await select(page, 'C1');
  await cellText(page, 'C1', '1/15/24');
  const second = await save(page, info, 'typed-date-reopened');
  expect(stableParts(second.parts)).toEqual(stableParts(first.parts));
});
