import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${relationships}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Merged" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${relationships}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><dimension ref="A1:D8"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1:C2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><sheetData><row r="1"><c r="D1"><v>1</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:C2"/><mergeCell ref="A4:B5"/></mergeCells></worksheet>`);
  return { name: 'merged.xlsx', mimeType, buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(await fixture());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  return scroll;
}

async function saved(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download', { timeout: 15_000 });
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('merged-saved.xlsx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const strings = await zip.file('xl/sharedStrings.xml')?.async('string') ?? '';
  const shared = [...strings.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1]!.matchAll(/<t[^>]*>([^<]*)<\/t>/g)]
    .map((t) => t[1]).join(''));
  const cells: Record<string, string> = {};
  for (const m of sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const body = m[3] ?? '';
    const value = /<v>([^<]*)<\/v>/.exec(body)?.[1];
    const inline = /<t[^>]*>([^<]*)<\/t>/.exec(body)?.[1];
    const text = /t="s"/.test(m[2]!) && value !== undefined ? shared[Number(value)] : inline ?? value;
    if (text !== undefined) cells[m[1]!] = text;
  }
  return { cells, merges: [...sheet.matchAll(/<mergeCell ref="([^"]+)"/g)].map((m) => m[1]) };
}

test('xlsx merge: type and Enter five times down through merged cells stores every value', async ({ page }, info) => {
  const scroll = await open(page);
  const nameBox = page.getByTestId('xlsx-name-box');
  const path = ['A1'];
  for (let i = 1; i <= 5; i++) {
    await page.keyboard.type(`v${i}`);
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
    path.push(await nameBox.inputValue());
  }
  expect.soft(path).toEqual(['A1', 'A3', 'A4', 'A6', 'A7', 'A8']);
  await expect.soft(page.getByTestId('xlsx-input-refusal')).toHaveCount(0);
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('v1');
  const { cells, merges } = await saved(page, info);
  expect(cells).toEqual({ A1: 'v1', A3: 'v2', A4: 'v3', A6: 'v4', A7: 'v5', D1: '1' });
  expect(merges).toEqual(['A1:C2', 'A4:B5']);
});

test('xlsx merge: arrows, Tab and Shift+Enter skip the cells a merge covers', async ({ page }) => {
  const scroll = await open(page);
  const nameBox = page.getByTestId('xlsx-name-box');
  const steps: [string, string][] = [
    ['ArrowRight', 'D1'], ['ArrowDown', 'D2'], ['ArrowLeft', 'A1'], ['Tab', 'D1'], ['Shift+Tab', 'A1'],
    ['ArrowDown', 'A3'], ['ArrowUp', 'A1'], ['Enter', 'A3'], ['Enter', 'A4'], ['ArrowRight', 'C4'],
    ['ArrowDown', 'C5'], ['ArrowLeft', 'A4'], ['ArrowDown', 'A6'], ['Shift+Enter', 'A4'], ['Shift+Enter', 'A3'],
  ];
  for (const [key, address] of steps) {
    await scroll.press(key);
    await expect(nameBox, `${key} lands on ${address}`).toHaveValue(address);
  }
});

test('xlsx merge: typing into a merge after reaching a covered cell edits its anchor', async ({ page }, info) => {
  const scroll = await open(page);
  const nameBox = page.getByTestId('xlsx-name-box');
  await scroll.press('ArrowRight');
  await scroll.press('ArrowDown');
  await scroll.press('ArrowLeft');
  await expect(nameBox).toHaveValue('A1');
  await page.keyboard.type('anchored');
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue('A3');
  const { cells } = await saved(page, info);
  expect(cells).toEqual({ A1: 'anchored', D1: '1' });
});
