import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { selectCell, workbookXml } from './xlsx-async-clipboard-helpers';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const chartPart = 'xl/charts/chart1.xml';
const chart = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>'
  + '<c:cat><c:multiLvlStrRef><c:f>Sheet!$A$1:$B$4</c:f><c:multiLvlStrCache><c:ptCount val="4"/>'
  + '<c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q1</c:v></c:pt><c:pt idx="3"><c:v>Q2</c:v></c:pt></c:lvl>'
  + '<c:lvl><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>'
  + '<c:val><c:numRef><c:f>Sheet!$C$1:$C$4</c:f><c:numCache><c:ptCount val="4"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val>'
  + '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

/** A sheet whose chart plots A1:B4 through a multi-level category cache, which a row insert inside it cannot rebuild. */
async function multiLevelChartWorkbook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
  const row = (r: number, region: string, quarter: string) =>
    `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>${region}</t></is></c><c r="B${r}" t="inlineStr"><is><t>${quarter}</t></is></c><c r="C${r}"><v>${r}</v></c></row>`;
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${ns}" xmlns:r="${rel}"><sheetData>${row(1, 'North', 'Q1')}${row(2, 'North', 'Q2')}${row(3, 'South', 'Q1')}${row(4, 'South', 'Q2')}<row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="C6"><v>10</v></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>`);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
  zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${rel}"><xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="${rel}/chart" Target="../charts/chart1.xml"/></Relationships>`);
  zip.file(chartPart, chart);
  return { name: 'chart-refusal.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function saved(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.xlsx');
  await (await downloading).saveAs(file);
  const bytes = await readFile(file);
  return { cells: (await workbookXml(page, bytes))['xl/worksheets/sheet1.xml'].cells, chart: await (await JSZip.loadAsync(bytes)).file(chartPart)!.async('string') };
}

async function insertRowsAbove(page: Page) {
  await page.getByTestId('xlsx-scroll').press('Shift+F10');
  const menu = page.getByRole('menu', { name: 'Cell menu', exact: true });
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitem', { name: 'Insert', exact: true }).hover();
  await page.getByRole('menu', { name: 'Insert', exact: true }).getByRole('menuitem', { name: 'Insert rows above', exact: true }).click();
  await expect(menu).toBeHidden();
}

test('a row insert a chart cannot follow is refused with a banner naming the chart, and nothing changes', async ({ page }, info) => {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(await multiLevelChartWorkbook());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await selectCell(page, 1, 0);
  await insertRowsAbove(page);
  await expect(page.getByTestId('xlsx-input-refusal'))
    .toHaveText("This edit can't be applied: a chart uses cells it would move. The workbook is unchanged.");
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('North');
  const after = await saved(page, info, 'refused');
  expect([after.cells.A2?.value, after.cells.A4?.value, after.cells.A5?.value ?? '']).toEqual(['North', 'South', '']);
  expect(after.chart).toBe(chart);

  await selectCell(page, 5, 0);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Total');
  await insertRowsAbove(page);
  await expect(page.getByTestId('xlsx-input-refusal')).toHaveCount(0);
  const below = await saved(page, info, 'below');
  expect([below.cells.A4?.value, below.cells.A6?.value ?? '', below.cells.A7?.value, below.cells.C7?.value])
    .toEqual(['South', '', 'Total', '10']);
  expect(below.chart).toContain('<c:f>Sheet!$A$1:$B$4</c:f>');
});
