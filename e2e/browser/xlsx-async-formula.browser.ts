import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

type Cell = {
  ref: string;
  formula: string | null;
  value: string;
  sheet?: number;
  type?: string;
};

const original: Cell[] = [
  { ref: 'A1', formula: null, value: '7' },
  { ref: 'B1', formula: 'A1*2', value: '14' },
  { ref: 'C1', formula: 'B1+1', value: '15' },
];

async function workbook() {
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
  zip.file('xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Source Data" sheetId="2" r:id="rId2"/></sheets>
    <calcPr calcMode="auto"/>
  </workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
    <borders count="1"><border/></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  </styleSheet>`);
  const sheet = (rows: string) => `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:F6"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>
    <sheetFormatPr defaultRowHeight="24"/>
    <cols><col min="1" max="6" width="20" customWidth="1"/></cols>
    <sheetData>${rows}</sheetData>
  </worksheet>`;
  zip.file('xl/worksheets/sheet1.xml', sheet(`
    <row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>A1*2</f><v>14</v></c><c r="C1"><f>B1+1</f><v>15</v></c></row>
    <row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>3</v></c></row>
    <row r="3"><c r="A3"><v>11</v></c></row>
    <row r="4"><c r="A4"><v>13</v></c></row>`));
  zip.file('xl/worksheets/sheet2.xml', sheet('<row r="1"><c r="A1"><v>20</v></c></row>'));
  return { name: 'formulas.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page, file?: string) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(file ?? await workbook());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
}

async function select(page: Page, ref: string) {
  const match = /^([A-F])([1-6])$/.exec(ref);
  if (!match) throw new Error('Unsupported test cell: ' + ref);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  for (let row = 1; row < Number(match[2]); row++) await scroll.press('ArrowDown');
  for (let col = 65; col < match[1].charCodeAt(0); col++) await scroll.press('ArrowRight');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(ref);
}

function paintedCell(page: Page, ref: string) {
  return page.getByRole('gridcell').filter({ hasText: new RegExp('^' + ref + ', ') });
}

async function painted(page: Page, ref: string, value: string) {
  const escaped = (value || 'empty').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  await expect(paintedCell(page, ref)).toHaveText(new RegExp('^' + ref + ', ' + escaped + '(?:, selected)?$'));
}

async function checkCells(page: Page, cells: Cell[]) {
  for (const cell of cells) {
    const tab = page.getByRole('tab', { name: cell.sheet === 2 ? 'Source Data' : 'Sheet1', exact: true });
    if (await tab.getAttribute('aria-selected') !== 'true') await tab.click();
    await expect(tab).toHaveAttribute('aria-selected', 'true');
    await select(page, cell.ref);
    await painted(page, cell.ref, cell.value);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(cell.formula === null ? cell.value : '=' + cell.formula);
  }
}

async function edit(page: Page, ref: string, value: string, commit = 'Enter', previous = '') {
  await select(page, ref);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(previous);
  await page.getByTestId('xlsx-scroll').press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue(previous);
  await input.press('ControlOrMeta+A');
  await input.pressSequentially(value);
  await expect(input).toHaveValue(value);
  await input.press(commit);
  await expect(input).toHaveCount(0);
}

async function undoRedo(page: Page, before: Cell[], after: Cell[]) {
  await expect(page.getByTestId('xlsx-undo')).toBeEnabled();
  await page.getByTestId('xlsx-undo').click();
  await checkCells(page, before);
  await expect(page.getByTestId('xlsx-redo')).toBeEnabled();
  await page.getByTestId('xlsx-redo').click();
  await checkCells(page, after);
}

async function save(page: Page, info: TestInfo, label: string, cells: Cell[]) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(label + '.xlsx');
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  for (const sheet of new Set(cells.map((cell) => cell.sheet ?? 1))) {
    const part = zip.file(`xl/worksheets/sheet${sheet}.xml`);
    expect(part).not.toBeNull();
    const values = await page.evaluate((xml) => {
      const document = new DOMParser().parseFromString(xml, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error('Invalid saved worksheet XML');
      return Object.fromEntries(Array.from(document.getElementsByTagNameNS('*', 'c')).map((cell) => [cell.getAttribute('r'), {
        formula: cell.getElementsByTagNameNS('*', 'f')[0]?.textContent ?? null,
        value: cell.getElementsByTagNameNS('*', 'v')[0]?.textContent ?? '',
        type: cell.getAttribute('t'),
      }]));
    }, await part!.async('string'));
    for (const cell of cells.filter((cell) => (cell.sheet ?? 1) === sheet)) {
      const saved = values[cell.ref] ?? { formula: null, value: '', type: null };
      expect(saved, `Saved sheet ${sheet} cell ${cell.ref}`).toMatchObject({
        formula: cell.formula,
        value: cell.value,
        ...(cell.type ? { type: cell.type } : {}),
      });
    }
  }
  return file;
}

async function roundTrip(page: Page, info: TestInfo, cells: Cell[]) {
  const file = await save(page, info, 'edited', cells);
  await open(page, file);
  await checkCells(page, cells);
  await save(page, info, 'reopened', cells);
}

test('xlsx: typing SUM paints its value and keeps the formula through undo, redo and reopen', async ({ page }, info) => {
  await open(page);
  await select(page, 'D1');
  await page.keyboard.type('=SUM(A1:A3)');
  await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue('=SUM(A1:A3)');
  await page.getByTestId('xlsx-cell-editor').press('Enter');
  const result: Cell = { ref: 'D1', formula: 'SUM(A1:A3)', value: '23' };
  await checkCells(page, [result]);
  await undoRedo(page, [{ ref: 'D1', formula: null, value: '' }], [result]);
  await roundTrip(page, info, [result]);
});

test('xlsx: changing a precedent recalculates a chain while formula bars retain expressions', async ({ page }, info) => {
  await open(page);
  await checkCells(page, original);
  await edit(page, 'A1', '9', 'Enter', '7');
  const changed: Cell[] = [
    { ref: 'A1', formula: null, value: '9' },
    { ref: 'B1', formula: 'A1*2', value: '18' },
    { ref: 'C1', formula: 'B1+1', value: '19' },
  ];
  await checkCells(page, changed);
  await undoRedo(page, original, changed);
  await roundTrip(page, info, changed);
});

test('xlsx: editing a formula in the formula bar commits with Enter and recalculates dependents', async ({ page }, info) => {
  await open(page);
  await select(page, 'B1');
  const input = page.getByTestId('xlsx-formula-input');
  await expect(input).toHaveValue('=A1*2');
  await input.click();
  await input.press('ControlOrMeta+A');
  await input.pressSequentially('=A1*3');
  await expect(input).toHaveValue('=A1*3');
  await input.press('Enter');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B2');
  const changed: Cell[] = [
    { ref: 'B1', formula: 'A1*3', value: '21' },
    { ref: 'C1', formula: 'B1+1', value: '22' },
  ];
  await checkCells(page, changed);
  await undoRedo(page, original.slice(1), changed);
  await roundTrip(page, info, changed);
});

test('xlsx: Escape in the formula bar restores the formula without a history entry', async ({ page }, info) => {
  await open(page);
  await select(page, 'B1');
  const input = page.getByTestId('xlsx-formula-input');
  await expect(input).toHaveValue('=A1*2');
  await input.click();
  await input.press('ControlOrMeta+A');
  await input.pressSequentially('=A1*99');
  await expect(input).toHaveValue('=A1*99');
  await input.press('Escape');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  await expect(page.getByTestId('xlsx-scroll')).toBeFocused();
  await checkCells(page, original);
  await expect(page.getByTestId('xlsx-undo')).toBeDisabled();
  await roundTrip(page, info, original);
});

test('xlsx: quoted cross-sheet references recalculate after editing the source sheet', async ({ page }, info) => {
  await open(page);
  const formula = "'Source Data'!A1+Sheet1!A1";
  await edit(page, 'D1', '=' + formula);
  const initial: Cell = { ref: 'D1', formula, value: '27' };
  await checkCells(page, [initial]);
  await undoRedo(page, [{ ref: 'D1', formula: null, value: '' }], [initial]);
  await page.getByRole('tab', { name: 'Source Data', exact: true }).click();
  await edit(page, 'A1', '30', 'Enter', '20');
  const changed: Cell[] = [
    { ref: 'A1', sheet: 2, formula: null, value: '30' },
    { ref: 'D1', formula, value: '37' },
  ];
  await checkCells(page, changed);
  await undoRedo(page, [{ ref: 'A1', sheet: 2, formula: null, value: '20' }, initial], changed);
  await roundTrip(page, info, changed);
});

test('xlsx: formula errors paint and persist as DIV/0, REF and NAME error cells', async ({ page }, info) => {
  await open(page);
  const errors: Cell[] = [
    { ref: 'D1', formula: '1/0', value: '#DIV/0!', type: 'e' },
    { ref: 'D2', formula: '#REF!+1', value: '#REF!', type: 'e' },
    { ref: 'D3', formula: 'UNKNOWN_FUNCTION(1)', value: '#NAME?', type: 'e' },
  ];
  for (const cell of errors) {
    await edit(page, cell.ref, '=' + cell.formula);
    await checkCells(page, [cell]);
    await undoRedo(page, [{ ref: cell.ref, formula: null, value: '' }], [cell]);
  }
  await roundTrip(page, info, errors);
});

for (const [key, destination] of [
  ['Enter', 'D4'],
  ['Shift+Enter', 'D2'],
  ['Tab', 'E3'],
  ['Shift+Tab', 'C3'],
] as const) {
  test(`xlsx: ${key} commits a formula and moves to ${destination}`, async ({ page }, info) => {
    await open(page);
    await edit(page, 'D3', '=SUM(A1:A2)', key);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue(destination);
    await expect(paintedCell(page, destination)).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('xlsx-scroll')).toBeFocused();
    const result: Cell = { ref: 'D3', formula: 'SUM(A1:A2)', value: '12' };
    await checkCells(page, [result]);
    await undoRedo(page, [{ ref: 'D3', formula: null, value: '' }], [result]);
    await roundTrip(page, info, [result]);
  });
}

test('xlsx: typing replaces a filled cell and Escape cancels a later replacement', async ({ page }, info) => {
  await open(page);
  await select(page, 'B1');
  await page.keyboard.type('=A1*5');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue('=A1*5');
  await input.press('Enter');
  const changed: Cell[] = [
    { ref: 'B1', formula: 'A1*5', value: '35' },
    { ref: 'C1', formula: 'B1+1', value: '36' },
  ];
  await checkCells(page, changed);
  await undoRedo(page, original.slice(1), changed);
  await select(page, 'B1');
  await page.keyboard.type('=A1*99');
  await expect(input).toHaveValue('=A1*99');
  await input.press('Escape');
  await expect(input).toHaveCount(0);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  await checkCells(page, changed);
  await undoRedo(page, original.slice(1), changed);
  await roundTrip(page, info, changed);
});

test('xlsx: double-click edits a formula in place with arrow-key caret movement', async ({ page }, info) => {
  await open(page);
  await select(page, 'B1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('=A1*2');
  const selection = await page.getByTestId('xlsx-selection').boundingBox();
  expect(selection).not.toBeNull();
  await page.mouse.dblclick(selection!.x + selection!.width / 2, selection!.y + selection!.height / 2);
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue('=A1*2');
  await expect(input).toBeFocused();
  await input.press('End');
  await input.press('ArrowLeft');
  expect(await input.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(4);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  await input.press('Shift+End');
  await input.pressSequentially('4');
  await expect(input).toHaveValue('=A1*4');
  await input.press('Enter');
  const changed: Cell[] = [
    { ref: 'B1', formula: 'A1*4', value: '28' },
    { ref: 'C1', formula: 'B1+1', value: '29' },
  ];
  await checkCells(page, changed);
  await undoRedo(page, original.slice(1), changed);
  await roundTrip(page, info, changed);
});

test('xlsx: Delete and Backspace clear a selected range and undo restores its formulas', async ({ page }, info) => {
  await open(page);
  const before: Cell[] = [
    ...original,
    { ref: 'A2', formula: null, value: '5' },
    { ref: 'B2', formula: null, value: '3' },
  ];
  const after: Cell[] = [
    { ref: 'A1', formula: null, value: '' },
    { ref: 'B1', formula: null, value: '' },
    { ref: 'A2', formula: null, value: '' },
    { ref: 'B2', formula: null, value: '' },
    { ref: 'C1', formula: 'B1+1', value: '1' },
  ];
  for (const key of ['Delete', 'Backspace']) {
    await select(page, 'A1');
    const scroll = page.getByTestId('xlsx-scroll');
    await scroll.press('Shift+ArrowRight');
    await scroll.press('Shift+ArrowDown');
    for (const ref of ['A1', 'B1', 'A2', 'B2']) await expect(paintedCell(page, ref)).toHaveAttribute('aria-selected', 'true');
    await expect(paintedCell(page, 'C1')).toHaveAttribute('aria-selected', 'false');
    await scroll.press(key);
    await checkCells(page, after);
    await undoRedo(page, before, after);
    await save(page, info, key.toLowerCase(), after);
    if (key === 'Delete') {
      await page.getByTestId('xlsx-undo').click();
      await checkCells(page, before);
    }
  }
  await roundTrip(page, info, after);
});
