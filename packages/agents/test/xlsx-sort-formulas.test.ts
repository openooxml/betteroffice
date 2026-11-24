import { expect, test } from 'bun:test';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import type { XlsxCellValue } from '@betteroffice/xlsx';
import JSZip from 'jszip';
import { createOfficeMcpServer } from '../src/mcp';
import { DocumentToolError } from '../src/types';
import { attachXlsx, openXlsx, type XlsxAgentWorkbook } from '../src/xlsx';
import type { XlsxWriteEdit } from '../src/xlsx-operations';
import { implicitFontWorkbook } from './format-fixtures';

type Call = (name: string, args: Record<string, unknown>, expectError?: boolean) => Promise<Record<string, unknown>>;
type SortOptions = Omit<Parameters<XlsxAgentWorkbook['sortRange']>[0], 'version'>;

const sales = { sheet: 'Sales', range: 'A1:B4', values: [['Month', 'North'], ['Jan', 100], ['Feb', 110], ['Mar', 120]] };
const doubled = { sheet: 'Sales', range: 'F2:F4', formulas: [['=B2*2'], ['=B3*2'], ['=B4*2']] };
const descending = { sheet: 'Sales', range: 'A1:F4', header: true, keys: [{ column: 'B', order: 'desc' }] } satisfies SortOptions;

async function withWorkbook(run: (call: Call) => Promise<void>, sheetData = '<sheetData/>') {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-sort-formulas-'));
  const client = new Client({ name: 'xlsx-sort-formulas', version: '1' });
  let server: Awaited<ReturnType<typeof createOfficeMcpServer>> | undefined;
  try {
    const zip = await JSZip.loadAsync(await implicitFontWorkbook());
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const sheet = await zip.file(sheetPath)!.async('string');
    zip.file(sheetPath, sheet.replace(/<sheetData>.*?<\/sheetData>/, () => sheetData));
    await writeFile(join(root, 'book.xlsx'), await zip.generateAsync({ type: 'uint8array' }));
    server = await createOfficeMcpServer({ root, formats: ['xlsx'] });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    const call: Call = async (name, args, expectError = false) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError === true, JSON.stringify(output.structuredContent)).toBe(expectError);
      return output.structuredContent!;
    };
    const { document } = await call('office_open', { path: 'book.xlsx' });
    await run((name, args, expectError) => call(name, { document, ...args }, expectError));
  } finally {
    await client.close();
    await server?.close();
    await rm(root, { recursive: true, force: true });
  }
}

async function withUnsupportedHandle(run: (book: XlsxAgentWorkbook, handle: XlsxAgentWorkbook['handle']) => void) {
  const owner = await openXlsx(await implicitFontWorkbook());
  const { handle } = owner;
  let book: XlsxAgentWorkbook | undefined;
  try {
    const written = owner.writeRanges({ version: handle.version(), edits: [sales, doubled] });
    expect(written.applied).toBe(true);
    book = attachXlsx(new Proxy(handle, {}));
    run(book, handle);
  } finally {
    book?.close();
    owner.close();
  }
}

async function readRange(call: Call, range: string, sheet = 'Sales') {
  return await call('xlsx_read_range', { sheet, range }) as ReturnType<XlsxAgentWorkbook['readCells']>;
}

function cellContents({ cells }: ReturnType<XlsxAgentWorkbook['readCells']>) {
  return cells.map(({ a1, formula, value }) => ({ a1, formula, value }));
}

async function writeAndSort(call: Call, edits: XlsxWriteEdit[], options: SortOptions) {
  const { version } = await readRange(call, options.range, options.sheet);
  const written = await call('xlsx_write_range', { version, edits });
  expect(written.applied).toBe(true);
  const sorted = await call('xlsx_sort', { ...options, version: written.version });
  expect(sorted.applied).toBe(true);
}

async function expectSortRefused(call: Call, options: SortOptions, message: string) {
  const before = await readRange(call, options.range, options.sheet);
  const refused = await call('xlsx_sort', { ...options, version: before.version }, true);
  expect(refused).toEqual({ code: 'INVALID_SORT', message });
  const after = await readRange(call, options.range, options.sheet);
  expect(after.version).toBe(before.version);
  expect(cellContents(after)).toEqual(cellContents(before));
}

async function expectValues(call: Call, range: string, expected: (string | number | null)[]) {
  const { cells } = await readRange(call, range);
  expect(cells.map(cell => cell.formula)).toEqual(expected.map(() => null));
  expect(cells.map(cell => cell.value)).toEqual(expected.map((value): XlsxCellValue => {
    if (value === null) return { kind: 'empty' };
    return typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value };
  }));
}

async function expectFormulas(call: Call, range: string, expected: [string, number | '#REF!'][]) {
  const { cells } = await readRange(call, range);
  expect(cells.map(({ formula, value }) => ({ formula, value }))).toEqual(expected.map(([formula, value]): { formula: string; value: XlsxCellValue } => ({
    formula,
    value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'error', value },
  })));
}

test('XLSX sort updates own-row formulas when sorting numbers descending', () => withWorkbook(async call => {
  await writeAndSort(call, [sales, doubled], descending);

  await expectValues(call, 'A1:F1', ['Month', 'North', null, null, null, null]);
  await expectValues(call, 'A2:B4', ['Mar', 120, 'Feb', 110, 'Jan', 100]);
  await expectFormulas(call, 'F2:F4', [['B2*2', 240], ['B3*2', 220], ['B4*2', 200]]);
}));

test('XLSX sort refuses moving formulas on an unsupported attached handle', () => withUnsupportedHandle((book, handle) => {
  const { sheet, range } = descending;
  const before = cellContents(book.readCells({ sheet, range }));
  const version = handle.version();
  let refusal: unknown;
  try { book.sortRange({ ...descending, version }); }
  catch (error) { refusal = error; }
  expect(refusal).toBeInstanceOf(DocumentToolError);
  expect(refusal).toMatchObject({ code: 'INVALID_SORT' });
  expect(handle.version()).toBe(version);
  expect(cellContents(book.readCells({ sheet, range }))).toEqual(before);
}));

test('XLSX sort reorders values on an unsupported attached handle when no formulas move', () => withUnsupportedHandle((book, handle) => {
  const sorted = book.sortRange({ ...descending, range: 'A1:B4', version: handle.version() });
  expect(sorted.applied).toBe(true);
  const { cells } = book.readCells({ sheet: 'Sales', range: 'A1:B4' });
  expect(cells.map(cell => cell.value)).toEqual([
    { kind: 'text', value: 'Month' }, { kind: 'text', value: 'North' },
    { kind: 'text', value: 'Mar' }, { kind: 'number', value: 120 },
    { kind: 'text', value: 'Feb' }, { kind: 'number', value: 110 },
    { kind: 'text', value: 'Jan' }, { kind: 'number', value: 100 },
  ]);
}));

test('XLSX sort refuses a range that cuts through an array formula', () => withWorkbook(async call => {
  const before = await readRange(call, 'B2:F4');
  const refused = await call('xlsx_sort', { ...descending, range: 'B3:F4', header: false, version: before.version }, true);
  expect(refused).toMatchObject({ code: 'locked-target', message: 'F3 is filled by an array formula' });
  const after = await readRange(call, 'B2:F4');
  expect(after.version).toBe(before.version);
  expect(cellContents(after)).toEqual(cellContents(before));
}, '<sheetData><row r="2"><c r="B2"><v>100</v></c><c r="F2"><f t="array" ref="F2:F3">B2:B3*2</f><v>200</v></c></row><row r="3"><c r="B3"><v>110</v></c><c r="F3"><v>220</v></c></row><row r="4"><c r="B4"><v>120</v></c></row></sheetData>'));

test('XLSX sort updates own-row formulas when sorting text ascending', () => withWorkbook(async call => {
  await writeAndSort(call, [sales, doubled], { ...descending, keys: [{ column: 'A', order: 'asc' }] });

  await expectValues(call, 'A2:B4', ['Feb', 110, 'Jan', 100, 'Mar', 120]);
  await expectFormulas(call, 'F2:F4', [['B2*2', 220], ['B3*2', 200], ['B4*2', 240]]);
}));

test('XLSX sort shifts relative rows and preserves anchors in formulas and ranges', () => withWorkbook(async call => {
  await writeAndSort(call, [sales, { sheet: 'Sales', range: 'C2:F4', formulas: [
    ['=$B2*2', '=B$2*2', '=$B$2*2', '=SUM(B$2:B2)'],
    ['=$B3*2', '=B$2*2', '=$B$2*2', '=SUM(B$2:B3)'],
    ['=$B4*2', '=B$2*2', '=$B$2*2', '=SUM(B$2:B4)'],
  ] }], descending);

  await expectFormulas(call, 'C2:F4', [
    ['$B2*2', 240], ['B$2*2', 240], ['$B$2*2', 240], ['SUM(B$2:B2)', 120],
    ['$B3*2', 220], ['B$2*2', 240], ['$B$2*2', 240], ['SUM(B$2:B3)', 230],
    ['$B4*2', 200], ['B$2*2', 240], ['$B$2*2', 240], ['SUM(B$2:B4)', 330],
  ]);
}));

test('XLSX sort preserves the offset of a previous-row reference inside the range', () => withWorkbook(async call => {
  await writeAndSort(call, [
    { ...sales, values: [['Month', 'North'], ['Jan', 100], ['Feb', 90], ['Mar', 120]] },
    { sheet: 'Sales', range: 'G3', formulas: [['=B2']] },
  ], { ...descending, range: 'A1:G4' });

  await expectValues(call, 'A2:B4', ['Mar', 120, 'Jan', 100, 'Feb', 90]);
  await expectFormulas(call, 'G4', [['B3', 100]]);
}));

test('XLSX sort matches paste translation for cross-sheet references', () => withWorkbook(async call => {
  const { version } = await readRange(call, 'A1');
  await call('xlsx_edit_sheets', { version, edits: [{ action: 'add', name: 'Rates' }] });
  await writeAndSort(call, [
    sales,
    { sheet: 'Rates', range: 'B2:B4', values: [[2], [3], [5]] },
    { sheet: 'Sales', range: 'F2:G4', formulas: [
      ['=Rates!B2*B2', '=Rates!$B$2'],
      ['=Rates!B3*B3', '=Rates!$B$2'],
      ['=Rates!B4*B4', '=Rates!$B$2'],
    ] },
  ], { ...descending, range: 'A1:G4' });

  await expectFormulas(call, 'F2:G4', [
    ['Rates!B2*B2', 240], ['Rates!$B$2', 2],
    ['Rates!B3*B3', 330], ['Rates!$B$2', 2],
    ['Rates!B4*B4', 500], ['Rates!$B$2', 2],
  ]);
}));

test('XLSX sort keeps equal keys stable and translates formulas for moved rows', () => withWorkbook(async call => {
  await writeAndSort(call, [
    { sheet: 'Sales', range: 'A1:B5', values: [['Name', 'Key'], ['First', 10], ['Second', 10], ['Higher', 20], ['Third', 10]] },
    { sheet: 'Sales', range: 'C2:C5', formulas: [['=B2 * 2'], ['=B3*3'], ['=B4*4'], ['=B5 * 5']] },
  ], { ...descending, range: 'A1:C5' });

  await expectValues(call, 'A2:A5', ['Higher', 'First', 'Second', 'Third']);
  await expectFormulas(call, 'C2:C5', [['B2*4', 80], ['B3 * 2', 20], ['B4*3', 30], ['B5 * 5', 50]]);
}));

test('XLSX sort translates formulas in a headerless block starting at row 10', () => withWorkbook(async call => {
  await writeAndSort(call, [
    { sheet: 'Sales', range: 'A10:B13', values: [['a', 30], ['b', 10], ['c', 40], ['d', 20]] },
    { sheet: 'Sales', range: 'C10:C13', formulas: [['=B10*2'], ['=B11*2'], ['=B12*2'], ['=B13*2']] },
  ], { sheet: 'Sales', range: 'A10:C13', keys: [{ column: 'B', order: 'asc' }] });

  await expectValues(call, 'A10:A13', ['b', 'd', 'a', 'c']);
  await expectFormulas(call, 'C10:C13', [['B10*2', 20], ['B11*2', 40], ['B12*2', 60], ['B13*2', 80]]);
}));

test('XLSX sort emits #REF! when a relative reference moves above row 1', () => withWorkbook(async call => {
  await writeAndSort(call, [
    sales,
    { sheet: 'Sales', range: 'F2', formulas: [['=B1']] },
    { sheet: 'Sales', range: 'F4', formulas: [['=B2']] },
  ], descending);

  await expectFormulas(call, 'F2', [['#REF!', '#REF!']]);
  await expectFormulas(call, 'F4', [['B3', 110]]);
}));

test.each(['[1]Sheet1!A1', "'[1]Sheet 1'!A1", '[Book.xlsx]Sheet1!A1'])('XLSX sort refuses to translate %s', formula => withWorkbook(async call => {
  await expectFormulas(call, 'C2', [[formula, 42]]);
  await expectSortRefused(call, { ...descending, range: 'B2:C3', header: false }, 'Cannot move the formula at C2. It refers to an external workbook.');
}, `<sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2"><f>${formula}</f><v>42</v></c></row><row r="3"><c r="B3"><v>2</v></c></row></sheetData>`));

test('XLSX sort ignores external reference syntax inside string literals and structured references', () => withWorkbook(async call => {
  await writeAndSort(call, [sales, { sheet: 'Sales', range: 'F2:G2', formulas: [['=B2+N("[1]Sheet1!A1")', '=B2+IF(FALSE,Table1[Col],0)']] }], { ...descending, range: 'A1:G4' });
  const { cells } = await readRange(call, 'F4:G4');
  expect(cells.map(cell => cell.formula)).toEqual(['B4+N("[1]Sheet1!A1")', 'B4+IF(FALSE,Table1[Col],0)']);
}));

test('XLSX sort refuses a translation that grows beyond 8192 characters', () => withWorkbook(async call => {
  const { version } = await readRange(call, 'B9:C10');
  const written = await call('xlsx_write_range', { version, edits: [
    { sheet: 'Sales', range: 'B9:B10', values: [[1], [2]] },
    { sheet: 'Sales', range: 'C9', formulas: [[`=B9${' '.repeat(8188)}+0`]] },
  ] });
  expect(written.applied).toBe(true);
  await expectSortRefused(call, { ...descending, range: 'B9:C10', header: false }, 'Cannot move the formula at C9. The translated formula exceeds 8192 characters.');
}));

test('XLSX sort reports protected formula cells by their recalculated values', () => withWorkbook(async call => {
  const { version } = await readRange(call, 'B2:D3');
  const sorted = await call('xlsx_sort', { ...descending, range: 'B2:D3', header: false, version, protectedSheets: 'allow' });
  expect(sorted).toMatchObject({ applied: true, protectedCells: ['Sales!B2', 'Sales!D2', 'Sales!B3', 'Sales!D3'] });
  await expectFormulas(call, 'C2:D3', [['E2', 10], ['B$2', 2], ['E3', 20], ['B$2', 2]]);
}, '<sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2"><f>E2</f><v>10</v></c><c r="D2"><f>B$2</f><v>1</v></c><c r="E2"><v>10</v></c></row><row r="3"><c r="B3"><v>2</v></c><c r="C3"><f>E3</f><v>20</v></c><c r="D3"><f>B$2</f><v>1</v></c><c r="E3"><v>20</v></c></row></sheetData><sheetProtection sheet="1"/>'));
