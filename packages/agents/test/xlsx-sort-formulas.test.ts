import { expect, test } from 'bun:test';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import JSZip from 'jszip';
import { createOfficeMcpServer } from '../src/mcp';
import type { XlsxAgentWorkbook } from '../src/xlsx';
import type { XlsxWriteEdit } from '../src/xlsx-operations';
import { implicitFontWorkbook } from './format-fixtures';

type Call = (name: string, args: Record<string, unknown>) => Promise<Record<string, unknown>>;
type SortOptions = Omit<Parameters<XlsxAgentWorkbook['sortRange']>[0], 'version'>;

const sales = { sheet: 'Sales', range: 'A1:B4', values: [['Month', 'North'], ['Jan', 100], ['Feb', 110], ['Mar', 120]] };
const doubled = { sheet: 'Sales', range: 'F2:F4', formulas: [['=B2*2'], ['=B3*2'], ['=B4*2']] };
const descending = { sheet: 'Sales', range: 'A1:F4', header: true, keys: [{ column: 'B', order: 'desc' }] } satisfies SortOptions;

async function withWorkbook(run: (call: Call) => Promise<void>) {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-sort-formulas-'));
  const client = new Client({ name: 'xlsx-sort-formulas', version: '1' });
  let server: Awaited<ReturnType<typeof createOfficeMcpServer>> | undefined;
  try {
    const zip = await JSZip.loadAsync(await implicitFontWorkbook());
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const sheet = await zip.file(sheetPath)!.async('string');
    zip.file(sheetPath, sheet.replace(/<sheetData>.*?<\/sheetData>/, '<sheetData/>'));
    await writeFile(join(root, 'book.xlsx'), await zip.generateAsync({ type: 'uint8array' }));
    server = await createOfficeMcpServer({ root, formats: ['xlsx'] });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    const call: Call = async (name, args) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent!;
    };
    const { document } = await call('office_open', { path: 'book.xlsx' });
    await run((name, args) => call(name, { document, ...args }));
  } finally {
    await client.close();
    await server?.close();
    await rm(root, { recursive: true, force: true });
  }
}

async function readRange(call: Call, range: string, sheet = 'Sales') {
  return await call('xlsx_read_range', { sheet, range }) as ReturnType<XlsxAgentWorkbook['readCells']>;
}

async function writeAndSort(call: Call, edits: XlsxWriteEdit[], options: SortOptions) {
  const { version } = await readRange(call, options.range, options.sheet);
  const written = await call('xlsx_write_range', { version, edits });
  expect(written.applied).toBe(true);
  const sorted = await call('xlsx_sort', { ...options, version: written.version });
  expect(sorted.applied).toBe(true);
}

async function expectValues(call: Call, range: string, expected: (string | number | null)[]) {
  const { cells } = await readRange(call, range);
  expect(cells.map(cell => cell.formula)).toEqual(expected.map(() => null));
  expect(cells.map(cell => cell.value)).toEqual(expected.map(value => {
    if (value === null) return { kind: 'empty' };
    return { kind: typeof value === 'number' ? 'number' : 'text', value };
  }));
}

async function expectFormulas(call: Call, range: string, expected: [string, number | '#REF!'][]) {
  const { cells } = await readRange(call, range);
  expect(cells.map(({ formula, value }) => ({ formula, value }))).toEqual(expected.map(([formula, value]) => ({
    formula,
    value: { kind: typeof value === 'number' ? 'number' : 'error', value },
  })));
}

test('XLSX sort updates own-row formulas when sorting numbers descending', () => withWorkbook(async call => {
  await writeAndSort(call, [sales, doubled], descending);

  await expectValues(call, 'A1:F1', ['Month', 'North', null, null, null, null]);
  await expectValues(call, 'A2:B4', ['Mar', 120, 'Feb', 110, 'Jan', 100]);
  await expectFormulas(call, 'F2:F4', [['B2*2', 240], ['B3*2', 220], ['B4*2', 200]]);
}));

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
