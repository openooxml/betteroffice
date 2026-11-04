import { afterAll, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook, translateFormulasOn } from '../index';
import type { CellPoint, WorkbookHandle } from '../index';
import type { WorkbookSession } from '../session/client';
import { createTestWorkbookSession } from '../session/testHelpers';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

async function book(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

const FORMULAS = [
  'A1+$B$1+$A1+A$1', 'SUM(A1:B2)', "'Other Sheet'!A1+Sheet1!$B2", 'Table1[A1]+Rate_A1+A1', 'A1+$A$1+B$1+$A2',
  'XFD1048576+$XFD$1048576', 'SUM(A:A)', 'SUM($A:B)', 'SUM(1:3)', 'SUM($2:4)', 'SUM(Sheet1:Sheet2!B2:C4)',
  '"A1"&A1', 'IF(A1>0,"x""A1",B2)', 'a1+b$2', 'A1 : B2', 'SUM( A1 , B2 )', 'LOG10(100)+A1%', '1E+3+A1',
  '#REF!+A1', 'INDEX(A1:A3,2)', "'It''s'!C3+[1]Sheet1!A1", 'SUM(A1:B2 C1:C3)', 'Ünïcode_名+A1', '{1,2;3,4}',
  '-A1^2', 'A1:A1048576', 'B$1048576+$A1', 'AAA1+ZZ10', 'SUM(A1:INDEX(B:B,3))', 'A1 +　B2',
];

const MOVES: [CellPoint, CellPoint][] = [
  [{ row: 0, col: 2 }, { row: 2, col: 2 }],
  [{ row: 5, col: 5 }, { row: 0, col: 0 }],
  [{ row: 0, col: 0 }, { row: 3, col: 1 }],
  [{ row: 1, col: 1 }, { row: 1, col: 1 }],
  [{ row: 0, col: 0 }, { row: 1_048_570, col: 16_380 }],
];

let handle: WorkbookHandle;
let session: WorkbookSession;

beforeAll(async () => {
  await initWasm(new Uint8Array(readFileSync(WASM)));
  handle = openWorkbook(await book());
  session = await createTestWorkbookSession(await book());
}, 30_000);

afterAll(async () => {
  handle.dispose();
  await session.dispose();
});

describe('translateFormulasOn', () => {
  it('moves references by the copy offset', () => {
    const from = { row: 0, col: 2 };
    const to = { row: 2, col: 2 };
    expect(translateFormulasOn(handle, [
      { formula: 'A1+$B$1+$A1+A$1', from, to },
      { formula: 'SUM(A1:B2)', from, to },
      { formula: "'Other Sheet'!A1+Sheet1!$B2", from, to },
      { formula: 'Table1[A1]+Rate_A1+A1', from, to },
      { formula: 'A1+$A$1+B$1+$A2', from: { row: 1, col: 1 }, to: { row: 0, col: 0 } },
      { formula: 'XFD1048576', from, to },
    ])).toEqual([
      { ok: true, formula: 'A3+$B$1+$A3+A$1' },
      { ok: true, formula: 'SUM(A3:B4)' },
      { ok: true, formula: "'Other Sheet'!A3+Sheet1!$B4" },
      { ok: true, formula: 'Table1[A1]+Rate_A1+A3' },
      { ok: true, formula: '#REF!+$A$1+A$1+$A1' },
      { ok: true, formula: '#REF!' },
    ]);
  });

  it('reports per item what a paste would refuse', () => {
    expect(translateFormulasOn(handle, [
      { formula: 'A1', from: { row: 1_048_576, col: 0 }, to: { row: 0, col: 0 } },
      { formula: `A1+"${'x'.repeat(32_768)}"`, from: { row: 0, col: 0 }, to: { row: 1, col: 0 } },
      { formula: 'A1', from: { row: 0, col: 0 }, to: { row: 1, col: 0 } },
    ])).toEqual([
      { ok: false, error: 'bad edit args: copy origin is outside the grid' },
      { ok: false, error: 'formula exceeds the length limit' },
      { ok: true, formula: 'A2' },
    ]);
  });

  it('gives the formula a paste writes, through the handle and in one session call', async () => {
    const items = FORMULAS.flatMap((formula) => MOVES.map(([from, to]) => ({ formula, from, to })));
    const moved = translateFormulasOn(handle, items);
    expect(await session.call.translateFormulas(items)).toEqual(moved);
    items.forEach((item, at) => {
      const result = moved[at];
      let written: string;
      try {
        handle.editCells(0, [{ ...item.to, input: `=${item.formula}`, copyFrom: item.from }]);
        written = handle.cell(0, item.to.row, item.to.col).input;
      } catch (error) {
        written = `error: ${(error as Error).message}`;
      }
      let typed: string;
      try {
        if (!result.ok) throw new Error(result.error);
        handle.editCells(0, [{ ...item.to, input: `=${result.formula}` }]);
        typed = handle.cell(0, item.to.row, item.to.col).input;
      } catch (error) {
        typed = `error: ${(error as Error).message}`;
      }
      expect({ ...item, typed }).toEqual({ ...item, typed: written });
    });
  });
});
