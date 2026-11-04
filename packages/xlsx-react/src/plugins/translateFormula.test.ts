import { beforeAll, describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { initWasm, openWorkbook, type CellAddr, type WorkbookHandle } from '@betteroffice/xlsx';
import { translateFormula } from './translateFormula';

const WASM = resolve(import.meta.dir, '../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');

beforeAll(async () => {
  await initWasm(new Uint8Array(readFileSync(WASM)));
}, 30_000);

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

const MOVES: [CellAddr, CellAddr][] = [
  [{ row: 0, col: 2 }, { row: 2, col: 2 }],
  [{ row: 5, col: 5 }, { row: 0, col: 0 }],
  [{ row: 0, col: 0 }, { row: 3, col: 1 }],
  [{ row: 1, col: 1 }, { row: 1, col: 1 }],
  [{ row: 0, col: 0 }, { row: 1_048_570, col: 16_380 }],
];

function written(handle: WorkbookHandle, at: CellAddr, input: string, copyFrom?: CellAddr): string {
  try {
    handle.editCells(0, [{ ...at, input, ...(copyFrom ? { copyFrom } : {}) }]);
    return handle.cell(0, at.row, at.col).input;
  } catch (error) {
    return `error: ${error instanceof Error ? error.message : String(error)}`;
  }
}

describe('translateFormula', () => {
  it('moves references by the copy offset like the engine tests expect', () => {
    const from = { row: 0, col: 2 };
    const to = { row: 2, col: 2 };
    expect(translateFormula('A1+$B$1+$A1+A$1', from, to)).toBe('A3+$B$1+$A3+A$1');
    expect(translateFormula('SUM(A1:B2)', from, to)).toBe('SUM(A3:B4)');
    expect(translateFormula("'Other Sheet'!A1+Sheet1!$B2", from, to)).toBe("'Other Sheet'!A3+Sheet1!$B4");
    expect(translateFormula('Table1[A1]+Rate_A1+A1', from, to)).toBe('Table1[A1]+Rate_A1+A3');
    expect(translateFormula('A1+$A$1+B$1+$A2', { row: 1, col: 1 }, { row: 0, col: 0 })).toBe('#REF!+$A$1+A$1+$A1');
    expect(translateFormula('XFD1048576', from, to)).toBe('#REF!');
  });

  it('refuses a formula past the length limit', () => {
    expect(() => translateFormula(`A1+"${'x'.repeat(32_768)}"`, { row: 0, col: 0 }, { row: 1, col: 0 }))
      .toThrow('formula exceeds the length limit');
  });

  it('writes what the engine writes for a copied formula', async () => {
    const handle = openWorkbook(await book());
    try {
      for (const formula of FORMULAS) {
        for (const [from, to] of MOVES) {
          const copied = written(handle, to, `=${formula}`, from);
          let ours: string;
          try { ours = `=${translateFormula(formula, from, to)}`; } catch (error) { ours = `error: ${(error as Error).message}`; }
          const typed = ours.startsWith('error: ') ? ours : written(handle, to, ours);
          expect({ formula, from, to, written: typed }).toEqual({ formula, from, to, written: copied });
        }
      }
    } finally {
      handle.dispose();
    }
  });
});
