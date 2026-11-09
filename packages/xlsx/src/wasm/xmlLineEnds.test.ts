import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import type { WorkbookHandle } from '../index';
import { initWasm, openWorkbook } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

async function book(eol: string): Promise<Uint8Array> {
  const price = `Sales[[#This Row],[Price${eol}'[GBP']]]`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rIdTable" Type="${RELS}/table" Target="../tables/table1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}" xmlns:r="${RELS}"><sheetData>`
    + '<row r="1"><c r="A1" t="inlineStr"><is><t>Year</t></is></c><c r="B1" t="inlineStr"><is><t>Price&#10;[GBP]</t></is></c><c r="C1" t="inlineStr"><is><t>Double</t></is></c>'
    + `<c r="E1"><f>SUM(Sales[Price${eol}'[GBP']])</f><v>30</v></c></row>`
    + `<row r="2"><c r="A2"><v>2003</v></c><c r="B2"><v>10</v></c><c r="C2"><f>${price}*2</f><v>20</v></c></row>`
    + `<row r="3"><c r="A3"><v>2004</v></c><c r="B3"><v>20</v></c><c r="C3"><f>${price}*2</f><v>40</v></c></row>`
    + '</sheetData><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>');
  zip.file('xl/tables/table1.xml', `<table xmlns="${MAIN}" id="1" name="Sales" displayName="Sales" ref="A1:C3"><tableColumns count="3"><tableColumn id="1" name="Year"/><tableColumn id="2" name="Price_x000a_[GBP]"/><tableColumn id="3" name="Double"/></tableColumns></table>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function values(handle: WorkbookHandle): unknown[] {
  const read = handle.readCells({
    ranges: ['C2', 'C3', 'E1'].map((a1) => ({ sheetId: 'sheet:0', range: { kind: 'a1' as const, a1 } })),
  });
  if (!read.ok) throw new Error(read.failure.message);
  return read.ranges.map((range) => range.cells[0][0].value);
}

describe('line ends in formula text', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  for (const [label, eol] of [['a line feed', '\n'], ['a literal crlf', '\r\n']] as const) {
    it(`${label} matches a multi-line table column after recalculation`, async () => {
      const handle = openWorkbook(await book(eol));
      try {
        handle.recalculate();
        expect(values(handle)).toEqual([20, 40, 30].map((value) => ({ kind: 'number', value })));
      } finally {
        handle.dispose();
      }
    });
  }

  it('keeps a literal crlf formula working across save and reopen', async () => {
    const handle = openWorkbook(await book('\r\n'));
    let saved: Uint8Array;
    try {
      handle.editCell(0, 1, 1, '11');
      expect(handle.cell(0, 1, 2).input).toBe("=Sales[[#This Row],[Price\n'[GBP']]]*2");
      saved = new Uint8Array(handle.save());
    } finally {
      handle.dispose();
    }
    const reopened = openWorkbook(saved);
    try {
      reopened.recalculate();
      expect(values(reopened)).toEqual([22, 40, 31].map((value) => ({ kind: 'number', value })));
    } finally {
      reopened.dispose();
    }
  });
});

async function controlFormulaBook(d1 = '"a&#1;b"'): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="D1" t="str"><f>${d1}</f><v>x</v></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function thrown(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  return '';
}

describe('control characters in formulas through applyOps', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('writes a new one, typed or applied, with CHAR and refuses what has no storable form', async () => {
    const handle = openWorkbook(await controlFormulaBook());
    try {
      const written = '=("a"&CHAR(1)&"b")';
      handle.applyOps([{ type: 'setCell', sheet: 0, at: { row: 0, col: 7 }, cell: { formula: '"a\u0001b"' } }]);
      handle.editCell(0, 0, 8, '="a\u0001b"');
      handle.editCells(0, [{ row: 0, col: 9, input: '="a\u0001b"' }]);
      for (const col of [7, 8, 9]) expect(handle.cell(0, 0, col).input).toBe(written);
      const read = handle.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'H1:J1' } }] });
      if (!read.ok) throw new Error(read.failure.message);
      expect(read.ranges[0].cells[0].map((cell) => cell.value)).toEqual(Array(3).fill({ kind: 'text', value: 'a\u0001b' }));
      expect(handle.cell(0, 0, 3).input).toBe('="a\u0001b"');
      const version = handle.version();
      expect(thrown(() => handle.editCell(0, 0, 10, '="a\u0000"'))).toContain('U+0000, which a workbook file cannot store and CHAR cannot produce');
      expect(thrown(() => handle.applyOps([{ type: 'setCell', sheet: 0, at: { row: 0, col: 10 }, cell: { formula: '{"a\u0001",1}' } }]))).toContain('inside an array constant');
      expect(handle.version()).toBe(version);
      const sheet = await (await JSZip.loadAsync(new Uint8Array(handle.save()))).file('xl/worksheets/sheet1.xml')!.async('string');
      expect(sheet.match(/<f>\(&quot;a&quot;&amp;CHAR\(1\)&amp;&quot;b&quot;\)<\/f>/g)?.length).toBe(3);
      expect(sheet.match(/a&#1;b/g)?.length).toBe(1);
    } finally {
      handle.dispose();
    }
  });

  it('pastes, moves, shifts and re-commits an opened one', async () => {
    const handle = openWorkbook(await controlFormulaBook());
    try {
      expect(thrown(() => handle.editCells(0, [{ row: 0, col: 9, input: '="a\u0001b"', copyFrom: { row: 0, col: 3 } }]))).toBe('');
      expect(thrown(() => handle.applyOps([
        { type: 'insertRows', sheet: 0, at: 0, count: 1 },
        { type: 'setCell', sheet: 0, at: { row: 1, col: 3 }, cell: { formula: '"a\u0001b"' } },
      ]))).toBe('');
      expect(thrown(() => handle.moveRange(0, 'D2', 'G2'))).toBe('');
      expect(handle.cell(0, 1, 6).input).toBe('="a\u0001b"');
      const sheet = await (await JSZip.loadAsync(new Uint8Array(handle.save()))).file('xl/worksheets/sheet1.xml')!.async('string');
      expect(sheet.match(/a&#1;b/g)?.length).toBe(2);
    } finally {
      handle.dispose();
    }
  });

  it('re-commits an opened prefixed formula as displayed', async () => {
    const handle = openWorkbook(await controlFormulaBook('_xlfn.CONCAT("a&#1;b","")'));
    try {
      const shown = 'CONCAT("a\u0001b","")';
      expect(thrown(() => handle.applyOps([{ type: 'setCell', sheet: 0, at: { row: 0, col: 3 }, cell: { formula: shown } }]))).toBe('');
      expect(thrown(() => handle.editCell(0, 0, 3, `=${shown}`))).toBe('');
      expect(handle.cell(0, 0, 3).input).toBe(`=${shown}`);
    } finally {
      handle.dispose();
    }
  });

  it('keeps a typed carriage return across save and reopen as Excel spells it', async () => {
    const handle = openWorkbook(await controlFormulaBook());
    let saved: Uint8Array;
    try {
      handle.editCell(0, 8, 9, 'a\rb');
      saved = new Uint8Array(handle.save());
    } finally {
      handle.dispose();
    }
    const sheet = await (await JSZip.loadAsync(saved)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toContain('a_x000D_b');
    const reopened = openWorkbook(saved);
    try {
      expect(reopened.cell(0, 8, 9).input).toBe('a\rb');
    } finally {
      reopened.dispose();
    }
  });
});
