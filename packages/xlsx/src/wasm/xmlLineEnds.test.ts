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

  it('keeps a typed carriage return across save and reopen', async () => {
    const handle = openWorkbook(await book('\n'));
    let saved: Uint8Array;
    try {
      handle.editCell(0, 8, 9, 'a\rb');
      saved = new Uint8Array(handle.save());
    } finally {
      handle.dispose();
    }
    const sheet = await (await JSZip.loadAsync(saved)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toContain('a&#13;b');
    const reopened = openWorkbook(saved);
    try {
      const read = reopened.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'J9' } }] });
      if (!read.ok) throw new Error(read.failure.message);
      expect(read.ranges[0].cells[0][0].value).toEqual({ kind: 'text', value: 'a\rb' });
    } finally {
      reopened.dispose();
    }
  });
});
