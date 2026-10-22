import { afterEach, beforeAll, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const ROWS = 20_000;

async function chainWorkbook(): Promise<Uint8Array> {
  const rows = ['<row r="1"><c r="A1"><v>1</v></c></row>'];
  for (let row = 2; row <= ROWS; row++) {
    rows.push(`<row r="${row}"><c r="A${row}"><f>A${row - 1}+1</f></c></row>`);
  }
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

afterEach(() => {
  delete (globalThis as { __betterofficeEngineProgress?: () => void }).__betterofficeEngineProgress;
});

test('the first open of a module reports progress from its recalculation', async () => {
  const bytes = await chainWorkbook();
  let ticks = 0;
  Object.assign(globalThis, { __betterofficeEngineProgress: () => { ticks++; } });

  const workbook = openWorkbook(bytes);
  try {
    expect(ticks).toBeGreaterThanOrEqual(4);
  } finally {
    workbook.dispose();
  }
});
