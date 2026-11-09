import { afterAll, beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook, translateFormulasOn } from '../index';
import type { WorkbookHandle } from '../index';
import type { WorkbookSession } from '../session/client';
import { createTestWorkbookSession } from '../session/testHelpers';

const LEAN_WASM = resolve(import.meta.dir, './generated/xlsx_wasm_lean_bg.wasm');

async function book(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

let handle: WorkbookHandle;
let session: WorkbookSession;

beforeAll(async () => {
  await initWasm(new Uint8Array(readFileSync(LEAN_WASM)));
  handle = openWorkbook(await book());
  session = await createTestWorkbookSession(await book());
}, 30_000);

afterAll(async () => {
  handle.dispose();
  await session.dispose();
});

it('moves formulas on the lean engine, in-thread and through the session', async () => {
  const items = [
    { formula: 'A1+$B$1', from: { row: 0, col: 0 }, to: { row: 2, col: 1 } },
    { formula: 'A1', from: { row: 1_048_576, col: 0 }, to: { row: 0, col: 0 } },
  ];
  const expected: ReturnType<typeof translateFormulasOn> = [
    { ok: true, formula: 'B3+$B$1' },
    { ok: false, error: 'bad edit args: copy origin is outside the grid' },
  ];
  expect(translateFormulasOn(handle, items)).toEqual(expected);
  expect(await session.call.translateFormulas(items)).toEqual(expected);
});
