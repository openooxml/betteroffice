import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CALCULATION = { nowSerial: 46_327.5, randSeed: 12_345 };

async function book(calcPr: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><f>SUM(2,3)</f><v>999</v></c><c r="B1"><f>REPRO_UNKNOWN_FUNCTION(1)</f><v>999</v></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function savedCaches(bytes: Uint8Array): Promise<(string | null)[]> {
  const sheet = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
  return [...sheet.matchAll(/<c [^>]*>(.*?)<\/c>/g)].map((cell) => /<v>(.*?)<\/v>/.exec(cell[1])?.[1] ?? null);
}

describe('saving recalculated values', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('keeps the file byte for byte when nothing asked for a recalculation', async () => {
    const source = await book('');
    const handle = openWorkbook(source, { calculation: CALCULATION });
    try {
      expect(new Uint8Array(handle.save())).toEqual(new Uint8Array(source));
    } finally {
      handle.dispose();
    }
  });

  it('saves the values recalculate() produced and keeps a cache it could not evaluate', async () => {
    const handle = openWorkbook(await book(''), { calculation: CALCULATION });
    try {
      const status = handle.recalculate();
      expect(status).toEqual(handle.calculationStatus());
      expect(await savedCaches(handle.save())).toEqual(['5', '999']);
    } finally {
      handle.dispose();
    }
  });

  it('saves the values calculated at open when the file asks for a full calculation', async () => {
    const handle = openWorkbook(await book('<calcPr fullCalcOnLoad="1" forceFullCalc="1"/>'), { calculation: CALCULATION });
    try {
      expect(await savedCaches(handle.save())).toEqual(['5', '999']);
    } finally {
      handle.dispose();
    }
  });
});
