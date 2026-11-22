import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

import { initWasm, openWorkbook, type WorkbookHandle } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CALCULATION = { nowSerial: 46_327.5, randSeed: 12_345 };

const STALE_CELLS = '<row r="1"><c r="A1"><f>SUM(2,3)</f><v>999</v></c><c r="B1"><f>REPRO_UNKNOWN_FUNCTION(1)</f><v>999</v></c></row>';

async function book(calcPr: string, cells = STALE_CELLS): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData>${cells}</sheetData></worksheet>`);
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

  it('fills missing caches when the file asks for a full calculation', async () => {
    const cells = STALE_CELLS.replace('</row>', '<c r="C1"><f>1+1</f></c></row>');
    const handle = openWorkbook(await book('<calcPr fullCalcOnLoad="1" forceFullCalc="1"/>', cells), { calculation: CALCULATION });
    try {
      expect(await savedCaches(handle.save())).toEqual(['999', '999', '2']);
    } finally {
      handle.dispose();
    }
  });

  it('keeps cached array booleans on a save without edits', async () => {
    const cells = '<row r="1"><c r="A1" t="b"><f t="array" ref="A1:A3">FALSE()</f><v>1</v></c></row>'
      + '<row r="2"><c r="A2" t="b"><v>1</v></c></row>'
      + '<row r="3"><c r="A3" t="b"><v>1</v></c></row>';
    const source = await book('<calcPr fullCalcOnLoad="1"/>', cells);
    const handle = openWorkbook(source);
    try {
      const sheet = await (await JSZip.loadAsync(handle.save())).file('xl/worksheets/sheet1.xml')!.async('string');
      expect(sheet).toContain(`<sheetData>${cells}</sheetData>`);
    } finally {
      handle.dispose();
    }
  });

  it('keeps iterative caches when the file asks for a full calculation', async () => {
    const cells = '<row r="1">'
      + '<c r="A1"><f>(A1+10)/2</f><v>10</v></c><c r="B1"><f>A1*2</f><v>20</v></c>'
      + '<c r="C1"><v>5</v></c><c r="D1"><f>C1+1</f><v>999</v></c>'
      + '<c r="E1"><f>F1/2+10</f><v>20</v></c><c r="F1"><f>E1</f><v>20</v></c></row>';
    const source = await book('<calcPr iterate="1" fullCalcOnLoad="1"/>', cells);
    const handle = openWorkbook(source);
    try {
      expect(await savedCaches(handle.save())).toEqual(['10', '20', '5', '999', '20', '20']);
    } finally {
      handle.dispose();
    }
  });
});

describe("recalculate: 'stale'", () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  const CACHED = '<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>A1*2</f><v>999</v></c><c r="C1"><f>SUM(2,3)</f><v>7</v></c></row>';
  const UNCACHED = CACHED.replace('</row>', '<c r="D1"><f>1+1</f></c></row>');

  function values(handle: WorkbookHandle, a1: string): unknown[] {
    const read = handle.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1 } }] });
    if (!read.ok) throw new Error(read.failure.message);
    return read.ranges[0].cells[0].map((cell) => cell.value);
  }

  const number = (value: number) => ({ kind: 'number', value });

  it('shows cached results, while the default recalculates them', async () => {
    const source = await book('', CACHED);
    for (const [options, expected] of [
      [{}, [5, 10, 5]],
      [{ recalculate: 'all' }, [5, 10, 5]],
      [{ recalculate: 'stale' }, [5, 999, 7]],
      [{ recalculate: 'stale', calculation: CALCULATION }, [5, 999, 7]],
    ] as const) {
      const handle = openWorkbook(source, options);
      try {
        expect(values(handle, 'A1:C1')).toEqual(expected.map(number));
      } finally {
        handle.dispose();
      }
    }
  });

  it('recalculates when a formula has no cached result', async () => {
    const handle = openWorkbook(await book('', UNCACHED), { recalculate: 'stale' });
    try {
      expect(values(handle, 'A1:D1')).toEqual([5, 10, 5, 2].map(number));
    } finally {
      handle.dispose();
    }
  });

  it('keeps the file byte for byte and recalculates what an edit reaches', async () => {
    const source = await book('', CACHED);
    const handle = openWorkbook(source, { recalculate: 'stale' });
    try {
      expect(new Uint8Array(handle.save())).toEqual(new Uint8Array(source));
      handle.editCell(0, 0, 0, '6');
      expect(values(handle, 'A1:C1')).toEqual([6, 12, 7].map(number));
    } finally {
      handle.dispose();
    }
  });

  it('rejects an unknown mode and a collaborative replica', async () => {
    const source = await book('', CACHED);
    expect(() => openWorkbook(source, { recalculate: 'none' as 'all' })).toThrow(TypeError);
    expect(() => openWorkbook(source, { recalculate: 'stale', collaborative: true })).toThrow(TypeError);
  });
});
