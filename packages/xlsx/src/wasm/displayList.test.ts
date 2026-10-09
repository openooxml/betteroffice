import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import * as xlsx from '../index';
import { workbookDisplayListJson } from './loader';

const LAST_ROW = 1_048_575;
const LAST_COL = 16_383;

async function workbookBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>edge</t></is></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

describe('display list viewport limits', () => {
  let bytes: Uint8Array;
  beforeAll(async () => {
    await xlsx.initWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm'))));
    bytes = await workbookBytes();
  });

  it('clamps viewports past the grid end', () => {
    const handle = xlsx.openWorkbook(bytes);
    try {
      const corner = handle.cellRect(0, LAST_ROW, LAST_COL);
      const viewport = { x: corner.x, y: corner.y, width: corner.w * 4, height: corner.h * 4 };
      const frame = handle.displayList(viewport);
      expect(frame.grid!.startRow).toBe(LAST_ROW);
      expect(frame.grid!.startCol).toBe(LAST_COL);
      expect(frame.grid!.rowOffsets).toEqual([0, corner.h]);
      expect(frame.grid!.colOffsets).toEqual([0, corner.w]);
      expect(frame.commands.some((command) => command.op === 'text' && command.text === 'edge')).toBe(true);
      expect(handle.displayListProfiled(viewport).displayList).toEqual(frame);
      expect(JSON.parse(workbookDisplayListJson(handle, viewport, 0))).toEqual(frame);
      const outside = handle.displayList({
        x: corner.x + corner.w * 2, y: corner.y + corner.h * 2, width: 200, height: 100,
      });
      expect(outside.grid!.rowOffsets).toEqual([]);
      expect(outside.grid!.colOffsets).toEqual([]);
      expect(outside.commands).toHaveLength(1);
    } finally {
      handle.dispose();
    }
  });

  it('reports typed oversized viewports and exposes the cell limit', () => {
    const handle = xlsx.openWorkbook(bytes);
    const viewport = { x: 0, y: 0, width: 100, height: 6_000_000 };
    try {
      for (const read of [
        () => handle.displayList(viewport),
        () => handle.displayListProfiled(viewport),
        () => workbookDisplayListJson(handle, viewport, 0),
      ]) {
        let failure: unknown;
        try { read(); } catch (error) { failure = error; }
        expect(failure).toMatchObject({ code: 'displayTooLarge', cells: 600_002, maxCells: 250_000 });
        expect(failure).toBeInstanceOf(xlsx.DisplayTooLargeError);
        expect((failure as xlsx.DisplayTooLargeError).maxCells).toBe(xlsx.getDisplayListCellLimit());
      }
    } finally {
      handle.dispose();
    }
  });
});
