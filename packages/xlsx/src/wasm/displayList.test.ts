import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import * as xlsx from '../index';
import { workbookDisplayListJson } from './loader';

const LAST_ROW = 1_048_575;
const LAST_COL = 16_383;

const EDGE_SHEET = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>edge</t></is></c></row></sheetData></worksheet>';

function hiddenTailSheet(): string {
  const rows = Array.from({ length: 20 }, (_, row) => {
    const cells = 'ABCDEFGH'.split('').map((col) => `<c r="${col}${row + 1}" t="inlineStr"><is><t>${col}${row + 1}</t></is></c>`).join('');
    return `<row r="${row + 1}" ht="15" customHeight="1">${cells}</row>`;
  }).join('');
  return '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/><cols>'
    + '<col min="1" max="2" width="10" customWidth="1"/><col min="3" max="5" width="10" hidden="1" customWidth="1"/>'
    + '<col min="6" max="8" width="10" customWidth="1"/><col min="9" max="16384" width="9.14" hidden="1"/>'
    + `</cols><sheetData>${rows}<row r="1048576" hidden="1"/></sheetData></worksheet>`;
}

async function workbookBytes(sheet = EDGE_SHEET): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', sheet);
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

  it('renders and hit-tests sheets whose unused columns and rows are hidden', async () => {
    const handle = xlsx.openWorkbook(await workbookBytes(hiddenTailSheet()));
    try {
      const frame = handle.displayList({ x: 0, y: 0, width: 1600, height: 1000 });
      const grid = frame.grid!;
      expect(grid.colOffsets).toHaveLength(9);
      expect(grid.rowOffsets).toHaveLength(21);
      const right = grid.colOffsets[8];
      const bottom = grid.rowOffsets[20];
      expect(xlsx.cellAtPoint(grid, grid.colOffsets[2] + 1, 1)).toEqual({ row: 0, col: 5 });
      expect(xlsx.cellAtPoint(grid, right - 1, bottom - 1)).toEqual({ row: 19, col: 7 });
      expect(xlsx.cellAtPoint(grid, right + 1, 1)).toBeNull();
      expect(xlsx.cellAtPoint(grid, 1, bottom + 1)).toBeNull();
      expect(frame.commands.some((command) => command.op === 'text' && command.text === 'H20')).toBe(true);
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
