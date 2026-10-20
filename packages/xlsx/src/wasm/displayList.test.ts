import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import * as xlsx from '../index';
import { workbookDisplayListJson } from './loader';
import { createTestWorkbookSession } from '../session/testHelpers';

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
      expect(grid.colOffsets).toHaveLength(6);
      expect(grid.colIndices).toEqual([0, 1, 5, 6, 7]);
      expect(grid.rowOffsets).toHaveLength(21);
      const right = grid.colOffsets[5];
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

  it('degrades oversized viewports and warns once across all frame readers', () => {
    const handle = xlsx.openWorkbook(bytes);
    const viewport = { x: 0, y: 0, width: 100, height: 6_000_000 };
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      for (const read of [
        () => handle.displayList(viewport),
        () => handle.displayListProfiled(viewport).displayList,
        () => JSON.parse(workbookDisplayListJson(handle, viewport, 0)) as xlsx.DisplayList,
      ]) {
        const frame = read();
        expect(frame.truncated).toBe(true);
        expect((frame.grid!.rowOffsets.length - 1) * (frame.grid!.colOffsets.length - 1))
          .toBeLessThanOrEqual(xlsx.getDisplayListCellLimit());
      }
      expect(warning).toHaveBeenCalledTimes(1);
      expect(String(warning.mock.calls[0][0])).toContain('displayTooLarge');
    } finally {
      warning.mockRestore();
      handle.dispose();
    }
  });

  it('opens a dense low-zoom worker frame and scrolls with bounded hit testing', async () => {
    const sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>'
      + '<sheetFormatPr defaultRowHeight="3" customHeight="1"/><cols><col min="1" max="16384" width="0.5" customWidth="1"/></cols>'
      + '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>origin</t></is></c></row>'
      + '<row r="102"><c r="CX102" t="inlineStr"><is><t>scrolled</t></is></c></row></sheetData></worksheet>';
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    let session: xlsx.WorkbookSession | undefined;
    try {
      session = await createTestWorkbookSession(await workbookBytes(sheet), undefined, {
        initialViewport: { width: 1600, height: 1000, zoom: 0.1, dpr: 1, requestGeneration: 1 },
      });
      expect(session.state.stage).toBe('ready');
      const first = session.initialPaint!.frame;
      expect(first.displayList.truncated).toBe(true);
      expect(first.displayList.commands.some((command) => command.op === 'text' && command.text === 'origin')).toBe(true);
      const at = await session.call.cellGeometry(0, 100, 100);
      const viewport = { x: at.rect.x, y: at.rect.y, width: 16000, height: 10000 };
      for (let index = 0; index < 2; index++) {
        const frame = await session.call.frame(viewport);
        const grid = frame.displayList.grid!;
        expect(frame.displayList.truncated).toBe(true);
        expect(frame.viewport).toEqual(viewport);
        expect(frame.displayList.commands.some((command) => command.op === 'text' && command.text === 'scrolled')).toBe(true);
        expect(xlsx.cellAtPoint(grid, 1, 1)).toEqual({ row: 0, col: 0 });
        expect(xlsx.cellAtPoint(grid, grid.colOffsets[1] + 1, grid.rowOffsets[1] + 1)).toEqual({ row: 101, col: 101 });
        expect(xlsx.cellAtPoint(grid, 1, grid.rowOffsets[grid.rowOffsets.length - 1] + 1)).toBeNull();
        expect(xlsx.cellRect(grid, grid.rowIndices![grid.rowIndices!.length - 1] + 1, 101)).toBeNull();
        expect(xlsx.rangeRect(grid, {
          top: grid.rowIndices![grid.rowIndices!.length - 1] + 1, bottom: grid.rowIndices![grid.rowIndices!.length - 1] + 2,
          left: 101, right: 101,
        })).toBeNull();
        expect((grid.rowOffsets.length - 1) * (grid.colOffsets.length - 1)).toBeLessThanOrEqual(250000);
      }
      expect(warning).toHaveBeenCalledTimes(1);
      expect(String(warning.mock.calls[0][0])).toContain('displayTooLarge');
      expect(session.failure).toBeUndefined();
    } finally {
      await session?.dispose();
      warning.mockRestore();
    }
  });

  it('keeps typed cell-limit errors for print ranges', () => {
    const handle = xlsx.openWorkbook(bytes);
    try {
      expect(() => handle.printDisplayList(0, 'A1:XFD1048576', {
        dpi: 96, maxDigitWidth: 7, defaultRowHeightPt: 15,
        fontSizePt: 11, fontFamily: 'Calibri', fontAscent: 10, fontDescent: 3,
      }, false)).toThrow(xlsx.DisplayTooLargeError);
    } finally {
      handle.dispose();
    }
  });
});
