import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { initWasm, openWorkbook, type WorkbookHandle } from '../wasm/loader';
import type { DisplayList } from '../display-list/types';
import { applyWorkbookCommand, createWorkbookCommandState } from './commandHandlers';
import type { WorkbookCommand } from './commandProtocol';

const view = { x: 0, y: 0, width: 640, height: 800 };
const red = '#ffc7ce';

beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

async function positiveWorkbook(zero?: string): Promise<Uint8Array> {
  const rows = Array.from({ length: 30 }, (_, index) => {
    const row = index + 1;
    const cells = ['A', 'B', 'C', 'D'].map((letter, col) =>
      `<c r="${letter}${row}"><v>${`${letter}${row}` === zero ? 0 : row + col}</v></c>`).join('');
    return `<row r="${row}">${cells}</row>`;
  }).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData><conditionalFormatting sqref="A1:D30"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function fills(frame: DisplayList): string[] {
  return (frame.commands as unknown as Record<string, unknown>[])
    .filter((command) => command.op === 'fillRect')
    .map((command) => `${command.x} ${command.y} ${command.w} ${command.h} ${command.color}`)
    .sort();
}

function settled(handle: WorkbookHandle): DisplayList {
  for (let frames = 0; frames < 20; frames++) {
    const frame = handle.displayList(view);
    if (!frame.conditionalPending) return frame;
  }
  throw new Error('conditional formats still pending after 20 frames');
}

function stopped(handle: WorkbookHandle): DisplayList {
  let now = 0;
  const clock = spyOn(performance, 'now').mockImplementation(() => (now += 1000));
  try { return handle.displayList(view); }
  finally { clock.mockRestore(); }
}

function typeIntoD30(handle: WorkbookHandle, input: string): void {
  const identity = { session: 'book', incarnation: 'worker' };
  const basis = { epoch: 1, version: handle.version() };
  const { sheetIds } = handle.sheetInfo();
  const sheetId = sheetIds[0]!;
  const state = createWorkbookCommandState(identity, basis, { sheetId, range: { kind: 'a1', a1: 'A1' } }, sheetIds);
  const command = { ...identity, seq: 1, basis, kind: 'editCell',
    payload: { args: [0, 29, 3, input], target: { sheetId, range: { kind: 'a1', a1: 'D30' } } } } as WorkbookCommand;
  expect(applyWorkbookCommand(handle, state, command).statuses).toMatchObject([{ status: 'applied', committed: true }]);
}

test('a typed edit that changes no result keeps every conditional fill on screen until decided', async () => {
  const handle = openWorkbook(await positiveWorkbook());
  try {
    const before = fills(settled(handle));
    expect(before.filter((fill) => fill.endsWith(red))).toHaveLength(120);
    typeIntoD30(handle, '7');
    const first = stopped(handle);
    expect(first.conditionalPending).toBeGreaterThan(0);
    expect(fills(first)).toEqual(before);
    let frame = first;
    let frames = 1;
    while (frame.conditionalPending) {
      expect(frames).toBeLessThan(40);
      frame = stopped(handle);
      frames++;
      expect(fills(frame)).toEqual(before);
    }
    expect(frames).toBeGreaterThan(1);
  } finally { handle.dispose(); }
});

test('a typed edit that changes a result shows its previous fill until decided', async () => {
  const handle = openWorkbook(await positiveWorkbook());
  const reference = openWorkbook(await positiveWorkbook('D30'));
  try {
    const before = fills(settled(handle));
    typeIntoD30(handle, '0');
    const first = stopped(handle);
    expect(first.conditionalPending).toBeGreaterThan(0);
    expect(fills(first)).toEqual(before);
    const expected = fills(settled(reference));
    expect(expected.filter((fill) => fill.endsWith(red))).toHaveLength(119);
    expect(fills(settled(handle))).toEqual(expected);
  } finally {
    handle.dispose();
    reference.dispose();
  }
});
