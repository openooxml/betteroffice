import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import type { XlsxExportResult } from '../exports';
import * as wasm from '../wasm/loader';
import { createWorkbookSession } from './client';
import { createWorkbookSessionHost } from './host';
import { translateWorkbookSheetIds } from './commandHandlers';

let bytes: Uint8Array;
const viewport = { x: 0, y: 0, width: 120, height: 48 };
const metrics = { dpi: 144, maxDigitWidth: 9, defaultRowHeightPt: 18, defaultColumnWidth: 10,
  fontSizePt: 12, fontFamily: 'Arial', fontAscent: 12, fontDescent: 4 };

beforeAll(async () => {
  await wasm.initWasm(new Uint8Array(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url))));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
  for (const sheet of [1, 2]) zip.file(`xl/worksheets/sheet${sheet}.xml`, `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>${sheet}</v></c><c r="B1"><f>A1*2</f><v>${sheet * 2}</v></c></row></sheetData></worksheet>`);
  bytes = await zip.generateAsync({ type: 'uint8array' });
});

function content<T>(result: XlsxExportResult<T>): T {
  if (!result.ok) throw new Error(result.failure.message);
  return result.content;
}

async function opened() {
  const pair = createInProcessPair();
  const transferred: ArrayBuffer[] = [];
  createWorkbookSessionHost({ ...pair.host, post(message, transfer) {
    const buffers = isHostMessage(message) && message.kind === 'reply' && message.ok
      ? (transfer ?? []).filter((item): item is ArrayBuffer => item instanceof ArrayBuffer) : [];
    pair.host.post(message, transfer);
    transferred.push(...buffers);
  } });
  const session = await createWorkbookSession(bytes, {}, pair.client);
  return { session, transferred };
}

test('structured and Markdown exports match stored results, anchors and limits after edits', async () => {
  const { session } = await opened();
  const main = wasm.openWorkbook(bytes);
  try {
    const edit = session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
      kind: 'editCell', payload: { args: [1, 0, 0, '17'] } });
    main.editCell(1, 0, 0, '17');
    const options = { scope: [{ sheet: 1, range: 'A1:B2' }] };
    const exported = await session.call.exportStructured(options);
    expect(await edit).toMatchObject({ status: 'applied' });
    expect(exported.version).toBe(session.confirmedBasis.version);
    expect(content(exported)).toEqual(content(main.exportStructured(options)));
    expect(content(exported)).toMatchObject({ anchorScope: 'session',
      calculation: { policy: 'asStored', freshness: 'unverified' } });
    for (const limits of [{}, { maxRows: 1 }, { maxColumns: 1 }, { maxCells: 1 }, { maxBytes: 128 }]) {
      const markdown = await session.call.exportMarkdown(options, limits);
      expect(markdown.version).toBe(exported.version);
      expect(markdown).toEqual({ ...main.exportMarkdown(options, limits), version: exported.version });
      if (markdown.ok) expect(Array.from(markdown.content.markdown.match(/<!-- xlsx-export:\d+ -->/g) ?? []))
        .toEqual(markdown.content.anchors.map((anchor) => anchor.marker));
    }
    for (const limits of [{ maxCells: 1 }, { maxBytes: 1 }, { scope: [{ sheet: 8 }] }]) {
      const actual = await session.call.exportStructured(limits);
      const expected = main.exportStructured(limits);
      expect(actual).toEqual({ ...expected, version: exported.version });
    }
    expect(session.state.activeSheet).toBe(0);
  } finally { main.dispose(); await session.dispose(); }
});

test('export anchors keep command catalog identities after removing a sheet', async () => {
  const { session } = await opened();
  const main = wasm.openWorkbook(bytes);
  try {
    await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
      kind: 'applyOps', payload: { args: [[{ type: 'removeSheet', index: 0 }]] } });
    main.applyOps([{ type: 'removeSheet', index: 0 }]);
    const ids = session.editState.sheets!.map((sheet) => sheet.id);
    const positional = main.sheetInfo().sheetIds;
    for (const method of ['exportStructured', 'exportMarkdown'] as const) {
      const actual = await session.call[method]({ scope: [{ sheet: 0 }] });
      const expected = translateWorkbookSheetIds(main[method]({ scope: [{ sheet: 0 }] }), positional, ids);
      expect(actual).toEqual({ ...expected, version: session.confirmedBasis.version });
    }
    const result = content(await session.call.exportStructured());
    expect(result.sheets[0].anchor).toMatchObject({ sheet: { sheetId: ids[0] } });
    const anchor = result.sheets[0].cells[0].anchor;
    if (anchor.kind !== 'cell') throw new Error('Missing cell anchor');
    expect(await session.call.readCells({ ranges: [{ sheetId: anchor.sheet.sheetId,
      range: { kind: 'a1', a1: anchor.a1 } }] })).toMatchObject({ ok: true });
  } finally { main.dispose(); await session.dispose(); }
});

test('exports budget command sheet identities after adding a sheet', async () => {
  const { session } = await opened();
  const main = wasm.openWorkbook(bytes);
  try {
    const operations = [{ type: 'addSheet', index: 2, name: 'Added' }];
    expect(await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
      kind: 'applyOps', payload: { args: [operations] } })).toMatchObject({ status: 'applied' });
    main.applyOps(operations);
    const edits = Array.from({ length: 128 }, (_, row) => ({ row, col: 0, input: `row-${row + 1}` }));
    expect(await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 2,
      kind: 'editCells', payload: { args: [2, edits] } })).toMatchObject({ status: 'applied' });
    main.editCells(2, edits);
    const scope = [{ sheet: 2, range: 'A1:A128' }];
    const raw = content(main.exportStructured({ scope }));
    const ids = session.editState.sheets!.map((sheet) => sheet.id);
    expect(ids[2]).toMatch(/^command-sheet:/);
    const translated = translateWorkbookSheetIds(raw, main.sheetInfo().sheetIds, ids);
    const encoder = new TextEncoder();
    const maxBytes = encoder.encode(JSON.stringify(raw)).byteLength + 512;
    expect(content(main.exportStructured({ scope, maxBytes }))).toEqual(raw);
    expect(encoder.encode(JSON.stringify(translated)).byteLength).toBeGreaterThan(maxBytes);
    for (const method of ['exportStructured', 'exportMarkdown'] as const) {
      expect(main[method]({ scope, maxBytes }).ok).toBe(true);
      expect(await session.call[method]({ scope, maxBytes })).toMatchObject({
        ok: false, version: session.confirmedBasis.version,
        failure: { code: 'limit-exceeded', target: null },
      });
      const result = await session.call[method]({ scope, maxBytes: encoder.encode(JSON.stringify(translated)).byteLength });
      const expected = translateWorkbookSheetIds(main[method]({ scope }), main.sheetInfo().sheetIds, ids);
      expect(result).toEqual({ ...expected, version: session.confirmedBasis.version });
    }
    expect(session.failure).toBeUndefined();
  } finally { main.dispose(); await session.dispose(); }
});

for (const method of ['exportStructured', 'exportMarkdown'] as const) {
  for (const maxBytes of [NaN, Infinity]) {
    test(`${method} budgets expanded sheet identities with maxBytes ${maxBytes}`, async () => {
      const { session } = await opened();
      const main = wasm.openWorkbook(bytes);
      try {
        const operations = [{ type: 'addSheet', index: 2, name: 'Added' }];
        expect(await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
          kind: 'applyOps', payload: { args: [operations] } })).toMatchObject({ status: 'applied' });
        main.applyOps(operations);
        const edits = Array.from({ length: 128 }, (_, row) => ({ row, col: 0, input: `row-${row + 1}` }));
        main.editCells(2, edits);
        const scope = [{ sheet: 2, range: 'A1:A128' }];
        const encoder = new TextEncoder();
        const defaultMaxBytes = 8 * 1024 * 1024;
        const size = encoder.encode(JSON.stringify(content(main.exportStructured({ scope })))).byteLength;
        const padding = 'x'.repeat(Math.floor((defaultMaxBytes - size - 512) / (2 * edits.length)));
        const padded = edits.map((edit) => ({ ...edit, input: edit.input + padding }));
        expect(await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 2,
          kind: 'editCells', payload: { args: [2, padded] } })).toMatchObject({ status: 'applied' });
        main.editCells(2, padded);
        const raw = content(main.exportStructured({ scope, maxBytes: 2 * defaultMaxBytes }));
        expect(raw.truncated).toBe(false);
        expect(raw.sheets[0].cells).toHaveLength(edits.length);
        expect(content(main.exportStructured({ scope, maxBytes }))).toEqual(raw);
        const ids = session.editState.sheets!.map((sheet) => sheet.id);
        expect(ids[2]).toMatch(/^command-sheet:/);
        const translated = translateWorkbookSheetIds(raw, main.sheetInfo().sheetIds, ids);
        expect(encoder.encode(JSON.stringify(translated)).byteLength).toBeGreaterThan(defaultMaxBytes);
        expect(main[method]({ scope, maxBytes }).ok).toBe(true);
        expect(await session.call[method]({ scope, maxBytes })).toMatchObject({
          ok: false, version: session.confirmedBasis.version,
          failure: { code: 'limit-exceeded', target: null },
        });
        expect(session.failure).toBeUndefined();
      } finally { main.dispose(); await session.dispose(); }
    }, 30_000);
  }
}

test('PNG and print exports use the requested sheet after edits without navigating', async () => {
  const { session, transferred } = await opened();
  const main = wasm.openWorkbook(bytes);
  try {
    const before = session.state;
    const editing = session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
      kind: 'editCell', payload: { args: [1, 0, 0, '23'] } });
    main.editCell(1, 0, 0, '23');
    const printing = session.call.printDisplayList(1, 'A1:B2', metrics, false);
    expect(await printing).toEqual(main.printDisplayList(1, 'A1:B2', metrics, false));
    await editing;
    main.setActiveSheet(1);
    expect(session.pngExport).toBe(wasm.isPngExportAvailable());
    const renders = [
      { call: () => session.call.renderPng(viewport, { sheet: 1 }), expected: () => main.renderPng(viewport) },
      { call: () => session.call.renderRangePng({ range: 'A1:B2', scale: 1.5 }, { sheet: 1 }),
        expected: () => main.renderRangePng({ range: 'A1:B2', scale: 1.5 }) },
    ];
    for (const render of renders) {
      if (session.pngExport) {
        const png = await render.call();
        expect(png).toEqual(render.expected());
        expect([...png.slice(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
      } else await expect(render.call()).rejects.toMatchObject({
        name: 'WorkbookPngUnavailableError', code: 'png-unavailable',
      });
    }
    expect(transferred.length).toBeGreaterThanOrEqual(session.pngExport ? 3 : 1);
    expect(transferred.every((buffer) => buffer.byteLength === 0)).toBe(true);
    expect(session.state.activeSheet).toBe(before.activeSheet);
    expect((await session.call.frame(viewport)).sheet).toBe(before.activeSheet);
    expect(session.failure).toBeUndefined();
  } finally { main.dispose(); await session.dispose(); }
});

test('PNG unavailability is typed and leaves the session usable', async () => {
  const availability = spyOn(wasm, 'isPngExportAvailable').mockReturnValue(false);
  try {
    const { session } = await opened();
    try {
      expect(session.pngExport).toBe(false);
      for (const call of [() => session.call.renderPng(viewport), () => session.call.renderRangePng()]) {
        await expect(call()).rejects.toMatchObject({ name: 'WorkbookPngUnavailableError', code: 'png-unavailable' });
      }
      expect(session.failure).toBeUndefined();
      expect(content(await session.call.exportStructured()).sheets).toHaveLength(2);
    } finally { await session.dispose(); }
  } finally { availability.mockRestore(); }
});
