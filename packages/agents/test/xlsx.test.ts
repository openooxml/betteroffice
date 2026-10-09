import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { openWorkbook } from '@betteroffice/xlsx';
import { attachXlsx, openXlsx } from '../src/xlsx';
import { xlsxFixture } from './format-fixtures';

test('XLSX header formatting preserves the implicit workbook font and reports it accurately', async () => {
  const source = await readFile(resolve(import.meta.dir, '../eval/formats/fixtures/xlsx-04-style-header.xlsx'));
  const book = await openXlsx(source);
  try {
    const result = book.formatRanges({ version: book.overview().version, edits: [{ sheet: 'sheet:0', range: 'A1:C1', bold: true, fill: '#D9EAF7' }] });
    expect(result.items[0].formatting).toMatchObject({ fontFamily: 'DejaVu Sans', fontSize: 11, bold: true });
    const zip = await JSZip.loadAsync(await book.export());
    const styles = await zip.file('xl/styles.xml')!.async('string');
    const original = await JSZip.loadAsync(source);
    const originalStyles = await original.file('xl/styles.xml')!.async('string');
    expect(styles.match(/<font><\/font>/g)?.length ?? 0).toBe(originalStyles.match(/<font><\/font>/g)?.length ?? 0);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const fonts = styles.match(/<font>.*?<\/font>/g)!;
    const xfs = styles.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/)![1].match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!;
    for (const address of ['A1', 'B1', 'C1']) {
      const xf = Number(sheet.match(new RegExp(`<c r="${address}"[^>]* s="(\\d+)"`))![1]);
      const font = Number(xfs[xf].match(/fontId="(\d+)"/)![1]);
      expect(fonts[font]).toContain('<name val="DejaVu Sans"/>');
      expect(fonts[font]).toContain('<sz val="11"/>');
      expect(fonts[font]).toContain('<b/>');
    }
  } finally { book.close(); }
});

async function tableFixture() {
  const zip = await JSZip.loadAsync(await xlsxFixture());
  const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('</worksheet>', `<tableParts count="1"><tablePart xmlns:r="${ns}" r:id="rIdTable"/></tableParts></worksheet>`));
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable" Type="${ns}/table" Target="../tables/table1.xml"/></Relationships>`);
  zip.file('xl/tables/table1.xml', '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Items" displayName="Items" ref="A3:C10" totalsRowShown="0"><autoFilter ref="A3:C10"/><tableColumns count="3"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Value"/><tableColumn id="3" name="Count"/></tableColumns></table>');
  return zip.generateAsync({ type: 'uint8array' });
}

function value(book: Awaited<ReturnType<typeof openXlsx>>, sheet: string, range: string) {
  return book.readCells({ sheet, range }).cells.map(cell => cell.value);
}

test('XLSX outline includes dimensions, used ranges, tables, names, and paginated sheets', async () => {
  const book = await openXlsx(await tableFixture());
  try {
    const page = book.outline({ limit: 1 });
    expect(page.items).toHaveLength(1);
    expect(page.nextOffset).toBe(1);
    expect(page.items[0].dimensions.rows).toBeGreaterThan(10);
    expect(page.items[0].usedRange).toBeString();
    expect(page.items[0].tables[0]).toMatchObject({ name: 'Items', range: 'A3:C10', columnCount: 3 });
    expect(book.outline({ offset: 1 }).items[0].sheetId).not.toBe(page.items[0].sheetId);
    expect(() => book.outline({ sheet: 'missing' })).toThrow('Valid sheets');
  } finally { book.close(); }
  const named = await openXlsx(await readFile(resolve(import.meta.dir, '../../xlsx/test-fixtures/defined-names.xlsx')));
  try { expect(named.outline().definedNameCount).toBeGreaterThan(0); }
  finally { named.close(); }
});

test('XLSX reads paginate blanks and return values, formulas, results, and number formats', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const sheet = book.list().items[0].sheetId;
    const page = book.readCells({ sheet, range: '$B$3:$E$3', limit: 2 });
    expect(page.nextOffset).toBe(2);
    expect(page.cells[0].computed).toEqual({ kind: 'number', value: 100 });
    expect(page.cells[0].numberFormat).toBeString();
    const next = book.readCells({ sheet, range: 'B3:E3', offset: page.nextOffset! });
    expect(next.cells[0].formula).toBe('B3+C3');
    expect(book.readCells({ sheet, range: 'XFD1048576' }).cells[0].value).toEqual({ kind: 'empty' });
    for (const range of ['XFE1', 'A0', 'B3:A1', 'A1,B2', 'Sheet1!A1']) expect(() => book.readCells({ sheet, range })).toThrow();
  } finally { book.close(); }
});

test('XLSX writes typed literals and formulas atomically and return recalculated dependents', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const initial = book.readCells({ sheet: 'sheet:0', range: 'B3:E3' });
    const result = book.writeRanges({ version: initial.version, edits: [{ sheet: 'sheet:0', range: 'B3', values: [[250]] }, { sheet: 'sheet:0', range: 'E3', formulas: [['=D3*2']] }] });
    expect(result.applied).toBe(true);
    expect(result.items[1].cells[0]).toMatchObject({ formula: 'D3*2', value: { kind: 'number', value: 614 } });
    expect(result.calculation.changedCells).toBeGreaterThan(0);
    expect(() => book.writeRanges({ version: initial.version, edits: [{ sheet: 'sheet:0', range: 'B3', values: [[1]] }] })).toThrow('Workbook changed');
    let version = result.version;
    const written = book.writeRanges({ version, edits: [{ sheet: 'sheet:0', range: 'A100:E100', values: [['001', '=2+2', true, null, 4]] }] });
    expect(written.items[0].cells.map(cell => cell.value)).toEqual([{ kind: 'text', value: '001' }, { kind: 'text', value: '=2+2' }, { kind: 'bool', value: true }, { kind: 'empty' }, { kind: 'number', value: 4 }]);
    version = written.version;
    for (const edits of [
      [{ sheet: 'sheet:0', range: 'A100', values: [[7]] }, { sheet: 'sheet:0', range: 'B1', values: [[8]] }],
      [{ sheet: 'sheet:0', range: 'A100', values: [[7]] }, { sheet: 'sheet:0', range: 'A100', values: [[8]] }],
      [{ sheet: 'sheet:0', range: 'A100:B101', values: [[7]] }],
    ]) {
      expect(() => book.writeRanges({ version, edits })).toThrow();
      expect(book.handle.version()).toBe(version);
      expect(value(book, 'sheet:0', 'A100')[0]).toEqual({ kind: 'text', value: '001' });
    }
    const reopened = await openXlsx(await book.export());
    try { expect(value(reopened, 'sheet:0', 'E3')[0]).toEqual({ kind: 'number', value: 614 }); }
    finally { reopened.close(); }
  } finally { book.close(); }
});

test('XLSX grid edits shift cells and formulas and roll back a refused batch', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    let version = book.overview().version;
    version = book.editGrid({ version, edits: [{ sheet: 'sheet:0', axis: 'rows', action: 'insert', at: 3, count: 1 }, { sheet: 'sheet:0', axis: 'columns', action: 'insert', at: 2, count: 1 }] }).version;
    expect(value(book, 'sheet:0', 'C4')[0]).toEqual({ kind: 'number', value: 100 });
    expect(book.readCells({ sheet: 'sheet:0', range: 'E4' }).cells[0].formula).toBe('C4+D4');
    version = book.editGrid({ version, edits: [{ sheet: 'sheet:0', axis: 'columns', action: 'delete', at: 2, count: 1 }, { sheet: 'sheet:0', axis: 'rows', action: 'delete', at: 3, count: 1 }] }).version;
    expect(value(book, 'sheet:0', 'B3')[0]).toEqual({ kind: 'number', value: 100 });
    expect(() => book.editGrid({ version, edits: [{ sheet: 'sheet:0', axis: 'rows', action: 'insert', at: 3, count: 1 }, { sheet: 'sheet:0', axis: 'columns', action: 'delete', at: 16384, count: 2 }] })).toThrow();
    expect(book.handle.version()).toBe(version);
  } finally { book.close(); }
});

test('XLSX sheet edits keep IDs stable and preserve content through rename, move, add and delete', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const initial = book.outline();
    const first = initial.items[0];
    const moved = book.editSheets({ version: initial.version, edits: [{ action: 'rename', sheet: first.sheetId, name: 'Budget' }, { action: 'move', sheet: first.sheetId, index: 2 }, { action: 'add', name: 'Scratch', index: 0 }] });
    expect(moved.items).toHaveLength(3);
    expect(book.outline({ sheet: first.sheetId }).items[0]).toMatchObject({ sheetId: first.sheetId, name: 'Budget', index: 3 });
    expect(value(book, first.sheetId, 'B3')[0]).toEqual({ kind: 'number', value: 100 });
    const version = moved.version;
    expect(() => book.editSheets({ version, edits: [{ action: 'add', name: 'Temporary' }, { action: 'rename', sheet: first.sheetId, name: 'Scratch' }] })).toThrow('already exists');
    expect(book.handle.version()).toBe(version);
    const scratch = moved.sheets.find(sheet => sheet.name === 'Scratch')!;
    const deleted = book.editSheets({ version, edits: [{ action: 'delete', sheet: scratch.sheetId }] });
    expect(deleted.sheets).toHaveLength(3);
    const reopened = await openXlsx(await book.export());
    try { expect(value(reopened, 'Budget', 'B3')[0]).toEqual({ kind: 'number', value: 100 }); }
    finally { reopened.close(); }
    book.handle.undo();
    expect(book.catalog().map(sheet => sheet.name)).toContain('Scratch');
  } finally { book.close(); }
});

test('XLSX formatting applies number patterns, emphasis, fill, borders, alignment and widths atomically', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const version = book.overview().version;
    const result = book.formatRanges({ version, edits: [{ sheet: 'sheet:0', range: 'B3:C3', numberFormat: '0.00', bold: true, italic: true, fill: '#FFCC00', borders: { preset: 'all', color: '#FF0000', style: 'dashed' }, alignment: { horizontal: 'center', vertical: 'middle', wrap: 'wrap' }, columnWidth: 20 }] });
    expect(result.items[0].changed).toBe(true);
    expect(result.items[0].formatting).toMatchObject({ bold: true, italic: true, fillColor: '#ffcc00', horizontalAlignment: 'center', verticalAlignment: 'middle', textWrapping: 'wrap', numberFormatPattern: '0.00' });
    expect(result.items[0].cells[0].displayText).toBe('100.00');
    expect(() => book.formatRanges({ version: result.version, edits: [{ sheet: 'sheet:0', range: 'B3', bold: false }, { sheet: 'sheet:0', range: 'C3', fill: 'red' }] })).toThrow('No edits');
    expect(book.handle.selectionFormatting(0, 'B3').bold).toBe(true);
    expect(book.handle.version()).toBe(result.version);
    expect(() => book.formatRanges({ version: result.version, edits: [{ sheet: 'sheet:0', range: 'B1', bold: true }] })).toThrow();
  } finally { book.close(); }
});

test('XLSX find/replace paginates, narrows ranges, replaces literals and refuses numeric/formula replacements', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const found = book.findReplace({ query: 'line item', sheet: 'sheet:0', range: 'A3:A6', limit: 2 });
    expect('matches' in found && found.matches).toHaveLength(2);
    expect('nextOffset' in found && found.nextOffset).toBe(2);
    const replaced = book.findReplace({ query: 'line item', replacement: 'Entry $&', version: book.handle.version(), sheet: 'sheet:0', range: 'A3:A6' });
    expect('replacedCells' in replaced && replaced.replacedCells).toBe(4);
    expect(book.readCells({ sheet: 'sheet:0', range: 'A3' }).cells[0].displayText).toStartWith('Entry $&');
    const version = book.handle.version();
    expect(() => book.findReplace({ query: '100', replacement: 'zero', version, sheet: 'sheet:0', range: 'B3' })).toThrow('non-text');
    expect(book.handle.version()).toBe(version);
    expect(() => book.findReplace({ query: 'Entry', replacement: 'X', sheet: 'sheet:0' })).toThrow('version');
  } finally { book.close(); }
});

test('XLSX sort moves whole rows and formatting, preserves headers and handles multi-key order', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    let version = book.writeRanges({ version: book.overview().version, edits: [{ sheet: 'sheet:0', range: 'A100:C104', values: [['Name', 'Value', 'Tie'], ['z', 2, 1], ['a', 1, 0], ['b', 2, 3], ['c', null, 2]] }] }).version;
    version = book.formatRanges({ version, edits: [{ sheet: 'sheet:0', range: 'A101:C101', bold: true }] }).version;
    const sorted = book.sortRange({ version, sheet: 'sheet:0', range: 'A100:C104', header: true, keys: [{ column: 'B', order: 'asc' }, { column: 'C', order: 'desc' }] });
    expect(value(book, 'sheet:0', 'A100:A104')).toEqual(['Name', 'a', 'b', 'z', 'c'].map(value => ({ kind: 'text', value })));
    expect(book.handle.selectionFormatting(0, 'A103').bold).toBe(true);
    expect(sorted.rows[0]).toEqual({ from: 102, to: 101 });
    expect(() => book.sortRange({ version: sorted.version, sheet: 'sheet:0', range: 'A1:C3', keys: [{ column: 'B', order: 'asc' }] })).toThrow('merged');
    expect(() => book.sortRange({ version: sorted.version, sheet: 'sheet:0', range: 'A100:C104', keys: [{ column: 'D', order: 'asc' }] })).toThrow('inside');
    expect(book.handle.version()).toBe(sorted.version);
  } finally { book.close(); }
});

test('XLSX previews return real PNG ranges, restore the active sheet and bound allocation', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const before = book.previewRange({ sheet: 'sheet:0', range: 'A3:C6' });
    expect([...before.png.slice(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(before.width).toBeGreaterThan(0);
    expect(before.height).toBeGreaterThan(0);
    book.formatRanges({ version: before.version, edits: [{ sheet: 'sheet:0', range: 'A3:C6', fill: '#FF0000' }] });
    expect(book.previewRange({ sheet: 'sheet:0', range: 'A3:C6' }).png).not.toEqual(before.png);
    book.handle.setActiveSheet(1);
    book.previewRange({ sheet: 'sheet:0', range: 'A3:C6' });
    expect(book.handle.sheetInfo().activeSheet).toBe(1);
    expect(() => book.previewRange({ sheet: 'sheet:0', range: 'A1:XFD1048576' })).toThrow('10000');
    expect(() => book.previewRange({ sheet: 'sheet:0', range: 'A1', scale: 100 })).toThrow('scale');
  } finally { book.close(); }
});

test('XLSX attachment shares a live editor handle and never disposes its owner', async () => {
  const bytes = await xlsxFixture();
  const init = await openXlsx(bytes);
  init.close();
  const handle = openWorkbook(bytes);
  const agent = attachXlsx(handle, { name: 'live.xlsx' });
  try {
    const read = agent.readCells({ sheet: 'sheet:0', range: 'B3' });
    handle.editCell(0, 2, 1, '42');
    expect(() => agent.writeRanges({ version: read.version, edits: [{ sheet: 'sheet:0', range: 'B3', values: [[99]] }] })).toThrow('Workbook changed');
    const result = agent.writeRanges({ version: agent.handle.version(), edits: [{ sheet: 'sheet:0', range: 'B3', values: [[99]] }] });
    expect(result.applied).toBe(true);
    expect(handle.cell(0, 2, 1).input).toBe('99');
    agent.close();
    expect(handle.cell(0, 2, 1).input).toBe('99');
    expect(() => agent.readCells({ sheet: 'sheet:0', range: 'B3' })).toThrow('closed');
  } finally { agent.close(); handle.dispose(); }
});

test('XLSX write/search outputs stay compact and paginate long text without dropping cells', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const edits = Array.from({ length: 32 }, (_, index) => ({ sheet: 'sheet:0', range: `A${100 + index}`, values: [[`${'long '.repeat(300)}${index}`]] }));
    const written = book.writeRanges({ version: book.overview().version, edits });
    expect(JSON.stringify(written).length).toBeLessThan(32000);
    const page = book.readCells({ sheet: 'sheet:0', range: 'A100:A131', limit: 100 });
    expect(JSON.stringify(page).length).toBeLessThan(17000);
    expect(page.nextOffset).toBeGreaterThan(0);
    const search = book.findReplace({ query: 'long', sheet: 'sheet:0', limit: 100 });
    expect(JSON.stringify(search).length).toBeLessThan(17000);
    if ('nextOffset' in search && search.nextOffset) expect(book.findReplace({ query: 'long', sheet: 'sheet:0', offset: search.nextOffset })).toHaveProperty('matches');
  } finally { book.close(); }
});

test('XLSX typed values ignore text number formats while preserving cell styles', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const formatted = book.formatRanges({ version: book.overview().version, edits: [{ sheet: 'sheet:0', range: 'A100:B100', numberFormat: '@', bold: true }] });
    const written = book.writeRanges({ version: formatted.version, edits: [{ sheet: 'sheet:0', range: 'A100:B100', values: [[123, false]] }] });
    expect(written.items[0].cells.map(cell => cell.value)).toEqual([{ kind: 'number', value: 123 }, { kind: 'bool', value: false }]);
    expect(book.handle.selectionFormatting(0, 'A100').bold).toBe(true);
    expect(book.handle.selectionFormatting(0, 'A100').numberFormatPattern).toBe('@');
    expect(book.writeRanges({ version: written.version, edits: [{ sheet: 'sheet:0', range: 'A100', values: [[123]] }] }).applied).toBe(false);
  } finally { book.close(); }
});

test('XLSX sheet moves preserve structured table formulas', async () => {
  const book = await openXlsx(await tableFixture());
  try {
    const formula = book.writeRanges({ version: book.overview().version, edits: [{ sheet: 'sheet:0', range: 'G100', formulas: [['SUM(Items[Value])']] }] });
    const result = formula.items[0].cells[0].value;
    const moved = book.editSheets({ version: formula.version, edits: [{ action: 'move', sheet: 'sheet:0', index: 2 }] });
    expect(value(book, 'sheet:0', 'G100')[0]).toEqual(result);
    expect(book.outline({ sheet: 'sheet:0' }).items[0].tables[0].name).toBe('Items');
    const reopened = await openXlsx(await book.export());
    try { expect(value(reopened, moved.sheets[2].name, 'G100')[0]).toEqual(result); }
    finally { reopened.close(); }
  } finally { book.close(); }
});

test('XLSX sheet batches expose valid IDs after adding and renaming a sheet in the same batch', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const result = book.editSheets({ version: book.overview().version, edits: [{ action: 'add', name: 'New' }, { action: 'rename', sheet: 'New', name: 'Renamed' }, { action: 'move', sheet: 'Renamed', index: 0 }] });
    expect(result.items[0].sheet).toBe(result.items[1].sheet);
    expect(result.items[0].sheet).toBe(result.sheets[0].sheetId);
    expect(book.readCells({ sheet: result.sheets[0].sheetId, range: 'A1' }).cells[0].value).toEqual({ kind: 'empty' });
    const last = book.editSheets({ version: result.version, edits: [{ action: 'delete', sheet: 'sheet:0' }, { action: 'delete', sheet: 'sheet:1' }, { action: 'delete', sheet: 'sheet:2' }] });
    expect(last.sheets).toHaveLength(1);
    expect(() => book.editSheets({ version: last.version, edits: [{ action: 'delete', sheet: last.sheets[0].sheetId }] })).toThrow('at least one');
  } finally { book.close(); }
});

test('XLSX attachment replays agent edits into a live workbook session', async () => {
  const { createInProcessPair } = await import('../../../shared/office-session/testing/inProcessTransport');
  const { createWorkbookSession, hydratePeer } = await import('../../xlsx/src/session/client');
  const { createWorkbookSessionHost } = await import('../../xlsx/src/session/host');
  const { createWorkbookEditPeer } = await import('../../xlsx/src/session/editPeer');
  const pair = createInProcessPair();
  createWorkbookSessionHost(pair.host);
  const wasm = new Uint8Array(await readFile(resolve(import.meta.dir, '../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm'))).buffer;
  const session = await createWorkbookSession(await xlsxFixture(), { wasm, retainPeerHydration: true }, pair.client);
  const peer = await hydratePeer(session);
  const edits = createWorkbookEditPeer({ session, peer });
  const agent = attachXlsx(peer, { editPeer: edits });
  try {
    const changed = agent.writeRanges({ version: peer.version(), edits: [{ sheet: 'sheet:0', range: 'B3', values: [[345]] }] });
    await edits.flush();
    const worker = await session.call.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }] });
    expect(worker.ok && worker.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 345 });
    expect(await session.call.version()).toBe(changed.version);
    const grid = agent.editGrid({ version: changed.version, edits: [{ sheet: 'sheet:0', axis: 'rows', action: 'insert', at: 3, count: 1 }] });
    await edits.flush();
    const moved = agent.editSheets({ version: grid.version, edits: [{ action: 'move', sheet: 'sheet:0', index: 2 }] });
    await edits.flush();
    expect(await session.call.version()).toBe(moved.version);
    const saved = await openXlsx(await agent.export());
    try { expect(saved.readCells({ sheet: moved.sheets[2].name, range: 'B4' }).cells[0].value).toEqual({ kind: 'number', value: 345 }); }
    finally { saved.close(); }
    agent.close();
    expect(edits.state).toBe('ready');
  } finally { agent.close(); edits.dispose(); peer.dispose(); await session.dispose(); }
}, 10000);

test('XLSX column-width receipts report changes outside the first selected column', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const initial = book.formatRanges({ version: book.overview().version, edits: [{ sheet: 'sheet:0', range: 'B100', columnWidth: 20 }, { sheet: 'sheet:0', range: 'C100', columnWidth: 10 }] });
    const changed = book.formatRanges({ version: initial.version, edits: [{ sheet: 'sheet:0', range: 'B100:C100', columnWidth: 20 }] });
    expect(changed.items[0].changed).toBe(true);
    expect(changed.items[0].columnWidth).toBe(20);
    expect(book.formatRanges({ version: changed.version, edits: [{ sheet: 'sheet:0', range: 'B100:C100', columnWidth: 20 }] }).items[0].changed).toBe(false);
  } finally { book.close(); }
});
