import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { openWorkbook } from '@betteroffice/xlsx';
import { attachXlsx, openXlsx } from '../src/xlsx';
import { implicitFontWorkbook } from './format-fixtures';

const protection = '<sheetProtection password="CAFE" sheet="1" objects="1" scenarios="1" formatCells="1" sort="1"/>';

async function fixture() {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('</worksheet>', `<mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>${protection}</worksheet>`).replace('<c r="A3"', '<c s="2" r="A3"').replace('<c r="C2" s="1"><v>50</v></c>', '<c r="C2" s="1"><f t="array" ref="C2:C3">1</f><v>1</v></c>'));
  const styles = await zip.file('xl/styles.xml')!.async('string');
  zip.file('xl/styles.xml', styles.replace('<cellXfs count="2">', '<cellXfs count="3">').replace('</cellXfs>', '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf></cellXfs>').replace('applyNumberFormat="1"/>', 'applyNumberFormat="1" applyProtection="1"><protection locked="1" hidden="1"/></xf>'));
  const workbook = await zip.file('xl/workbook.xml')!.async('string');
  zip.file('xl/workbook.xml', workbook.replace('<bookViews>', '<workbookProtection workbookPassword="BEEF" lockStructure="1"/><bookViews>').replace('</sheets>', '<sheet name="Other" sheetId="2" r:id="rId3"/><sheet name="Chart" sheetId="3" r:id="rId4"/></sheets>'));
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', `<Relationship Id="rId3" Type="${ns}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId4" Type="${ns}/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>`));
  zip.file('xl/worksheets/sheet2.xml', sheet);
  zip.file('xl/chartsheets/sheet1.xml', '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>');
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/></Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}

test('XLSX tools explicitly write through protection and preserve it through save and reopen', async () => {
  const root = await mkdtemp(join(tmpdir(), 'protected-writes-'));
  const client = new Client({ name: 'protected-writes', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  try {
    const source = await fixture();
    await writeFile(join(root, 'book.xlsx'), source);
    await client.connect(transport);
    const call = async (name: string, args: Record<string, unknown>) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    for (const tool of (await client.listTools()).tools.filter(tool => ['xlsx_write_range', 'xlsx_format', 'xlsx_sort', 'xlsx_find_replace'].includes(tool.name))) {
      expect(tool.inputSchema.properties).toHaveProperty('protectedSheets');
      expect(tool.description).toContain('only when the user explicitly asks');
      expect(tool.description).toContain('Never set protectedSheets:"allow" preemptively');
    }
    const { document } = await call('office_open', { path: 'book.xlsx' });
    let { version } = await call('xlsx_outline', { document });
    const edit = { sheet: 'Sales', range: 'A2', values: [['Updated']] };
    for (const protectedSheets of [undefined, 'refuse']) {
      const refusal = await client.callTool({ name: 'xlsx_write_range', arguments: { document, version, edits: [{ sheet: 'Sales', range: 'A3', values: [['Must stay unchanged']] }, edit], ...(protectedSheets ? { protectedSheets } : {}) } });
      expect(refusal.isError).toBe(true);
      expect(refusal.structuredContent).toMatchObject({ code: 'locked-target' });
      expect(JSON.stringify(refusal)).toContain('protectedSheets:');
      expect((await call('xlsx_outline', { document })).version).toBe(version);
    }
    await call('office_export', { document, path: 'unchanged.xlsx' });
    const original = await JSZip.loadAsync(source);
    const unchanged = await JSZip.loadAsync(await readFile(join(root, 'unchanged.xlsx')));
    for (const [path, part] of Object.entries(original.files)) if (!part.dir) expect(await unchanged.file(path)!.async('uint8array')).toEqual(await part.async('uint8array'));
    for (const tool of ['xlsx_format', 'xlsx_sort']) {
      const args = tool === 'xlsx_format' ? { edits: [{ sheet: 'Sales', range: 'B2', bold: true }] } : { sheet: 'Sales', range: 'A2:B4', keys: [{ column: 'B', order: 'asc' }] };
      const refused = await client.callTool({ name: tool, arguments: { document, version, ...args } });
      expect(refused.structuredContent).toMatchObject({ code: 'LOCKED_SHEET' });
      expect(JSON.stringify(refused)).toContain('protectedSheets:');
    }
    const invalid = await client.callTool({ name: 'xlsx_write_range', arguments: { document, version, edits: [edit], protectedSheets: 'always' } });
    expect(invalid.isError).toBe(true);
    expect(invalid.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(JSON.stringify(invalid)).toContain('protectedSheets');
    const written = await call('xlsx_write_range', { document, version, edits: [edit], protectedSheets: 'allow' });
    version = written.version;
    expect(written.protectedCells).toEqual(['Sales!A2']);
    expect(written.items[0].cells[0].value).toEqual({ kind: 'text', value: 'Updated' });
    const formatted = await call('xlsx_format', { document, version, edits: [{ sheet: 'Sales', range: 'B2', bold: true, numberFormat: '0.00' }], protectedSheets: 'allow' });
    version = formatted.version;
    expect(formatted.protectedCells).toContain('Sales!B2');
    expect(formatted.items[0].formatting).toMatchObject({ bold: true, numberFormatPattern: '0.00' });
    const sorted = await call('xlsx_sort', { document, version, sheet: 'Sales', range: 'A2:B4', keys: [{ column: 'B', order: 'asc' }], protectedSheets: 'allow' });
    version = sorted.version;
    expect(sorted.protectedCells).toContain('Sales!A2');
    expect((await call('xlsx_read_range', { document, sheet: 'Sales', range: 'A2' })).cells[0].value).toEqual({ kind: 'text', value: 'Bottle' });
    const replaced = await call('xlsx_find_replace', { document, version, sheet: 'Sales', query: 'Bottle', replacement: 'Final', protectedSheets: 'allow' });
    version = replaced.version;
    expect(replaced.protectedCells).toHaveLength(1);
    for (const [sheet, range, code] of [['Sales', 'E1', 'locked-target'], ['Sales', 'C2', 'locked-target'], ['Chart', 'A1', 'unsupported']]) {
      const refused = await client.callTool({ name: 'xlsx_write_range', arguments: { document, version, edits: [{ sheet, range, values: [[7]] }], protectedSheets: 'allow' } });
      expect(refused.structuredContent).toMatchObject({ code });
      expect((await call('xlsx_outline', { document })).version).toBe(version);
    }
    const mergedFormat = await client.callTool({ name: 'xlsx_format', arguments: { document, version, edits: [{ sheet: 'Sales', range: 'E1', bold: true }], protectedSheets: 'allow' } });
    expect(mergedFormat.structuredContent).toMatchObject({ code: 'locked-target' });
    const mergedSort = await client.callTool({ name: 'xlsx_sort', arguments: { document, version, sheet: 'Sales', range: 'D1:E2', keys: [{ column: 'D', order: 'asc' }], protectedSheets: 'allow' } });
    expect(mergedSort.structuredContent).toMatchObject({ code: 'MERGED_SORT' });
    await call('office_export', { document, path: 'saved.xlsx' });
    const saved = await JSZip.loadAsync(await readFile(join(root, 'saved.xlsx')));
    const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml.match(/<sheetProtection\b[^>]*\/>/)?.[0]).toBe(protection);
    expect(await saved.file('xl/worksheets/sheet2.xml')!.async('string')).toBe(await original.file('xl/worksheets/sheet2.xml')!.async('string'));
    expect((await saved.file('xl/workbook.xml')!.async('string')).match(/<workbookProtection\b[^>]*\/>/)?.[0]).toBe('<workbookProtection workbookPassword="BEEF" lockStructure="1"/>');
    const styles = await saved.file('xl/styles.xml')!.async('string');
    const originalStyles = await original.file('xl/styles.xml')!.async('string');
    const xfs = (text: string) => text.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/)![1].match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!;
    const beforeXfs = xfs(originalStyles);
    const afterXfs = xfs(styles);
    expect(afterXfs.slice(0, beforeXfs.length)).toEqual(beforeXfs);
    const originalSheet = await original.file('xl/worksheets/sheet1.xml')!.async('string');
    const cellProtection = (sheet: string, formats: string[], address: string) => {
      const cell = sheet.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))?.[0];
      const style = Number(cell?.match(/\bs="(\d+)"/)?.[1] ?? 0);
      return formats[style].match(/<protection\b[^>]*\/>/)?.[0] ?? null;
    };
    for (const address of ['A2', 'A3', 'A4', 'B2', 'B3', 'B4']) expect(cellProtection(xml, afterXfs, address)).toBe(cellProtection(originalSheet, beforeXfs, address));
    const reopened = await call('office_open', { path: 'saved.xlsx' });
    const outline = await call('xlsx_outline', { document: reopened.document });
    expect(outline.items[0].editable).toBe(false);
    const refused = await client.callTool({ name: 'xlsx_write_range', arguments: { document: reopened.document, version: outline.version, edits: [edit] } });
    expect(refused.structuredContent).toMatchObject({ code: 'locked-target' });
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);

test('XLSX formatting preserves cell protection on an unprotected sheet without opt-in', async () => {
  const zip = await JSZip.loadAsync(await fixture());
  const path = 'xl/worksheets/sheet1.xml';
  zip.file(path, (await zip.file(path)!.async('string')).replace(protection, ''));
  const book = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const result = book.formatRanges({ version: book.handle.version(), edits: [{ sheet: 'Sales', range: 'A3:B3', bold: true, numberFormat: '0.00' }] });
    expect(result.applied).toBe(true);
    const saved = await JSZip.loadAsync(await book.export());
    const sheet = await saved.file(path)!.async('string');
    expect(sheet).not.toContain('<sheetProtection');
    const styles = await saved.file('xl/styles.xml')!.async('string');
    const xfs = styles.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/)![1].match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!;
    for (const [address, locked] of [['A3', '0'], ['B3', '1']]) {
      const cell = sheet.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))![0];
      const style = Number(cell.match(/\bs="(\d+)"/)![1]);
      expect(xfs[style]).toContain(`applyProtection="1"`);
      expect(xfs[style]).toContain(`<protection locked="${locked}" hidden="1"/>`);
    }
  } finally { book.close(); }
});

test('XLSX protected writes remain opt-in and unavailable on attached editor handles', async () => {
  const source = await fixture();
  const owned = await openXlsx(source);
  try {
    const version = owned.handle.version();
    const unchanged = owned.writeRanges({ version, protectedSheets: 'allow', edits: [{ sheet: 'Sales', range: 'A2', values: [['Cap']] }] });
    expect(unchanged.applied).toBe(false);
    expect(unchanged.protectedCells).toEqual([]);
  } finally { owned.close(); }
  const handle = openWorkbook(source);
  const attached = attachXlsx(handle);
  try {
    expect(() => attached.writeRanges({ version: handle.version(), protectedSheets: 'allow', edits: [{ sheet: 'Sales', range: 'A2', values: [['Changed']] }] })).toThrow('standalone');
    expect(() => attached.writeRanges({ version: handle.version(), edits: [{ sheet: 'Sales', range: 'A2', values: [['Changed']] }] })).toThrow('locked');
  } finally { attached.close(); handle.dispose(); }
});

test('XLSX read-only server refuses protected writes and omits the protection option', async () => {
  const root = await mkdtemp(join(tmpdir(), 'protected-readonly-'));
  const client = new Client({ name: 'protected-readonly', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, '--read-only'], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'book.xlsx'), await fixture());
    await client.connect(transport);
    const tools = (await client.listTools()).tools;
    const opened = await client.callTool({ name: 'office_open', arguments: { path: 'book.xlsx' } });
    const document = (opened.structuredContent as Record<string, unknown>).document;
    for (const name of ['xlsx_write_range', 'xlsx_format', 'xlsx_sort']) {
      expect(tools.some(tool => tool.name === name)).toBe(false);
      expect((await client.callTool({ name, arguments: { document, protectedSheets: 'allow' } })).isError).toBe(true);
    }
    const find = tools.find(tool => tool.name === 'xlsx_find_replace')!;
    expect(find.inputSchema.properties).not.toHaveProperty('protectedSheets');
    expect(find.inputSchema.properties).not.toHaveProperty('replacement');
    expect((await client.callTool({ name: 'xlsx_find_replace', arguments: { document, query: 'Cap', replacement: 'Changed', protectedSheets: 'allow' } })).isError).toBe(true);
    const read = await client.callTool({ name: 'xlsx_read_range', arguments: { document, sheet: 'Sales', range: 'A2' } });
    expect(read.structuredContent).toMatchObject({ cells: [{ value: { kind: 'text', value: 'Cap' } }] });
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);

test('XLSX sort through protection keeps source style XML and reports moved formats on unlocked cells', async () => {
  const extension = '<extLst><ext uri="{6E1A2B3C-4D5E-4F60-8A9B-0C1D2E3F4A5B}"><x:keep xmlns:x="urn:protected-sort" value="kept"/></ext></extLst>';
  const zip = await JSZip.loadAsync(await fixture());
  const styles = await zip.file('xl/styles.xml')!.async('string');
  zip.file('xl/styles.xml', styles.replace('<cellXfs count="3">', '<cellXfs count="4">').replace('</cellXfs>', `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0">${extension}</xf></cellXfs>`));
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('<c r="A2" t="inlineStr">', '<c r="A2" s="3" t="inlineStr">'));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const book = await openXlsx(source);
  try {
    const sorted = book.sortRange({ version: book.handle.version(), sheet: 'Sales', range: 'A2:B4', keys: [{ column: 'B', order: 'asc' }], protectedSheets: 'allow' });
    expect(sorted.applied).toBe(true);
    expect(sorted.protectedCells).toContain('Sales!A3');
    const saved = await JSZip.loadAsync(await book.export());
    const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('Cap');
    const xfs = (text: string) => text.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/)![1].match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!;
    const before = xfs(await (await JSZip.loadAsync(source)).file('xl/styles.xml')!.async('string'));
    const after = xfs(await saved.file('xl/styles.xml')!.async('string'));
    expect(after.slice(0, before.length)).toEqual(before);
    const cell = xml.match(/<c\b(?=[^>]*\br="A3")[^>]*>/)![0];
    const moved = after[Number(cell.match(/\bs="(\d+)"/)![1])];
    expect(moved).toContain(extension);
    expect(moved).toContain('<protection locked="0" hidden="1"/>');
  } finally { book.close(); }
});
