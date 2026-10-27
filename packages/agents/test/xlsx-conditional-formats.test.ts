import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { xlsxFixture } from './format-fixtures';
import { openXlsx } from '../src/xlsx';

test('xlsx conditional formats add list delete export reopen via MCP', async () => {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-cf-'));
  const client = new Client({ name: 'xlsx-cf-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'book.xlsx'), await xlsxFixture());
    await client.connect(transport);
    const tools = (await client.listTools()).tools;
    const listTool = tools.find(tool => tool.name === 'xlsx_list_conditional_formats')!;
    expect(listTool).toBeDefined();
    expect(listTool.inputSchema.additionalProperties).toBe(false);
    expect(listTool.annotations?.readOnlyHint).toBe(true);
    expect(tools.find(tool => tool.name === 'xlsx_read_conditional_format_formula')?.annotations?.readOnlyHint).toBe(true);
    const editTool = tools.find(tool => tool.name === 'xlsx_edit_conditional_formats')!;
    expect(editTool).toBeDefined();
    expect(editTool.inputSchema.additionalProperties).toBe(false);
    const call = async (name: string, args: Record<string, unknown> = {}) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const opened = await call('office_open', { path: 'book.xlsx' });
    const document = opened.document;
    const outline = await call('xlsx_outline', { document, limit: 1 });
    const sheet = outline.items[0].sheetId;
    let version = outline.version;
    const empty = await call('xlsx_list_conditional_formats', { document, sheet });
    expect(empty).toMatchObject({ total: 0, items: [], nextOffset: null });
    const added = await call('xlsx_edit_conditional_formats', { document, version, edits: [
      { action: 'add_conditional_format', sheet, range: 'A1:A5', rule: { type: 'cellIs', operator: 'greaterThan', formulas: ['5'] }, format: { fontColor: '#FF0000', bold: true, fill: '#00FF00' } },
      { action: 'add_conditional_format', sheet, range: 'B1:B5', rule: { type: 'expression', formulas: ['=A1>5'] }, format: { bold: false } },
    ] });
    version = added.version;
    expect(added.items).toHaveLength(2);
    const listed = await call('xlsx_list_conditional_formats', { document, sheet });
    expect(listed.total).toBe(2);
    expect(listed.items[0]).toMatchObject({ range: 'A1:A5', type: 'cellIs', operator: 'greaterThan', formulas: ['5'], editable: true, format: { fontColor: '#FF0000', bold: true, fill: '#00FF00' } });
    const formula = await call('xlsx_read_conditional_format_formula', { document, sheet, id: listed.items[1].id, length: 2 });
    expect(formula).toMatchObject({ text: 'A1', nextStart: 2, length: 4 });
    const page = await call('xlsx_list_conditional_formats', { document, sheet, offset: 1, limit: 1 });
    expect(page.items).toHaveLength(1);
    expect(page.nextOffset).toBe(null);
    const stale = await client.callTool({ name: 'xlsx_edit_conditional_formats', arguments: { document, version: 'stale', edits: [{ action: 'delete_conditional_format', sheet, id: listed.items[0].id }] } });
    expect(stale.isError).toBe(true);
    const rolledBack = await client.callTool({ name: 'xlsx_edit_conditional_formats', arguments: { document, version, edits: [
      { action: 'add_conditional_format', sheet, range: 'C1:C2', rule: { type: 'cellIs', operator: 'between', formulas: ['1'] }, format: {} },
    ] } });
    expect(rolledBack.isError).toBe(true);
    expect((await call('xlsx_list_conditional_formats', { document, sheet })).total).toBe(2);
    const badColor = await client.callTool({ name: 'xlsx_edit_conditional_formats', arguments: { document, version, edits: [
      { action: 'add_conditional_format', sheet, range: 'C1:C2', rule: { type: 'cellIs', operator: 'equal', formulas: ['1'] }, format: { fontColor: 'red' } },
    ] } });
    expect(badColor.isError).toBe(true);
    const deleted = await call('xlsx_edit_conditional_formats', { document, version, edits: [{ action: 'delete_conditional_format', sheet, id: listed.items[0].id }] });
    version = deleted.version;
    expect((await call('xlsx_list_conditional_formats', { document, sheet })).total).toBe(1);
    await call('office_export', { document, path: 'result.xlsx' });
    const reopened = await openXlsx(await readFile(join(root, 'result.xlsx')));
    try {
      const again = reopened.listConditionalFormats({ sheet, limit: 10 });
      expect(again.total).toBe(1);
      expect(again.items[0]).toMatchObject({ type: 'expression', range: 'B1:B5' });
    } finally { reopened.close(); }
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);

test('conditional format result IDs are claimed per sheet', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const outline = book.outline();
    const added = book.editSheets({ version: outline.version, edits: [{ action: 'add', name: 'CF target' }] });
    const sheets = [outline.items[0].sheetId, added.sheets.find(sheet => sheet.name === 'CF target')!.sheetId];
    const result = book.editConditionalFormats({ version: added.version, edits: [
      { action: 'add_conditional_format', sheet: sheets[0], range: 'A1:A5', rule: { type: 'expression', formulas: ['A1>0'] }, format: {} },
      { action: 'add_conditional_format', sheet: sheets[1], range: 'A1:A5', rule: { type: 'expression', formulas: ['A1>0'] }, format: {} },
      { action: 'add_conditional_format', sheet: sheets[0], range: 'A1:A5', rule: { type: 'expression', formulas: ['A1>5'] }, format: {} },
    ] });
    expect(result.items.map(item => item.id)).toEqual(['cf1', 'cf1', 'cf2']);
    expect(result.items.map(item => item.sheet)).toEqual([sheets[0], sheets[1], sheets[0]]);
  } finally { book.close(); }
});


test('replacement IDs follow each operation including reused and deleted IDs', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const outline = book.outline();
    const sheet = outline.items[0].sheetId;
    const add = (formula: string) => ({ action: 'add_conditional_format' as const, sheet, range: 'A1:A5', rule: { type: 'expression' as const, formulas: [formula] }, format: {} });
    const initial = book.editConditionalFormats({ version: outline.version, edits: [add('A1>0'), add('A1>1')] });
    const result = book.editConditionalFormats({ version: initial.version, edits: [
      { action: 'delete_conditional_format', sheet, id: 'cf2' },
      add(' =A1>2 '),
      add('A1>3'),
      { action: 'delete_conditional_format', sheet, id: 'cf3' },
      add('A1>4'),
    ] });
    expect(result.items.map(item => item.id)).toEqual(['cf2', 'cf2', 'cf3', 'cf3', 'cf3']);
    expect(result.items[1]).toMatchObject({ formulas: ['A1>2'] });
    expect(result.items[2]).toMatchObject({ formulas: ['A1>3'] });
    expect(result.items[4]).toMatchObject({ formulas: ['A1>4'] });
    expect(book.listConditionalFormats({ sheet }).items.map(item => [item.id, item.formulas])).toEqual([
      ['cf1', ['A1>0']], ['cf2', ['A1>2']], ['cf3', ['A1>4']],
    ]);
  } finally { book.close(); }
});


test('conditional format pages bound text and expose complete formula windows', async () => {
  const book = await openXlsx(await xlsxFixture());
  try {
    const sheet = book.outline().items[0].sheetId;
    const formula = `"${'x'.repeat(8188)}"`;
    book.handle.applyOps(Array.from({ length: 100 }, (_, index) => ({ type: 'addConditionalFormat', sheet: 0, range: { start: { row: index, col: 0 }, end: { row: index, col: 0 } }, kind: 'cellIs', operator: 'between', formulas: [formula, formula], format: {} })));
    const seen: string[] = [];
    let offset = 0;
    while (true) {
      const page = book.listConditionalFormats({ sheet, offset, limit: 100 });
      expect(JSON.stringify(page).length).toBeLessThanOrEqual(16000);
      expect(page.items.length).toBeGreaterThan(0);
      expect(page.items.every(item => item.truncated && item.formulas.every(text => text.length <= 400))).toBe(true);
      seen.push(...page.items.map(item => item.id));
      if (page.nextOffset === null) break;
      expect(page.nextOffset).toBeGreaterThan(offset);
      offset = page.nextOffset;
    }
    expect(seen).toHaveLength(100);
    expect(new Set(seen).size).toBe(100);
    for (const index of [0, 1]) {
      let text = '';
      let start = 0;
      while (true) {
        const window = book.readConditionalFormatFormula({ sheet, id: seen[0], formula: index, start });
        expect(JSON.stringify(window).length).toBeLessThanOrEqual(16000);
        text += window.text;
        if (window.nextStart === null) break;
        start = window.nextStart;
      }
      expect(text).toBe(formula);
    }
    expect(() => book.readConditionalFormatFormula({ sheet, id: seen[0], formula: 2 })).toThrow();
    expect(() => book.readConditionalFormatFormula({ sheet, id: 'missing' })).toThrow();
  } finally { book.close(); }
});

test('xlsx conditional formats preserve italic true and false through MCP export', async () => {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-cf-italic-'));
  const client = new Client({ name: 'xlsx-cf-italic-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'book.xlsx'), await xlsxFixture());
    await client.connect(transport);
    const call = async (name: string, args: Record<string, unknown> = {}) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const { document } = await call('office_open', { path: 'book.xlsx' });
    const outline = await call('xlsx_outline', { document, limit: 1 });
    const sheet = outline.items[0].sheetId;
    await call('xlsx_edit_conditional_formats', { document, version: outline.version, edits: [
      { action: 'add_conditional_format', sheet, range: 'F12:F18', rule: { type: 'cellIs', operator: 'greaterThan', formulas: ['0.5'] }, format: { fill: '#FFC7CE', fontColor: '#9C0006', bold: true, italic: true } },
      { action: 'add_conditional_format', sheet, range: 'G12:G18', rule: { type: 'expression', formulas: ['F12>0.5'] }, format: { italic: false } },
    ] });
    const listed = await call('xlsx_list_conditional_formats', { document, sheet });
    expect(listed.items.map((item: any) => item.format)).toEqual([
      { fill: '#FFC7CE', fontColor: '#9C0006', bold: true, italic: true },
      { italic: false },
    ]);
    await call('office_export', { document, path: 'italic.xlsx' });
    const bytes = await readFile(join(root, 'italic.xlsx'));
    const zip = await JSZip.loadAsync(bytes);
    const styles = await zip.file('xl/styles.xml')!.async('string');
    expect(styles).toContain('<font><b/><i/><color rgb="FF9C0006"/></font>');
    expect(styles).toContain('<font><i val="0"/></font>');
    const reopened = await openXlsx(bytes);
    try {
      expect(reopened.listConditionalFormats({ sheet }).items.map(item => item.format.italic)).toEqual([true, false]);
    } finally { reopened.close(); }
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);
