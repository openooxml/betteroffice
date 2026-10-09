import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { xlsxFixture } from './format-fixtures';
import { openXlsx } from '../src/xlsx';

const toolNames = ['xlsx_outline', 'xlsx_read_range', 'xlsx_write_range', 'xlsx_edit_grid', 'xlsx_edit_sheets', 'xlsx_format', 'xlsx_find_replace', 'xlsx_sort', 'xlsx_preview'];

test('XLSX MCP discovers tight schemas and executes every spreadsheet tool through export', async () => {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-mcp-'));
  const client = new Client({ name: 'xlsx-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  let stderr = '';
  transport.stderr?.on('data', chunk => { stderr += String(chunk); });
  try {
    await writeFile(join(root, 'book.xlsx'), await xlsxFixture());
    await client.connect(transport).catch(error => { throw new Error(`${String(error)}\n${stderr}`); });
    const tools = (await client.listTools()).tools;
    for (const name of toolNames) {
      const tool = tools.find(tool => tool.name === name)!;
      expect(tool).toBeDefined();
      expect(tool.inputSchema.additionalProperties).toBe(false);
    }
    const call = async (name: string, args: Record<string, unknown> = {}) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const opened = await call('office_open', { path: 'book.xlsx' });
    const document = opened.document;
    const outline = await call('xlsx_outline', { document, limit: 1 });
    const sheet = outline.items[0].sheetId;
    expect(outline.items[0].dimensions.rows).toBeGreaterThan(0);
    let version = (await call('xlsx_read_range', { document, sheet, range: 'B3:C3' })).version;
    const written = await call('xlsx_write_range', { document, version, edits: [{ sheet, range: 'A100:B102', values: [['c', 3], ['a', 1], ['b', 2]] }, { sheet, range: 'C100', formulas: [['=B100*2']] }] });
    version = written.version;
    expect(written.items[1].cells[0].value).toEqual({ kind: 'number', value: 6 });
    const format = await call('xlsx_format', { document, version, edits: [{ sheet, range: 'B100:B102', bold: true, numberFormat: '0.00', columnWidth: 18 }] });
    version = format.version;
    expect(format.items[0].formatting.bold).toBe(true);
    const sorted = await call('xlsx_sort', { document, version, sheet, range: 'A100:B102', keys: [{ column: 'B', order: 'asc' }] });
    version = sorted.version;
    expect((await call('xlsx_read_range', { document, sheet, range: 'A100' })).cells[0].value).toEqual({ kind: 'text', value: 'a' });
    const found = await call('xlsx_find_replace', { document, query: 'a', sheet, range: 'A100:A102' });
    const replaced = await call('xlsx_find_replace', { document, query: 'a', sheet, range: 'A100:A102', replacement: 'alpha', version: found.version });
    version = replaced.version;
    expect(replaced.replacedCells).toBe(1);
    const grid = await call('xlsx_edit_grid', { document, version, edits: [{ sheet, axis: 'rows', action: 'insert', at: 100, count: 1 }] });
    version = grid.version;
    const sheets = await call('xlsx_edit_sheets', { document, version, edits: [{ action: 'rename', sheet, name: 'Budget' }, { action: 'move', sheet, index: 2 }] });
    version = sheets.version;
    expect(sheets.sheets[2].sheetId).toBe(sheet);
    const preview = await client.callTool({ name: 'xlsx_preview', arguments: { document, sheet, range: 'A100:C103' } });
    expect(preview.isError).not.toBe(true);
    expect((preview.content as Array<{ type: string }>).some(item => item.type === 'image')).toBe(true);
    for (const [name, args] of [
      ['xlsx_write_range', { document, version, edits: [{ sheet, range: 'A100', values: [[1]], formulas: [['1']] }] }],
      ['xlsx_read_range', { document, sheet, range: 'A1', limit: 101 }],
      ['xlsx_preview', { document, sheet, range: 'A1', scale: 4 }],
      ['xlsx_edit_grid', { document, version, edits: [{ sheet, axis: 'rows', action: 'insert', at: 0, count: 1 }] }],
      ['xlsx_format', { document, version, edits: [{ sheet, range: 'A100', fill: 'red' }] }],
      ['xlsx_sort', { document, version, sheet, range: 'A100:B103', keys: [{ column: 'B', order: 'up' }] }],
    ] as const) {
      const invalid = await client.callTool({ name, arguments: args });
      expect(invalid.isError).toBe(true);
      expect(invalid.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    }
    const missing = await client.callTool({ name: 'xlsx_read_range', arguments: { document, sheet: 'missing', range: 'A1' } });
    expect(missing.structuredContent).toMatchObject({ code: 'UNKNOWN_SHEET' });
    expect(JSON.stringify(missing.structuredContent)).toContain('Budget');
    await call('office_export', { document, path: 'result.xlsx' });
    const reopened = await openXlsx(await readFile(join(root, 'result.xlsx')));
    try { expect(reopened.readCells({ sheet: 'Budget', range: 'A101' }).cells[0].value).toEqual({ kind: 'text', value: 'alpha' }); }
    finally { reopened.close(); }
    expect(stderr).toBe('');
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);

test('XLSX MCP read-only mode removes writes and excludes replacement from its schema', async () => {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-mcp-readonly-'));
  const client = new Client({ name: 'xlsx-readonly', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, '--read-only'], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'book.xlsx'), await xlsxFixture());
    await client.connect(transport);
    const tools = (await client.listTools()).tools;
    for (const name of ['xlsx_write_range', 'xlsx_edit_grid', 'xlsx_edit_sheets', 'xlsx_format', 'xlsx_sort']) {
      expect(tools.some(tool => tool.name === name)).toBe(false);
      expect((await client.callTool({ name, arguments: {} })).isError).toBe(true);
    }
    const find = tools.find(tool => tool.name === 'xlsx_find_replace')!;
    expect(find.annotations?.readOnlyHint).toBe(true);
    expect(find.inputSchema.properties).not.toHaveProperty('replacement');
    const opened = await client.callTool({ name: 'office_open', arguments: { path: 'book.xlsx' } });
    const document = (opened.structuredContent as Record<string, unknown>).document;
    expect((await client.callTool({ name: 'xlsx_find_replace', arguments: { document, query: 'Line', replacement: 'X' } })).isError).toBe(true);
    const read = await client.callTool({ name: 'xlsx_read_range', arguments: { document, sheet: 'sheet:0', range: 'B3' } });
    expect(read.structuredContent).toMatchObject({ cells: [{ value: { kind: 'number', value: 100 } }] });
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);
