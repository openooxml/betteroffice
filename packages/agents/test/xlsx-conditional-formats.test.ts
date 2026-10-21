import { expect, test } from 'bun:test';
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
