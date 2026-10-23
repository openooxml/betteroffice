import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { openXlsx } from '../src/xlsx';
import { DocumentToolError } from '../src/types';

async function fixture(table = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data 7" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><autoFilter ref="A3:A30"><filterColumn colId="0"><filters><filter val="x"/></filters></filterColumn></autoFilter>${table ? '<tableParts count="1"><tablePart r:id="t1"/></tableParts>' : ''}</worksheet>`);
  if (table) {
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    zip.file('xl/tables/table1.xml', '<table id="1" name="Items" displayName="Items" ref="E3:F30"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>');
  }
  return zip.generateAsync({ type: 'uint8array' });
}

test('auto filter extension appears in outline and can be removed', async () => {
  const book = await openXlsx(await fixture());
  try {
    const outline = book.outline();
    const sheet = outline.items[0].sheetId;
    expect(outline.items[0].autoFilter).toBe('A3:A30');
    const edited = book.setAutoFilter({ version: outline.version, sheet, range: 'A3:F30' });
    expect(edited.applied).toBe(true);
    expect(book.outline().items[0].autoFilter).toBe('A3:F30');
    book.setAutoFilter({ version: edited.version, sheet, range: null });
    expect(book.outline().items[0].autoFilter).toBeNull();
  } finally { book.close(); }
});

test('table overlap is an actionable document tool error', async () => {
  const book = await openXlsx(await fixture(true));
  try {
    const outline = book.outline();
    expect(() => book.setAutoFilter({ version: outline.version, sheet: outline.items[0].sheetId, range: 'A3:F30' })).toThrow(DocumentToolError);
    expect(book.outline().items[0].autoFilter).toBe('A3:A30');
  } finally { book.close(); }
});

test('MCP tool extends the filter using the latest version', async () => {
  const root = await mkdtemp(join(tmpdir(), 'xlsx-filter-'));
  const client = new Client({ name: 'xlsx-filter-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'book.xlsx'), await fixture());
    await client.connect(transport);
    const tool = (await client.listTools()).tools.find(tool => tool.name === 'xlsx_set_auto_filter')!;
    expect(tool).toBeDefined();
    expect(tool.inputSchema.additionalProperties).toBe(false);
    expect(tool.annotations?.readOnlyHint).toBe(false);
    const call = async (name: string, args: Record<string, unknown>) => {
      const result = await client.callTool({ name, arguments: args });
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
      return result.structuredContent as Record<string, any>;
    };
    const { document } = await call('office_open', { path: 'book.xlsx' });
    const outline = await call('xlsx_outline', { document });
    const sheet = outline.items[0].sheetId;
    expect(outline.items[0].autoFilter).toBe('A3:A30');
    const edited = await call('xlsx_set_auto_filter', { document, version: outline.version, sheet, range: 'A3:F30' });
    expect(edited.applied).toBe(true);
    expect((await call('xlsx_outline', { document })).items[0].autoFilter).toBe('A3:F30');
    const stale = await client.callTool({ name: 'xlsx_set_auto_filter', arguments: { document, version: outline.version, sheet, range: null } });
    expect(stale.isError).toBe(true);
    await call('xlsx_set_auto_filter', { document, version: edited.version, sheet, range: null });
    expect((await call('xlsx_outline', { document })).items[0].autoFilter).toBeNull();
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);
