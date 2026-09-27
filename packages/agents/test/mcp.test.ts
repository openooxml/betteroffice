import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { fixture } from './fixture';

test('stdio MCP supports discovery through verified export and confines file access', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-'));
  const outside = await mkdtemp(join(tmpdir(), 'betteroffice-outside-'));
  const client = new Client({ name: 'test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  let stderr = '';
  transport.stderr?.on('data', chunk => { stderr += String(chunk); });
  try {
    await writeFile(join(root, 'report.docx'), await fixture(200));
    await writeFile(join(outside, 'secret.docx'), await fixture());
    await symlink(outside, join(root, 'escape'));
    await client.connect(transport);
    const tools = (await client.listTools()).tools;
    expect(tools.some(tool => tool.name === 'office_grep')).toBe(true);
    const call = async (name: string, args: Record<string, unknown> = {}) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const files = await call('office_files');
    expect(files.files).toContainEqual({ path: 'report.docx', kind: 'docx' });
    const opened = await call('office_open', { path: 'report.docx' });
    const document = opened.document;
    const hits = await call('office_grep', { document, query: '€4.2 million' });
    const hit = hits.matches[0];
    await call('office_read', { document, ref: hit.ref });
    const staged = await call('office_propose', { document, author: 'MCP test', edits: [{ match: hit.match, newText: '€5.1 million' }] });
    const proposal = staged.id;
    expect((await call('office_review', { document, proposal })).status).toBe('pending');
    const rendered = await client.callTool({ name: 'office_render', arguments: { document, proposal, page: 1 } });
    expect(rendered.isError).not.toBe(true);
    expect((rendered.content as Array<{ type: string }>).some(item => item.type === 'image')).toBe(true);
    expect((await call('office_verify', { document, proposal })).reopened).toBe(true);
    await call('office_export', { document, proposal, path: 'result.docx' });
    expect((await readFile(join(root, 'result.docx'))).length).toBeGreaterThan(0);
    expect((await call('office_read', { document, ref: hit.ref })).text).toContain('€4.2 million');
    expect((await client.callTool({ name: 'office_export', arguments: { document, path: 'report.docx' } })).isError).toBe(true);
    expect((await client.callTool({ name: 'office_open', arguments: { path: join(outside, 'secret.docx') } })).isError).toBe(true);
    expect((await client.callTool({ name: 'office_open', arguments: { path: 'escape/secret.docx' } })).isError).toBe(true);
    expect((await client.callTool({ name: 'office_export', arguments: { document, path: 'escape/leak.docx' } })).isError).toBe(true);
    expect((await client.callTool({ name: 'office_grep', arguments: { document, query: 'risk', limit: 100000 } })).isError).toBe(true);
    await call('office_accept', { document, proposal });
    expect((await call('office_read', { document, ref: hit.ref })).text).toContain('€5.1 million');
    await call('office_close', { document });
    expect(stderr).toBe('');
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
    await rm(outside, { recursive: true, force: true });
  }
}, 30000);

test('stdio MCP edits XLSX and PPTX prototypes with discoverable tools and same-format exports', async () => {
  const { xlsxFixture, pptxFixture } = await import('./format-fixtures');
  const { openXlsx, openPptx } = await import('../src');
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-formats-mcp-'));
  const client = new Client({ name: 'formats-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  let stderr = '';
  transport.stderr?.on('data', chunk => { stderr += String(chunk); });
  try {
    await writeFile(join(root, 'budget.xlsx'), await xlsxFixture());
    await writeFile(join(root, 'slides.pptx'), await pptxFixture());
    await client.connect(transport);
    const call = async (name: string, args: Record<string, unknown> = {}) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    expect((await call('office_files')).files).toEqual([{ path: 'budget.xlsx', kind: 'xlsx' }, { path: 'slides.pptx', kind: 'pptx' }]);
    const workbook = await call('office_open', { path: 'budget.xlsx' });
    const document = workbook.document;
    const sheet = (await call('office_outline', { document })).items[0].sheetId;
    const cells = (await call('office_cells', { document, sheet, range: 'B3:E3' })).cells;
    const proposal = (await call('office_propose_cells', { document, author: 'MCP', edits: [{ cell: cells[0].cell, input: '250' }, { cell: cells[3].cell, input: '=D3*2' }] })).id;
    expect((await call('office_review', { document, proposal })).stale).toBe(false);
    expect((await call('office_verify', { document, proposal })).reopened).toBe(true);
    await call('office_export', { document, proposal, path: 'budget-revised.xlsx' });
    expect((await client.callTool({ name: 'office_export', arguments: { document, path: 'wrong.pptx' } })).isError).toBe(true);
    expect((await call('office_read', { document, ref: cells[0].ref })).value.value).toBe(100);
    await call('office_accept', { document, proposal });
    expect((await call('office_read', { document, ref: cells[3].ref, field: 'formula' })).text).toBe('D3*2');
    const saved = await openXlsx(await readFile(join(root, 'budget-revised.xlsx')));
    try { expect(saved.readCells({ sheet, range: 'E3' }).cells[0].value).toEqual({ kind: 'number', value: 614 }); }
    finally { saved.close(); }
    const deck = await call('office_open', { path: 'slides.pptx' });
    const hits = await call('office_grep', { document: deck.document, query: '€4.2 million' });
    const staged = await call('office_propose', { document: deck.document, author: 'MCP', edits: [{ match: hits.matches[0].match, newText: '€5.1 million' }] });
    await call('office_verify', { document: deck.document, proposal: staged.id });
    await call('office_export', { document: deck.document, proposal: staged.id, path: 'slides-revised.pptx' });
    const slides = await openPptx(await readFile(join(root, 'slides-revised.pptx')));
    try { expect(slides.grep({ query: '€5.1 million' }).matches).toHaveLength(1); }
    finally { slides.close(); }
    expect((await client.callTool({ name: 'office_render', arguments: { document: deck.document } })).isError).toBe(true);
    expect(stderr).toBe('');
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('read-only MCP omits both text and cell proposal tools', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-readonly-'));
  const client = new Client({ name: 'readonly-test', version: '1' });
  try {
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, '--read-only'], stderr: 'pipe' }));
    const tools = (await client.listTools()).tools.map(tool => tool.name);
    expect(tools).toContain('office_cells');
    for (const name of ['office_propose', 'office_propose_cells', 'office_accept', 'office_export']) expect(tools).not.toContain(name);
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);
