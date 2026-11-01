import { expect, test } from 'bun:test';
import { spawn } from 'node:child_process';
import { mkdtemp, readFile, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { fixture } from './fixture';

test('stdio closes when unterminated input exceeds 1 MiB', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-limit-'));
  const child = spawn(process.execPath, [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], { stdio: ['pipe', 'pipe', 'pipe'] });
  child.stdin.on('error', () => {});
  try {
    const exited = new Promise<number | null>((resolveExit, reject) => {
      child.once('exit', resolveExit);
      child.once('error', reject);
    });
    child.stdin.write('x'.repeat(1024 * 1024 + 1));
    expect(await exited).toBe(0);
    expect(await readdir(root)).toEqual([]);
  } finally {
    child.kill();
    await rm(root, { recursive: true, force: true });
  }
}, 10000);

test('stdio MCP initialize reports the package version', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-version-'));
  const client = new Client({ name: 'test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  try {
    await client.connect(transport);
    const { version } = JSON.parse(await readFile(resolve(import.meta.dir, '../package.json'), 'utf8'));
    expect(version).not.toBe('0.0.0');
    expect(client.getServerVersion()).toEqual({ name: 'betteroffice', version });
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

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
    expect((await call('office_grep', { document, query: 'risk', limit: 100000 })).clamping).toContain('limit clamped to 100; follow pagination cursors for remaining content.');
    for (const [name, args] of [
      ['office_open', { path: 'x'.repeat(4097) }],
      ['office_read', { document: 'x'.repeat(129), ref: hit.ref }],
      ['office_propose', { document, author: 'test', edits: [{ match: hit.match, newText: 'x'.repeat(16001) }] }],
    ] as const) {
      const invalid = await client.callTool({ name, arguments: args });
      expect(invalid.isError).toBe(true);
      expect(invalid.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    }
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

const DEFAULT_TOOLS = [
  'xlsx_outline', 'xlsx_list_defined_names', 'xlsx_read_range', 'xlsx_list_charts', 'xlsx_read_chart_text',
  'xlsx_list_conditional_formats', 'xlsx_read_conditional_format_formula', 'xlsx_preview', 'xlsx_find_replace',
  'xlsx_write_range', 'xlsx_edit_grid', 'xlsx_edit_sheets', 'xlsx_edit_defined_names', 'xlsx_format', 'xlsx_sort',
  'xlsx_set_auto_filter', 'xlsx_edit_charts', 'xlsx_edit_conditional_formats',
  'office_files', 'office_open', 'office_outline', 'office_grep', 'office_read', 'office_cells', 'office_render',
  'office_verify', 'office_close', 'office_propose', 'office_propose_cells', 'office_review', 'office_accept',
  'office_reject', 'docx_edit', 'office_export',
  'pptx_outline', 'pptx_read_slide', 'pptx_preview', 'pptx_edit',
];

async function connectFiltered(root: string, extra: string[], env?: Record<string, string>) {
  const client = new Client({ name: 'filter-test', version: '1' });
  const merged = env ? Object.fromEntries(Object.entries(process.env).filter((entry): entry is [string, string] => entry[1] !== undefined)) : undefined;
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, ...extra], stderr: 'pipe', ...(env ? { env: { ...merged, ...env } } : {}) });
  await client.connect(transport);
  return client;
}

test('default stdio tool list equals the unfiltered registration order', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-default-'));
  const client = await connectFiltered(root, []);
  try {
    expect((await client.listTools()).tools.map(tool => tool.name)).toEqual(DEFAULT_TOOLS);
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('--formats docx filters tools, calls, file listing and office_open', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-docx-'));
  await writeFile(join(root, 'report.docx'), await fixture());
  await writeFile(join(root, 'budget.xlsx'), 'not a workbook');
  await writeFile(join(root, 'slides.pptx'), 'not a deck');
  const client = await connectFiltered(root, ['--formats', 'docx']);
  try {
    const names = (await client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('docx_edit');
    expect(names).toContain('office_grep');
    expect(names).toContain('office_render');
    for (const name of ['xlsx_outline', 'pptx_outline', 'pptx_edit', 'office_cells', 'office_propose_cells']) expect(names).not.toContain(name);
    const instructions = (client as unknown as { getInstructions?: () => string | undefined }).getInstructions?.() ?? '';
    expect(instructions).not.toContain('XLSX:');
    expect(instructions).not.toContain('PPTX:');
    const filtered = await client.callTool({ name: 'xlsx_outline', arguments: { document: 'doc1' } });
    expect(filtered.isError).toBe(true);
    expect(filtered.structuredContent).toMatchObject({ code: 'UNKNOWN_TOOL' });
    const files = await client.callTool({ name: 'office_files', arguments: {} });
    expect(files.isError).not.toBe(true);
    expect((files.structuredContent as Record<string, any>).files).toEqual([{ path: 'report.docx', kind: 'docx' }]);
    const refused = await client.callTool({ name: 'office_open', arguments: { path: 'budget.xlsx' } });
    expect(refused.isError).toBe(true);
    expect(refused.structuredContent).toMatchObject({ code: 'UNSUPPORTED_FORMAT' });
    expect(JSON.stringify(refused.structuredContent)).toContain('docx');
    expect((await client.callTool({ name: 'office_open', arguments: { path: 'report.docx' } })).isError).not.toBe(true);
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('--formats xlsx,pptx keeps those formats and shared tools only', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-xp-'));
  const client = await connectFiltered(root, ['--formats', 'xlsx,pptx']);
  try {
    const names = (await client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('xlsx_outline');
    expect(names).toContain('pptx_outline');
    expect(names).toContain('office_cells');
    expect(names).toContain('office_propose_cells');
    expect(names).toContain('office_propose');
    expect(names).toContain('office_grep');
    expect(names).not.toContain('docx_edit');
    expect(names).not.toContain('office_render');
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('--tools, combined intersection, env vars and flag precedence over stdio', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-tools-'));
  try {
    const only = await connectFiltered(root, ['--tools', 'office_files,office_open']);
    try { expect((await only.listTools()).tools.map(tool => tool.name)).toEqual(['office_files', 'office_open']); }
    finally { await only.close(); }
    const both = await connectFiltered(root, ['--formats', 'docx', '--tools', 'office_files,xlsx_outline,docx_edit']);
    try { expect((await both.listTools()).tools.map(tool => tool.name)).toEqual(['office_files', 'docx_edit']); }
    finally { await both.close(); }
    const fromEnv = await connectFiltered(root, [], { BETTEROFFICE_MCP_FORMATS: 'xlsx', BETTEROFFICE_MCP_TOOLS: 'office_files,xlsx_outline,pptx_outline' });
    try { expect((await fromEnv.listTools()).tools.map(tool => tool.name)).toEqual(['xlsx_outline', 'office_files']); }
    finally { await fromEnv.close(); }
    const flagWins = await connectFiltered(root, ['--formats', 'docx'], { BETTEROFFICE_MCP_FORMATS: 'xlsx' });
    try {
      const names = (await flagWins.listTools()).tools.map(tool => tool.name);
      expect(names).toContain('docx_edit');
      expect(names).not.toContain('xlsx_outline');
    } finally { await flagWins.close(); }
    const readOnlyBoth = await connectFiltered(root, ['--read-only', '--formats', 'docx']);
    try {
      const names = (await readOnlyBoth.listTools()).tools.map(tool => tool.name);
      expect(names).toContain('office_grep');
      expect(names).not.toContain('docx_edit');
      expect(names).not.toContain('office_propose');
    } finally { await readOnlyBoth.close(); }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}, 60000);

test('unknown tool names and bad formats exit 1 with a stderr message', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-bad-'));
  const run = (extra: string[]) => new Promise<{ code: number | null; stderr: string }>((resolveRun, reject) => {
    const child = spawn(process.execPath, [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, ...extra], { stdio: ['pipe', 'pipe', 'pipe'] });
    let stderr = '';
    child.stderr.on('data', chunk => { stderr += String(chunk); });
    child.once('error', reject);
    child.once('exit', code => resolveRun({ code, stderr }));
  });
  try {
    const unknown = await run(['--tools', 'office_files,nope']);
    expect(unknown.code).toBe(1);
    expect(unknown.stderr).toContain('nope');
    expect(unknown.stderr).toContain('office_files');
    const badFormat = await run(['--formats', 'docx,odt']);
    expect(badFormat.code).toBe(1);
    expect(badFormat.stderr).toContain('docx, xlsx, pptx');
    const emptyFormat = await run(['--formats', '']);
    expect(emptyFormat.code).toBe(1);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('read-only stdio exposes inspection tools and refuses writes', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-readonly-'));
  const client = new Client({ name: 'test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, '--read-only'], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'report.docx'), await fixture());
    await client.connect(transport);
    const names = (await client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('office_open');
    expect(names).toContain('office_cells');
    for (const name of ['office_propose', 'office_propose_cells', 'office_accept', 'office_reject', 'office_export']) {
      expect(names).not.toContain(name);
      expect((await client.callTool({ name, arguments: {} })).isError).toBe(true);
    }
    expect((await client.callTool({ name: 'office_open', arguments: { path: 'report.docx' } })).isError).not.toBe(true);
    expect(await readdir(root)).toEqual(['report.docx']);
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);
