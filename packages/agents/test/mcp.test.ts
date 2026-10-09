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
      expect(output.isError).not.toBe(true);
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

test('read-only stdio exposes inspection tools and refuses writes', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-mcp-readonly-'));
  const client = new Client({ name: 'test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, '--read-only'], stderr: 'pipe' });
  try {
    await writeFile(join(root, 'report.docx'), await fixture());
    await client.connect(transport);
    const names = (await client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('office_open');
    for (const name of ['office_propose', 'office_accept', 'office_reject', 'office_export']) {
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
