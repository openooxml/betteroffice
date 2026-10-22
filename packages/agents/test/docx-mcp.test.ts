import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { openDocx } from '../src';
import { createOfficeMcpServer } from '../src/mcp';
import { commentFixture, editFixture, tableCell, tableFixture } from './fixture';

test('stdio docx_edit applies a style, format and table-row batch that survives export', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-edit-mcp-'));
  const client = new Client({ name: 'docx-edit-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' });
  let stderr = '';
  transport.stderr?.on('data', chunk => { stderr += String(chunk); });
  try {
    await writeFile(join(root, 'report.docx'), await editFixture());
    await client.connect(transport);
    const call = async (name: string, args: Record<string, unknown>) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const tool = (await client.listTools()).tools.find(item => item.name === 'docx_edit');
    expect(tool?.annotations?.readOnlyHint).toBe(false);
    expect(tool?.inputSchema.additionalProperties).toBe(false);
    const schema = JSON.stringify(tool?.inputSchema);
    for (const op of ['set_paragraph_style', 'format_text', 'insert_paragraphs', 'delete_paragraph', 'insert_table_row', 'delete_table_row']) expect(schema).toContain(op);
    expect(schema).toContain('runFormat');
    expect(schema).toContain('anchor');

    const { document } = await call('office_open', { path: 'report.docx' });
    const outline = await call('office_outline', { document, limit: 100 });
    const summary = outline.items.find((item: any) => item.text.startsWith('Revenue grew'));
    const q1 = outline.items.find((item: any) => item.text === 'Q1');
    expect(q1.table).toEqual({ story: 'body', index: 1, row: 2, column: 1, rows: 2, columns: 2 });
    const [strongly] = (await call('office_grep', { document, query: 'strongly' })).matches;
    const edited = await call('docx_edit', { document, version: outline.version, edits: [
      { op: 'set_paragraph_style', paragraph: summary.ref, style: 'Heading2' },
      { op: 'format_text', match: strongly.match, italic: true },
      { op: 'insert_table_row', row: q1.ref, position: 'after', cells: ['Q2', '1.4M'] },
    ] });
    expect(edited.results.map((item: any) => item.op)).toEqual(['set_paragraph_style', 'format_text', 'insert_table_row']);
    expect(edited.version).not.toBe(outline.version);
    expect((await call('office_read', { document, ref: edited.results[2].refs[0] })).table).toMatchObject({ row: 3, rows: 3 });
    const stale = await client.callTool({ name: 'docx_edit', arguments: { document, version: outline.version, edits: [{ op: 'delete_paragraph', paragraph: summary.ref }] } });
    expect(stale.isError).toBe(true);
    expect(stale.structuredContent).toMatchObject({ code: 'STALE_VERSION' });

    await call('office_export', { document, path: 'edited.docx' });
    const reopened = await openDocx(await readFile(join(root, 'edited.docx')));
    try {
      const items = reopened.list({ limit: 100 }).items;
      expect(items.find(item => item.text.startsWith('Revenue grew'))?.style).toBe('Heading2');
      const [match] = reopened.grep({ query: 'strongly' }).matches;
      expect(reopened.read(match.ref).runs.find(run => run.start <= match.start && run.end >= match.end)?.formatting.italic).toBe(true);
      expect(items.filter(item => item.table).map(item => item.text).sort()).toEqual(['1.0M', '1.4M', 'Q1', 'Q2', 'Quarter', 'Revenue']);
      expect(items.find(item => item.text === 'Q2')?.table).toMatchObject({ row: 3, rows: 3 });
    } finally { reopened.close(); }
    expect(stderr).toBe('');
  } finally {
    await client.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30000);

test('read-only server does not list docx_edit', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-edit-readonly-'));
  const server = await createOfficeMcpServer({ root, readOnly: true });
  const client = new Client({ name: 'docx-edit-readonly', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  try {
    const names = (await client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('office_outline');
    expect(names).not.toContain('docx_edit');
    expect((await client.callTool({ name: 'docx_edit', arguments: {} })).structuredContent).toMatchObject({ code: 'UNKNOWN_TOOL' });
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});

test('docx_edit refuses merged and final row batches atomically', async () => {
  const variants = [
    [tableCell('Merged', '<w:vMerge w:val="restart"/>') + tableCell('Top'), tableCell('', '<w:vMerge/>') + tableCell('Bottom')],
    [tableCell('Only') + tableCell('Row')],
  ];
  const root = await mkdtemp(join(tmpdir(), 'docx-row-refusal-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'docx-row-refusal', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  const call = async (name: string, args: Record<string, unknown>) => {
    const output = await client.callTool({ name, arguments: args });
    expect(output.isError).not.toBe(true);
    return output.structuredContent as Record<string, any>;
  };
  try {
    for (const [index, tableRows] of variants.entries()) {
      const path = `input-${index}.docx`;
      await writeFile(join(root, path), await tableFixture(tableRows));
      const { document } = await call('office_open', { path });
      const outline = await call('office_outline', { document, limit: 100 });
      const heading = outline.items.find((item: any) => item.text === 'Quarterly report');
      await call('office_export', { document, path: `before-${index}.docx` });
      const baseline = await (await JSZip.loadAsync(await readFile(join(root, `before-${index}.docx`)))).file('word/document.xml')!.async('string');
      const targets = index === 0 ? ['Merged', 'Bottom'] : ['Only'];
      let attempt = 0;
      for (const text of targets) {
        const row = outline.items.find((item: any) => item.text === text).ref;
        const edits = index === 0 ? [
          { op: 'delete_table_row', row },
          { op: 'insert_table_row', row, position: 'before', cells: ['New', 'Row'] },
          { op: 'insert_table_row', row, position: 'after', cells: ['New', 'Row'] },
        ] : [{ op: 'delete_table_row', row }];
        for (const edit of edits) {
          const result = await client.callTool({ name: 'docx_edit', arguments: { document, version: outline.version, edits: [
            { op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading2' }, edit,
          ] } });
          expect(result.isError).toBe(true);
          expect(result.structuredContent).toMatchObject({ code: 'UNSUPPORTED_TARGET', details: { failedIndex: 1 } });
          expect(await call('office_outline', { document, limit: 100 })).toEqual(outline);
          const after = `after-${index}-${attempt++}.docx`;
          await call('office_export', { document, path: after });
          const xml = await (await JSZip.loadAsync(await readFile(join(root, after)))).file('word/document.xml')!.async('string');
          expect(xml).toBe(baseline);
        }
      }
    }
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});

test('docx_edit add_comment anchors a comment that survives export and refreshes the version', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-comment-mcp-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'docx-comment', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  const call = async (name: string, args: Record<string, unknown>) => {
    const output = await client.callTool({ name, arguments: args });
    expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
    return output.structuredContent as Record<string, any>;
  };
  try {
    await writeFile(join(root, 'report.docx'), await commentFixture());
    const tool = (await client.listTools()).tools.find(item => item.name === 'docx_edit');
    expect(JSON.stringify(tool?.inputSchema)).toContain('add_comment');
    expect(tool?.description).toContain('one grep match in the body or a table;');
    expect(tool?.description).not.toContain('or a content control');
    const { document } = await call('office_open', { path: 'report.docx' });
    const found = await call('office_grep', { document, query: 'third quarter' });
    const edits = [{ op: 'add_comment', match: found.matches[0].match, author: 'Scheduling Review', text: 'Please confirm the dates.' }];
    const edited = await call('docx_edit', { document, version: found.version, edits });
    expect(edited.results).toEqual([expect.objectContaining({ op: 'add_comment', ref: found.matches[0].ref, comment: expect.any(String) })]);
    expect(edited.version).not.toBe(found.version);
    const stale = await client.callTool({ name: 'docx_edit', arguments: { document, version: found.version, edits } });
    expect(stale.structuredContent).toMatchObject({ code: 'STALE_VERSION' });
    await call('office_export', { document, path: 'commented.docx' });
    const zip = await JSZip.loadAsync(await readFile(join(root, 'commented.docx')));
    const comments = await zip.file('word/comments.xml')!.async('string');
    expect(comments).toContain('w:author="Scheduling Review"');
    expect(comments).toContain('Please confirm the dates.');
    expect(comments).toContain('Keep this note.');
    expect((await zip.file('word/document.xml')!.async('string')).match(/<w:commentReference /g)).toHaveLength(2);
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});
