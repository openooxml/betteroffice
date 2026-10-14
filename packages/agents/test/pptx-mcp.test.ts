import { expect, test } from 'bun:test';
import { createCanvas } from '@napi-rs/canvas';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createOfficeMcpServer } from '../src/mcp';
import { openPptx } from '../src/pptx';
import { presentationFixture } from './pptx-fixture';

async function connect(readOnly = false) {
  const root = await mkdtemp(join(tmpdir(), 'pptx-mcp-'));
  await writeFile(join(root, 'deck.pptx'), await presentationFixture());
  const server = await createOfficeMcpServer({ root, readOnly });
  const client = new Client({ name: 'pptx-test', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  return { root, client, dispose: async () => { await client.close(); await server.close(); await rm(root, { recursive: true, force: true }); } };
}

test('PPTX MCP tools discover, inspect, batch edit, preview and export a deck', async () => {
  const context = await connect();
  const { client, root } = context;
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await client.callTool({ name, arguments: args });
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
    return result.structuredContent as Record<string, any>;
  };
  try {
    const listed = (await client.listTools()).tools;
    expect(listed.filter(tool => tool.name.startsWith('pptx_')).map(tool => tool.name).sort()).toEqual(['pptx_edit', 'pptx_outline', 'pptx_preview', 'pptx_read_slide']);
    for (const tool of listed.filter(tool => tool.name.startsWith('pptx_'))) expect(tool.inputSchema.additionalProperties).toBe(false);
    const { document } = await call('office_open', { path: 'deck.pptx' });
    const outline = await call('pptx_outline', { document });
    const slide = outline.slides[0].id;
    const shape = outline.slides[0].shapes[0].id;
    expect((await call('pptx_read_slide', { document, slide, shape })).shapes[0].stories[0].paragraphs[0].runs[0].formatting.bold).toBe(true);
    const edited = await call('pptx_edit', { document, version: outline.version, edits: [{ op: 'replace_text', slide, shape, text: 'Annual result' }, { op: 'set_notes', slide, text: 'Present annual result' }] });
    expect(edited.results.map((result: any) => result.changed)).toEqual([true, true]);
    const preview = await client.callTool({ name: 'pptx_preview', arguments: { document, slide, scale: 0.5 } });
    expect(preview.isError).not.toBe(true);
    expect((preview.content as Array<{ type: string }>).some(block => block.type === 'image')).toBe(true);
    expect(preview.structuredContent).toMatchObject({ page: 1, pageCount: 1, width: 480, height: 360 });
    await call('office_export', { document, path: 'saved.pptx' });
    const saved = await openPptx(await readFile(join(root, 'saved.pptx')));
    try { expect(saved.outline().slides[0].title).toBe('Annual result'); }
    finally { saved.close(); }
    const stale = await client.callTool({ name: 'pptx_edit', arguments: { document, version: outline.version, edits: [{ op: 'set_notes', slide, text: 'wrong' }] } });
    expect(stale.structuredContent).toMatchObject({ code: 'STALE_TARGET' });
    const invalid = await client.callTool({ name: 'pptx_edit', arguments: { document, version: edited.version, edits: [{ op: 'move_slide', slide, index: 0, unexpected: true }] } });
    expect(invalid.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    const badShape = await client.callTool({ name: 'pptx_read_slide', arguments: { document, slide, shape: 'wrong' } });
    expect(badShape.structuredContent).toMatchObject({ code: 'UNKNOWN_SHAPE', details: { nearestShapeIds: [{ id: shape }] } });
    const escape = await client.callTool({ name: 'office_export', arguments: { document, path: '../outside.pptx' } });
    expect(escape.structuredContent).toMatchObject({ code: 'OUTSIDE_WORKSPACE' });
  } finally { await context.dispose(); }
});

test('PPTX MCP read-only mode excludes writes and rejects oversized read inputs', async () => {
  const context = await connect(true);
  try {
    const names = (await context.client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('pptx_preview');
    expect(names).not.toContain('pptx_edit');
    const invalid = await context.client.callTool({ name: 'pptx_outline', arguments: { document: 'doc1', limit: 101 } });
    expect(invalid.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    const unknown = await context.client.callTool({ name: 'pptx_edit', arguments: {} });
    expect(unknown.structuredContent).toMatchObject({ code: 'UNKNOWN_TOOL' });
  } finally { await context.dispose(); }
});

test('PPTX MCP adds workspace images and refuses traversal and symlink escapes atomically', async () => {
  const context = await connect();
  const outside = await mkdtemp(join(tmpdir(), 'pptx-image-outside-'));
  try {
    const bytes = createCanvas(2, 2).toBuffer('image/png');
    await writeFile(join(context.root, 'image.png'), bytes);
    await writeFile(join(outside, 'outside.png'), bytes);
    await symlink(join(outside, 'outside.png'), join(context.root, 'escape.png'));
    const opened = await context.client.callTool({ name: 'office_open', arguments: { path: 'deck.pptx' } });
    const document = (opened.structuredContent as any).document;
    const listed = await context.client.callTool({ name: 'pptx_outline', arguments: { document } });
    const outline = listed.structuredContent as any;
    const slide = outline.slides[0].id;
    const rect = { x: 0, y: 0, width: 50, height: 50 };
    for (const path of [join(outside, 'outside.png'), 'escape.png']) {
      const rejected = await context.client.callTool({ name: 'pptx_edit', arguments: {
        document, version: outline.version,
        edits: [{ op: 'set_notes', slide, text: 'Should not apply' }, { op: 'add_image_file', slide, name: 'Image', rect, path }],
      } });
      expect(rejected.structuredContent).toMatchObject({ code: 'OUTSIDE_WORKSPACE', details: { failedIndex: 1 } });
      const current = await context.client.callTool({ name: 'pptx_read_slide', arguments: { document, slide } });
      expect((current.structuredContent as any).notes.text).toBe('Present the quarterly result.');
      expect((current.structuredContent as any).version).toBe(outline.version);
    }
    const added = await context.client.callTool({ name: 'pptx_edit', arguments: {
      document, version: outline.version, edits: [{ op: 'add_image_file', slide, name: 'Image', rect, path: 'image.png' }],
    } });
    expect(added.isError).not.toBe(true);
    expect(added.structuredContent).toMatchObject({ results: [{ op: 'add_image_file', changed: true }] });
  } finally { await context.dispose(); await rm(outside, { recursive: true, force: true }); }
});
