import { expect, test } from 'bun:test';
import { createCanvas } from '@napi-rs/canvas';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createOfficeMcpServer } from '../src/mcp';
import { openPptx } from '../src/pptx';
import { presentationFixture, layoutFixture } from './pptx-fixture';

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

test('PPTX MCP read-only mode excludes writes and clamps oversized read inputs', async () => {
  const context = await connect(true);
  try {
    const names = (await context.client.listTools()).tools.map(tool => tool.name);
    expect(names).toContain('pptx_preview');
    expect(names).not.toContain('pptx_edit');
    const opened = await context.client.callTool({ name: 'office_open', arguments: { path: 'deck.pptx' } });
    const clamped = await context.client.callTool({ name: 'pptx_outline', arguments: { document: (opened.structuredContent as any).document, limit: 101 } });
    expect(clamped.isError).not.toBe(true);
    expect(JSON.stringify(clamped.structuredContent)).toContain('limit clamped to 100');
    const unknown = await context.client.callTool({ name: 'pptx_edit', arguments: {} });
    expect(unknown.structuredContent).toMatchObject({ code: 'UNKNOWN_TOOL' });
  } finally { await context.dispose(); }
});

test('PPTX MCP edit validation names the field that is wrong', async () => {
  const context = await connect();
  try {
    const tool = (await context.client.listTools()).tools.find(tool => tool.name === 'pptx_edit') as any;
    const item = tool.inputSchema.properties.edits.items;
    expect(Object.keys(item)).toEqual(['oneOf']);
    expect(item.oneOf.map((variant: any) => variant.properties.op.const)).toEqual([
      'replace_text', 'set_paragraphs', 'find_replace', 'add_slide', 'set_slide_layout', 'duplicate_slide', 'delete_slide',
      'move_slide', 'add_text_box', 'add_shape', 'add_image', 'insert_table_row', 'delete_table_row', 'insert_table_column', 'delete_table_column', 'add_comment', 'reply_comment', 'resolve_comment', 'set_notes', 'add_image_file',
    ]);
    for (const variant of item.oneOf) expect(variant.additionalProperties).toBe(false);
    const opened = await context.client.callTool({ name: 'office_open', arguments: { path: 'deck.pptx' } });
    const document = (opened.structuredContent as any).document;
    const outline = (await context.client.callTool({ name: 'pptx_outline', arguments: { document } })).structuredContent as any;
    const slide = outline.slides[0].id;
    const edit =(edits: object[]) => context.client.callTool({ name: 'pptx_edit', arguments: { document, version: outline.version, edits } });
    const missing = await edit([{ op: 'replace_text', shape: 'shape', text: 'x' }]);
    expect(missing.structuredContent).toMatchObject({ code: 'UNKNOWN_SHAPE' });
    const unknownKey = await edit([{ op: 'set_slide_layout', slide, layout: 'layout', mapping: {} }]);
    expect(unknownKey.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT', issues: [{ path: ['edits', 0] }] });
    expect(JSON.stringify(unknownKey.structuredContent)).toContain('mapping');
    const image = await edit([{ op: 'add_image_file', slide, name: 'Logo', rect: { x: 0, y: 0, width: 10, height: 10 } }]);
    expect(image.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT', issues: [{ path: ['edits', 0, 'path'] }] });
    const author = await edit([{ op: 'add_comment', slide, text: 'Comment' }]);
    expect(author.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT', issues: [{ path: ['edits', 0, 'author'] }] });
    const position = await edit([{ op: 'add_comment', slide, author: 'Writer', text: 'Comment', position: { x: 'wrong', y: 0 } }]);
    expect(position.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT', issues: [{ path: ['edits', 0, 'position', 'x'] }] });
    const comment = await edit([{ op: 'reply_comment', comment: 'missing', author: 'Writer', text: 'Reply' }]);
    expect(comment.structuredContent).toMatchObject({ code: 'UNKNOWN_REF', details: { validComments: [] } });
    expect(JSON.stringify(comment.structuredContent)).toContain('comment must be an ID');
    const resolved = await edit([{ op: 'resolve_comment', comment: 'missing', done: 'yes' }]);
    expect(resolved.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT', issues: [{ path: ['edits', 0, 'done'] }] });
  } finally { await context.dispose(); }
});

test('PPTX MCP set_notes text describes its paragraph lines', async () => {
  const context = await connect();
  try {
    const edit = (await context.client.listTools()).tools.find(tool => tool.name === 'pptx_edit');
    const variants: Record<string, any>[] = [];
    const visit = (node: unknown): void => {
      if (Array.isArray(node)) node.forEach(visit);
      else if (node && typeof node === 'object') {
        const schema = node as Record<string, any>;
        if (schema.properties?.op?.const !== undefined) variants.push(schema);
        Object.values(schema).forEach(visit);
      }
    };
    visit(edit?.inputSchema);
    expect(variants.find(variant => variant.properties.op.const === 'set_notes')?.properties.text.description)
      .toBe('One paragraph per line; a blank line adds an empty paragraph.');
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

test('PPTX MCP exposes and applies explicit layout mappings', async () => {
  const context = await connect();
  try {
    await writeFile(join(context.root, 'layout.pptx'), await layoutFixture());
    const call = async (name: string, args: Record<string, unknown>) => {
      const result = await context.client.callTool({ name, arguments: args });
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
      return result.structuredContent as any;
    };
    const { document } = await call('office_open', { path: 'layout.pptx' });
    const outline = await call('pptx_outline', { document });
    const layout = outline.layouts.find((layout: any) => layout.name === 'Title Only');
    expect(layout.placeholders[0].name).toBe('Heading');
    const slide = outline.slides[0];
    const edit = { op: 'set_slide_layout', slide: slide.id, layout: layout.id, placeholders: [{ shape: slide.shapes[0].id, layoutPlaceholder: '8' }] };
    const rejected = await context.client.callTool({ name: 'pptx_edit', arguments: { document, version: outline.version, edits: [{ op: 'set_notes', slide: slide.id, text: 'rollback' }, { ...edit, layout: 'missing' }] } });
    expect(rejected.isError).toBe(true);
    await call('pptx_edit', { document, version: outline.version, edits: [edit] });
    const stale = await context.client.callTool({ name: 'pptx_edit', arguments: { document, version: outline.version, edits: [edit] } });
    expect(stale.structuredContent).toMatchObject({ code: 'STALE_TARGET' });
    await call('office_export', { document, path: 'layout-saved.pptx' });
    const reopened = await openPptx(await readFile(join(context.root, 'layout-saved.pptx')));
    try { expect(reopened.outline().slides[0].layoutId).toBe(layout.id); }
    finally { reopened.close(); }
  } finally { await context.dispose(); }
});

for (const flavor of ['legacy', 'modern'] as const) test(`PPTX MCP ${flavor} comments read, edit, duplicate and round trip atomically`, async () => {
  const context = await connect();
  const call = async (name: string, args: Record<string, unknown>) => {
    const result = await context.client.callTool({ name, arguments: args });
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
    return result.structuredContent as any;
  };
  try {
    const { openPresentation } = await import('@betteroffice/pptx');
    const initialized = await openPptx(await presentationFixture());
    initialized.close();
    const seed = openPresentation(await presentationFixture());
    seed.setCommentFlavor(flavor);
    seed.addComment(seed.snapshot().slides[0].id, { author: 'Reviewer', initials: 'RV', text: 'Review this slide', created: '2026-01-01T10:00:00Z', xEmu: 12700, yEmu: 25400 });
    await writeFile(join(context.root, 'deck.pptx'), seed.save());
    seed.dispose();
    const { document } = await call('office_open', { path: 'deck.pptx' });
    let outline = await call('office_outline', { document, story: 'comments' });
    const root = outline.items[0];
    expect(root).toMatchObject({ author: 'Reviewer', initials: 'RV', text: 'Review this slide', date: '2026-01-01T10:00:00Z', parentId: null, resolved: false, position: { x: 1, y: 2 }, slide: 1 });
    const slide = root.slideId;
    const text = await call('office_read', { document, ref: root.ref, length: 6 });
    expect(text).toMatchObject({ text: 'Review', nextStart: 6, slideId: slide });
    const reply = await call('pptx_edit', { document, version: outline.version, edits: [{ op: 'reply_comment', comment: root.id, author: 'Writer', initials: 'WR', text: 'Addressed & checked', date: '2026-01-01T11:00:00Z' }] });
    expect(typeof reply.results[0].comment).toBe('string');
    const replyId = reply.results[0].comment;
    expect(reply.results[0].parentId).toBe(root.id);
    const rejected = await context.client.callTool({ name: 'pptx_edit', arguments: { document, version: reply.version, edits: [{ op: 'add_comment', slide, author: 'Writer', text: 'Must roll back' }, { op: 'resolve_comment', comment: flavor === 'legacy' ? root.id : 'missing' }] } });
    expect(rejected.isError).toBe(true);
    expect(rejected.structuredContent).toMatchObject({ code: flavor === 'legacy' ? 'UNSUPPORTED_TARGET' : 'UNKNOWN_REF', details: { failedIndex: 1 } });
    if (flavor === 'legacy') expect(JSON.stringify(rejected.structuredContent)).toContain('legacy');
    outline = await call('office_outline', { document, story: 'comments' });
    expect(outline.version).toBe(reply.version);
    expect(outline.total).toBe(2);
    expect(outline.items.find((c: any) => c.parentId === root.id)).toMatchObject({ author: 'Writer', initials: 'WR', position: root.position, date: '2026-01-01T11:00:00Z' });
    if (flavor === 'modern') {
      const resolved = await call('pptx_edit', { document, version: outline.version, edits: [{ op: 'resolve_comment', comment: replyId }] });
      expect(resolved.results[0]).toMatchObject({ comment: root.id, resolved: true });
      outline = await call('office_outline', { document, story: 'comments' });
      expect(outline.items.find((c: any) => c.id === root.id).resolved).toBe(true);
      const conflict = await context.client.callTool({ name: 'pptx_edit', arguments: { document, version: outline.version, edits: [{ op: 'resolve_comment', comment: root.id, resolved: false, done: true }] } });
      expect(conflict.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
      await call('pptx_edit', { document, version: outline.version, edits: [{ op: 'resolve_comment', comment: root.id, done: false }] });
      outline = await call('office_outline', { document, story: 'comments' });
      expect(outline.items.find((c: any) => c.id === root.id).resolved).toBe(false);
    }
    await call('pptx_edit', { document, version: outline.version, edits: [{ op: 'add_comment', slide, author: 'Other', text: 'A new comment', position: { x: 3, y: 4 } }, { op: 'duplicate_slide', slide }] });
    const final = await call('office_outline', { document, story: 'comments', limit: 2 });
    expect(final.total).toBe(6);
    expect(final.nextOffset).toBe(2);
    await call('office_export', { document, path: 'comments-saved.pptx' });
    const reopened = await openPptx(await readFile(join(context.root, 'comments-saved.pptx')));
    try {
      const comments = reopened.list({ story: 'comments', limit: 100 });
      expect(comments.total).toBe(6);
      expect(comments.items.filter((c: any) => c.parentId !== null)).toHaveLength(2);
    } finally { reopened.close(); }
  } finally { await context.dispose(); }
});
