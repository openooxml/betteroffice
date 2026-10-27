import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createOfficeMcpServer } from '../src/mcp';
import { layoutFixture, presentationFixture, slideIdsFixture, tableFixture } from './pptx-fixture';
import { fixture } from './fixture';
import { implicitFontWorkbook } from './format-fixtures';

async function connect() {
  const root = await mkdtemp(join(tmpdir(), 'office-ergonomics-'));
  await writeFile(join(root, 'deck.pptx'), await slideIdsFixture([261, 279]));
  await writeFile(join(root, 'layout.pptx'), await layoutFixture());
  await writeFile(join(root, 'group.pptx'), await presentationFixture({ group: true }));
  await writeFile(join(root, 'table.pptx'), await tableFixture());
  await writeFile(join(root, 'report.docx'), await fixture(130));
  await writeFile(join(root, 'book.xlsx'), await implicitFontWorkbook());
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'ergonomics-test', version: '1' });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(b);
  await client.connect(a);
  const raw = async (name: string, args: Record<string, unknown>) => (await client.callTool({ name, arguments: args })).structuredContent as any;
  const call = async (name: string, args: Record<string, unknown>) => {
    const value = await raw(name, args);
    expect(value.code, JSON.stringify(value)).toBeUndefined();
    return value;
  };
  return { root, client, raw, call, dispose: async () => { await client.close(); await server.close(); await rm(root, { recursive: true, force: true }); } };
}

test('layout inputs, discovery and receipts use placeholders consistently', async () => {
  const c = await connect();
  try {
    const tool = (await c.client.listTools()).tools.find(t => t.name === 'pptx_edit')!;
    expect(tool.description).toContain('placeholders:[{shape, layoutPlaceholder}]');
    const variants = (tool.inputSchema.properties!.edits as any).items.oneOf;
    const layoutSchema = variants.find((v: any) => v.properties.op.const === 'set_slide_layout');
    expect(layoutSchema.properties.placeholders.items.properties.shape.description).toContain('slide');
    expect(layoutSchema.properties.placeholders.items.properties.layoutPlaceholder.description).toContain('layout');
    const { document } = await c.call('office_open', { path: 'layout.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const slide = outline.slides[0];
    const layout = outline.layouts.find((l: any) => l.name === 'Title Only');
    const invalid = await c.raw('pptx_edit', { document, version: outline.version, edits: [{ op: 'set_slide_layout', slide: slide.id, layout: layout.id, mapping: {} }] });
    expect(invalid.code).toBe('INVALID_ARGUMENT');
    expect(JSON.stringify(invalid)).toContain('placeholders:[{shape, layoutPlaceholder}]');
    const placeholders = [{ shape: slide.shapes[0].id, layoutPlaceholder: '8' }];
    const edited = await c.call('pptx_edit', { document, version: outline.version, edits: [{ op: 'set_slide_layout', slide: slide.id, layout: layout.id, placeholders }] });
    expect(edited.results[0].placeholders).toEqual(placeholders);
    expect(edited.results[0].mapping).toBeUndefined();
  } finally { await c.dispose(); }
});

test('shape targets infer slides and reject disagreement without applying edits', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'deck.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const [first, second] = outline.slides;
    const shape = first.shapes[0].id;
    const edited = await c.call('pptx_edit', { document, version: outline.version, edits: [{ op: 'replace_text', shape, text: 'Revised heading' }] });
    expect((await c.call('pptx_read_slide', { document, shape })).shapes[0].stories[0].text).toBe('Revised heading');
    const formatted = await c.call('pptx_edit', { document, version: edited.version, edits: [{ op: 'set_paragraphs', shape, paragraphs: [{ runs: [{ text: 'Scoped text', bold: true }] }] }] });
    const rejected = await c.raw('pptx_edit', { document, version: formatted.version, edits: [{ op: 'set_notes', slide: first.id, text: 'rollback' }, { op: 'replace_text', slide: second.id, shape, text: 'wrong' }] });
    expect(rejected.code).toBe('INVALID_ARGUMENT');
    expect(rejected.message).toContain('disagree');
    const current = await c.call('pptx_read_slide', { document, slide: first.id });
    expect(current.version).toBe(formatted.version);
    expect(current.shapes[0].stories[0].text).toBe('Scoped text');
  } finally { await c.dispose(); }
});

test('unknown edit discriminators list valid operations and the closest input shape', async () => {
  const c = await connect();
  try {
    for (const [name, edit, expected] of [
      ['pptx_edit', { op: 'set_shape_text' }, 'replace_text {shape, text, slide?'],
      ['pptx_edit', { op: 'set_text' }, 'replace_text {shape, text, slide?'],
      ['docx_edit', { op: 'set_paragraph_styles' }, 'set_paragraph_style {paragraph, style}'],
      ['xlsx_edit_sheets', { action: 'renam' }, 'rename {sheet, name}'],
    ] as const) {
      const value = await c.raw(name, { document: 'doc1', version: name === 'docx_edit' ? 0 : 'v0', edits: [edit] });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(JSON.stringify(value)).toContain('Valid');
      expect(JSON.stringify(value)).toContain(expected);
    }
  } finally { await c.dispose(); }
});

test('numeric slide addresses explain both interpretations in read and edit tools', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'deck.pptx' });
    const outline = await c.call('pptx_outline', { document });
    expect(outline.ids).toContain('0-based');
    for (const name of ['pptx_read_slide', 'pptx_preview', 'pptx_edit']) {
      const slide = '1';
      const value = await c.raw(name, { document, ...(name === 'pptx_edit' ? { version: outline.version, edits: [{ op: 'set_notes', slide, text: 'wrong' }] } : { slide }) });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(value.message).toContain('0-based');
      expect(value.message).toContain('1-based');
      expect(value.message).toContain(outline.slides[0].id);
      expect(value.message).toContain(outline.slides[1].id);
      expect(value.message).toContain('current 0-based position 1');
    }
  } finally { await c.dispose(); }
});

test('numeric slide errors try the source ID reading and always list valid slides', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'deck.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const validSlides = outline.slides.map((slide: any, index: number) => ({ id: slide.id, index: index + 1 }));
    for (const slide of ['261', '9']) {
      const value = await c.raw('pptx_read_slide', { document, slide });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(value.details.validSlides).toEqual(validSlides);
    }
    const source = await c.raw('pptx_edit', { document, version: outline.version, edits: [{ op: 'set_notes', slide: '261', text: 'wrong' }] });
    expect(source.code).toBe('INVALID_ARGUMENT');
    expect(source.message).toContain(`source ID 261: ${outline.slides[0].id}`);
    expect(source.details.validSlides).toEqual(validSlides);
  } finally { await c.dispose(); }
});

test('published slide fields are strings and numbers get a slide ID hint', async () => {
  const c = await connect();
  try {
    const tools = (await c.client.listTools()).tools;
    expect((tools.find(t => t.name === 'pptx_read_slide')!.inputSchema.properties!.slide as any).type).toBe('string');
    const move = (tools.find(t => t.name === 'pptx_edit')!.inputSchema.properties!.edits as any).items.oneOf.find((v: any) => v.properties.op.const === 'move_slide');
    for (const key of ['slide', 'after', 'before']) expect(move.properties[key].type).toBe('string');
    const { document } = await c.call('office_open', { path: 'deck.pptx' });
    const outline = await c.call('pptx_outline', { document });
    for (const name of ['pptx_read_slide', 'pptx_preview', 'pptx_edit']) {
      const value = await c.raw(name, { document, ...(name === 'pptx_edit' ? { version: outline.version, edits: [{ op: 'move_slide', slide: outline.slides[0].id, after: 1 }] } : { slide: 1 }) });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(value.message).toContain('0-based');
      expect(value.message).toContain('1-based');
      expect(value.message).toContain('pptx_outline');
    }
  } finally { await c.dispose(); }
});

test('invalid argument messages list the first five issues', async () => {
  const c = await connect();
  try {
    const value = await c.raw('pptx_edit', { document: 'doc1', version: 'v0', edits: Array.from({ length: 8 }, () => ({ op: 'set_notes' })) });
    expect(value.code).toBe('INVALID_ARGUMENT');
    expect(value.issues).toHaveLength(16);
    expect(value.message.split('; ')).toHaveLength(6);
    expect(value.message.endsWith('+11 more')).toBe(true);
  } finally { await c.dispose(); }
});

test('bounded reads clamp oversized requests and keep pagination', async () => {
  const c = await connect();
  try {
    const doc = await c.call('office_open', { path: 'report.docx' });
    const page = await c.call('office_outline', { document: doc.document, limit: 999 });
    expect(page.items).toHaveLength(100);
    expect(page.nextOffset).toBe(100);
    expect(page.clamping.join(' ')).toContain('limit clamped to 100');
    const deck = await c.call('office_open', { path: 'deck.pptx' });
    const outline = await c.call('pptx_outline', { document: deck.document, limit: 999, shapeLimit: 999 });
    expect(outline.clamping.join(' ')).toContain('shapeLimit clamped to 100');
    await c.call('pptx_edit', { document: deck.document, version: outline.version, edits: [{ op: 'replace_text', shape: outline.slides[0].shapes[0].id, text: 'X'.repeat(4500) }] });
    const read = await c.call('pptx_read_slide', { document: deck.document, slide: outline.slides[0].id, limit: 999, textLength: 99999 });
    expect(read.shapes[0].stories[0].text.length).toBe(4000);
    expect(read.shapes[0].stories[0].nextStart).toBe(4000);
    expect(read.notes.text).toBeDefined();
    expect(read.clamping.join(' ')).toContain('textLength clamped to 4000');
    expect(read).toHaveProperty('nextOffset');
    const bad = await c.raw('pptx_outline', { document: deck.document, limit: 0 });
    expect(bad.code).toBe('INVALID_ARGUMENT');
  } finally { await c.dispose(); }
});

test('date-only acceptance and comment dates normalize to midnight UTC', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'report.docx' });
    const hits = await c.call('office_grep', { document, query: '€4.2 million' });
    const proposal = await c.call('office_propose', { document, author: 'Reviewer', edits: [{ match: hits.matches[0].match, newText: '€5 million' }] });
    await c.call('office_accept', { document, proposal: proposal.id, tracked: true, date: '2026-04-12' });
    const current = await c.call('office_grep', { document, query: 'Executive summary' });
    const added = await c.call('docx_edit', { document, version: current.version, edits: [{ op: 'add_comment', match: current.matches[0].match, author: 'Reviewer', text: 'Check', date: '2026-04-12' }] });
    await c.call('docx_edit', { document, version: added.version, edits: [{ op: 'reply_comment', comment: added.results[0].comment, author: 'Reviewer', text: 'Checked', date: '2026-04-13' }] });
    await c.call('office_export', { document, path: 'dates.docx' });
    const zip = await JSZip.loadAsync(await readFile(join(c.root, 'dates.docx')));
    expect(await zip.file('word/document.xml')!.async('string')).toContain('2026-04-12T00:00:00Z');
    expect(await zip.file('word/comments.xml')!.async('string')).toContain('2026-04-13T00:00:00Z');
    const invalid = await c.raw('office_accept', { document, proposal: proposal.id, date: '2026-02-30' });
    expect(invalid.code).toBe('INVALID_ARGUMENT');
  } finally { await c.dispose(); }
});

test('xlsx range edits inherit the top-level sheet and allow explicit overrides', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'book.xlsx' });
    const initial = await c.call('xlsx_outline', { document });
    await c.call('xlsx_edit_sheets', { document, version: initial.version, edits: [{ action: 'add', name: 'Other' }] });
    const outline = await c.call('xlsx_outline', { document });
    const [first, second] = outline.items;
    await c.call('xlsx_write_range', { document, version: outline.version, sheet: first.sheetId, edits: [{ range: 'A100', values: [['default']] }, { sheet: second.sheetId, range: 'A100', formulas: [['=2+3']] }] });
    expect((await c.call('xlsx_read_range', { document, sheet: first.sheetId, range: 'A100' })).cells[0].value.value).toBe('default');
    expect((await c.call('xlsx_read_range', { document, sheet: second.sheetId, range: 'A100' })).cells[0].value.value).toBe(5);
    const current = await c.call('xlsx_outline', { document });
    const missing = await c.raw('xlsx_write_range', { document, version: current.version, edits: [{ range: 'A101', values: [[1]] }] });
    expect(missing.code).toBe('INVALID_ARGUMENT');
    expect(missing.message).toContain('sheet');
  } finally { await c.dispose(); }
});

test('find_replace can target one shape without changing neighboring shapes or slides', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'deck.pptx' });
    const opened = await c.call('pptx_outline', { document });
    await c.call('pptx_edit', { document, version: opened.version, edits: [{ op: 'replace_text', shape: opened.slides[1].shapes[0].id, text: 'Keep this paragraph.' }] });
    await c.call('office_export', { document, path: 'before.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const slide = outline.slides[0];
    const shape = slide.shapes[0].id;
    const added = await c.call('pptx_edit', { document, version: outline.version, edits: [{ op: 'add_text_box', slide: slide.id, name: 'Neighbor', rect: { x: 1, y: 1, width: 100, height: 20 }, text: 'Keep this paragraph.' }] });
    const neighbor = await c.call('pptx_edit', { document, version: added.version, edits: [{ op: 'replace_text', shape: added.results[0].shapeId, text: 'Keep this paragraph.' }] });
    const replaced = await c.call('pptx_edit', { document, version: neighbor.version, edits: [{ op: 'find_replace', shape, query: 'Keep', replacement: 'Change' }] });
    expect(replaced.results[0].changedShapes).toBe(1);
    const read = await c.call('pptx_read_slide', { document, slide: slide.id });
    expect(read.shapes.find((s: any) => s.id === shape).stories[0].text).toContain('Change');
    expect(read.shapes.find((s: any) => s.id === added.results[0].shapeId).stories[0].text).toBe('Keep this paragraph.');
    const mismatch = await c.raw('pptx_edit', { document, version: replaced.version, edits: [{ op: 'find_replace', slide: outline.slides[1].id, shape, query: 'Change', replacement: 'Wrong' }] });
    expect(mismatch.code).toBe('INVALID_ARGUMENT');
    await c.call('office_export', { document, path: 'scoped.pptx' });
    const before = await JSZip.loadAsync(await readFile(join(c.root, 'before.pptx')));
    const saved = await JSZip.loadAsync(await readFile(join(c.root, 'scoped.pptx')));
    expect(await before.file('ppt/slides/slide2.xml')!.async('string')).toContain('Keep this paragraph.');
    expect(await saved.file('ppt/slides/slide2.xml')!.async('string')).toBe(await before.file('ppt/slides/slide2.xml')!.async('string'));
  } finally { await c.dispose(); }
});

test('find_replace scoped to a shape without text names its child shapes', async () => {
  const c = await connect();
  try {
    const { document } = await c.call('office_open', { path: 'group.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const group = outline.slides[0].shapes.find((s: any) => s.childCount > 0);
    const child = outline.slides[0].shapes.find((s: any) => s.parentId === group.id);
    const value = await c.raw('pptx_edit', { document, version: outline.version, edits: [{ op: 'find_replace', shape: group.id, query: 'Keep', replacement: 'Change' }] });
    expect(value.code).toBe('INVALID_ARGUMENT');
    expect(value.message).toContain(child.id);
    const scoped = await c.call('pptx_edit', { document, version: outline.version, edits: [{ op: 'find_replace', shape: child.id, query: 'Keep', replacement: 'Change' }] });
    expect(scoped.results[0].replacements).toBeGreaterThan(0);
  } finally { await c.dispose(); }
});

test('table structure ops share slide inference, op hints and the published op list', async () => {
  const c = await connect();
  try {
    const tool = (await c.client.listTools()).tools.find(t => t.name === 'pptx_edit')!;
    for (const op of ['insert_table_row', 'delete_table_row', 'insert_table_column', 'delete_table_column']) expect(tool.description).toContain(`${op} {`);
    const insert = (tool.inputSchema.properties!.edits as any).items.oneOf.find((v: any) => v.properties.op.const === 'insert_table_row');
    expect(insert.properties.slide.description).toContain('inferred');
    const { document } = await c.call('office_open', { path: 'table.pptx' });
    const outline = await c.call('pptx_outline', { document });
    const [first, second] = outline.slides;
    const shape = first.shapes.find((s: any) => s.name === 'Table 8').id;
    const ref = (await c.call('office_grep', { document, query: 'R1C0' })).matches.find((m: any) => m.shapeId === shape).ref;
    for (const edit of [{ op: 'delete_table_row', shape, row: 1, slide: second.id }, { op: 'delete_table_column', ref, slide: second.id }]) {
      const value = await c.raw('pptx_edit', { document, version: outline.version, edits: [edit] });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(value.message).toContain('disagree');
      expect(value.message).toContain(first.id);
    }
    const numeric = await c.raw('pptx_edit', { document, version: outline.version, edits: [{ op: 'insert_table_row', shape, row: 0, position: 'after', slide: '0' }] });
    expect(numeric.code).toBe('INVALID_ARGUMENT');
    expect(numeric.message).toContain('0-based');
    expect((await c.call('pptx_outline', { document })).version).toBe(outline.version);
    const edited = await c.call('pptx_edit', { document, version: outline.version, edits: [{ op: 'delete_table_row', shape, row: 1, slide: first.id }] });
    expect(edited.results[0]).toMatchObject({ rowCount: 2, columnCount: 3 });
    for (const [op, expected] of [['insert_row', 'insert_table_row {position'], ['delete_column', 'delete_table_column {']] as const) {
      const value = await c.raw('pptx_edit', { document, version: edited.version, edits: [{ op }] });
      expect(value.code).toBe('INVALID_ARGUMENT');
      expect(value.message).toContain(expected);
    }
  } finally { await c.dispose(); }
});
