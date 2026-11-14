import { beforeAll, expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { initWasm, openPresentation } from '@betteroffice/pptx';
import { createOfficeMcpServer } from '../src/mcp';
import { attachPptx, openPptx } from '../src/pptx';
import { pptxEditSchema } from '../src/pptx-schema';
import { tableFixture } from './pptx-fixture';
import { initialize } from '../src/prototype';

beforeAll(() => initialize('pptx', initWasm));

async function connect(merged = false, fixture?: Uint8Array) {
  const root = await mkdtemp(join(tmpdir(), 'pptx-table-'));
  const source = fixture ?? await tableFixture(merged);
  await writeFile(join(root, 'deck.pptx'), source);
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'table-test', version: '1' });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(b);
  await client.connect(a);
  const raw = async (name: string, args: object) => client.callTool({ name, arguments: args as Record<string, unknown> });
  const call = async (name: string, args: object): Promise<any> => {
    const result = await raw(name, args);
    expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
    return result.structuredContent;
  };
  const { document } = await call('office_open', { path: 'deck.pptx' });
  const outline = await call('pptx_outline', { document });
  const slide = outline.slides[0].id;
  const shape = (outline.slides[0].shapes.find((s: any) => s.name === 'Table 8') ?? outline.slides[0].shapes[0]).id;
  return { root, source, document, slide, shape, call, raw, version: outline.version,
    close: async () => { await client.close(); await server.close(); await rm(root, { recursive: true, force: true }); } };
}

test('pptx_edit row and column operations preserve other content and reject stale refs', async () => {
  const c = await connect();
  try {
    const { document, slide, shape } = c;
    const hits = await c.call('office_grep', { document, query: 'R0C2' });
    const oldRef = hits.matches.find((m: any) => m.shapeId === shape).ref;
    let result = await c.call('pptx_edit', { document, version: c.version, edits: [{ op: 'delete_table_column', shape, column: 1 }] });
    expect(result.results[0]).toMatchObject({ rowCount: 3, columnCount: 2, newCellRefs: [] });
    const stale = await c.raw('pptx_edit', { document, version: result.version, edits: [{ op: 'delete_table_column', ref: oldRef }] });
    expect(stale.structuredContent).toMatchObject({ code: 'UNKNOWN_REF' });
    const fresh = await c.call('office_grep', { document, query: 'R0C2' });
    const ref = fresh.matches.find((m: any) => m.shapeId === shape).ref;
    expect(ref).not.toBe(oldRef);
    result = await c.call('pptx_edit', { document, version: result.version, edits: [{ op: 'insert_table_column', ref, position: 'before', cells: ['New column'] }] });
    expect(result.results[0]).toMatchObject({ rowCount: 3, columnCount: 3 });
    expect(result.results[0].newCellRefs).toHaveLength(3);
    expect((await c.call('office_read', { document, ref: result.results[0].newCellRefs[0] })).text).toBe('New column');
    result = await c.call('pptx_edit', { document, version: result.version, edits: [{ op: 'insert_table_row', slide, shape, row: 2, position: 'after', cells: ['New row', 'Second'] }] });
    expect(result.results[0]).toMatchObject({ rowCount: 4, columnCount: 3 });
    expect(result.results[0].newCellRefs).toHaveLength(3);
    const rowRef = result.results[0].newCellRefs[2];
    result = await c.call('pptx_edit', { document, version: result.version, edits: [{ op: 'delete_table_row', ref: rowRef }] });
    expect(result.results[0]).toMatchObject({ rowCount: 3, columnCount: 3 });
    await c.call('office_export', { document, path: 'saved.pptx' });
    const saved = await readFile(join(c.root, 'saved.pptx'));
    const before = await JSZip.loadAsync(c.source);
    const after = await JSZip.loadAsync(saved);
    for (const path of Object.keys(before.files).filter(path => path !== 'ppt/slides/slide1.xml' && !before.files[path].dir)) {
      expect(await after.file(path)!.async('uint8array'), path).toEqual(await before.file(path)!.async('uint8array'));
    }
    const xml = await after.file('ppt/slides/slide1.xml')!.async('string');
    const original = await before.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml.match(/<p:sp>.*?<\/p:sp>/su)?.[0]).toBe(original.match(/<p:sp>.*?<\/p:sp>/su)?.[0]);
    expect(xml.match(/<p:graphicFrame>.*?<\/p:graphicFrame>/gsu)?.[1]).toBe(original.match(/<p:graphicFrame>.*?<\/p:graphicFrame>/gsu)?.[1]);
    expect(xml).toContain('strike="sngStrike"');
    expect(xml).toContain('marL="91440"');
    const reopened = await openPptx(saved);
    try { expect(reopened.grep({ query: 'New column' }).matches).toHaveLength(1); }
    finally { reopened.close(); }
  } finally { await c.close(); }
});

test('pptx_edit table batches roll back and name invalid indices and cells', async () => {
  const c = await connect();
  try {
    const { document, slide, shape } = c;
    const failed = await c.raw('pptx_edit', { document, version: c.version, edits: [
      { op: 'delete_table_column', shape, column: 1 }, { op: 'replace_text', slide, shape: 'missing', text: 'Wrong' },
    ] });
    expect(failed.structuredContent).toMatchObject({ details: { failedIndex: 1 } });
    expect((await c.call('pptx_outline', { document })).version).toBe(c.version);
    expect((await c.call('office_grep', { document, query: 'R0C1' })).matches).toHaveLength(2);
    for (const edit of [
      { op: 'delete_table_row', shape, row: 3 },
      { op: 'delete_table_column', shape, column: 3 },
      { op: 'insert_table_row', shape, row: 0, position: 'before', cells: ['a', 'b', 'c', 'd'] },
      { op: 'insert_table_column', shape, column: 0, position: 'after', cells: ['two\nlines'] },
    ]) {
      const invalid = await c.raw('pptx_edit', { document, version: c.version, edits: [edit] });
      expect(invalid.isError).toBe(true);
      expect(JSON.stringify(invalid.structuredContent)).toContain(edit.op === 'delete_table_row' ? 'row' : edit.op === 'delete_table_column' ? 'column' : 'cells');
    }
  } finally { await c.close(); }
});

test('pptx_edit refuses merged table structure without changing the deck', async () => {
  const c = await connect(true);
  try {
    const result = await c.raw('pptx_edit', { document: c.document, version: c.version, edits: [{ op: 'delete_table_column', shape: c.shape, column: 0 }] });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.structuredContent)).toContain('merged');
    expect((await c.call('pptx_outline', { document: c.document })).version).toBe(c.version);
  } finally { await c.close(); }
});

test('pptx_edit refuses table structure on a collaborative handle without changing it', async () => {
  const handle = openPresentation(await tableFixture());
  handle.setCollaborative();
  const deck = attachPptx(handle);
  try {
    const shape = handle.snapshot().slides[0].shapes.find(s => s.name === 'Table 8')!.id;
    const version = handle.version();
    let code = '';
    try { deck.edit({ version, edits: [{ op: 'insert_table_row', shape, row: 0, position: 'after' }] }); }
    catch (error) { code = (error as { code?: string }).code ?? ''; expect(String(error)).toContain('collaborative session'); }
    expect(code).toBe('UNSUPPORTED');
    expect(handle.version()).toBe(version);
  } finally { deck.close(); handle.dispose(); }
});

test('pptx_edit copy column width preserves the grid and refits the frame', async () => {
  const c = await connect();
  try {
    const before = await JSZip.loadAsync(c.source);
    const original = await before.file('ppt/slides/slide1.xml')!.async('string');
    const frame = original.match(/<p:graphicFrame>.*?<\/p:graphicFrame>/su)![0];
    const widths = [...frame.matchAll(/<a:gridCol w="(\d+)"/gu)].map(match => Number(match[1]));
    await c.call('pptx_edit', { document: c.document, version: c.version, edits: [
      { op: 'insert_table_column', shape: c.shape, column: 1, position: 'after', width: 'copy' },
    ] });
    await c.call('office_export', { document: c.document, path: 'copy.pptx' });
    const saved = await JSZip.loadAsync(await readFile(join(c.root, 'copy.pptx')));
    const xml = await saved.file('ppt/slides/slide1.xml')!.async('string');
    const after = xml.match(/<p:graphicFrame>.*?<\/p:graphicFrame>/su)![0];
    const expected = [widths[0], widths[1], widths[1], widths[2]];
    expect([...after.matchAll(/<a:gridCol w="(\d+)"/gu)].map(match => Number(match[1]))).toEqual(expected);
    expect(Number(after.match(/<a:ext\b[^>]*\bcx="(\d+)"/u)![1])).toBe(expected.reduce((sum, width) => sum + width, 0));
    expect(after.match(/<a:off[^>]*\/>/u)?.[0]).toBe(frame.match(/<a:off[^>]*\/>/u)?.[0]);
  } finally { await c.close(); }
});

test('insert_table_column schema names invalid width values', () => {
  const result = pptxEditSchema.safeParse({ op: 'insert_table_column', shape: 'table', column: 1, position: 'after', width: 'wide' });
  expect(result.success).toBe(false);
  if (!result.success) {
    expect(result.error.issues.some(issue => issue.path.includes('width'))).toBe(true);
    expect(result.error.message).toContain('width');
  }
});

test.each(['insertTableRow', 'deleteTableRow', 'insertTableColumn', 'deleteTableColumn'] as const)('office_propose exposes %s and accepts table edits atomically', async type => {
  const c = await connect();
  try {
    const { document, slide, shape } = c;
    const before = await c.call('pptx_read_slide', { document, slide });
    const edit = type === 'insertTableRow' || type === 'insertTableColumn'
      ? { type, slideId: slide, shapeId: shape, index: 0, position: 'after', cells: ['Agent cell'] }
      : { type, slideId: slide, shapeId: shape, index: 0 };
    const pending = await c.call('office_propose', { document, author: 'agent', edits: [edit] });
    expect(await c.call('pptx_read_slide', { document, slide })).toEqual(before);
    await c.call('office_reject', { document, proposal: pending.id });
    expect(await c.call('pptx_read_slide', { document, slide })).toEqual(before);
    const accepted = await c.call('office_propose', { document, author: 'agent', edits: [edit] });
    await c.call('office_accept', { document, proposal: accepted.id });
    const after = await c.call('pptx_read_slide', { document, slide });
    const table = after.shapes.find((entry: any) => entry.id === shape).table;
    expect(table.rows).toBe(type === 'insertTableRow' ? 4 : type === 'deleteTableRow' ? 2 : 3);
    expect(table.columns).toBe(type === 'insertTableColumn' ? 4 : type === 'deleteTableColumn' ? 2 : 3);
    const invalid = await c.raw('office_propose', { document, author: 'agent', edits: [{ ...edit, index: 99 }] });
    expect(invalid.isError).toBe(true);
    expect(await c.call('pptx_read_slide', { document, slide })).toEqual(after);
  } finally { await c.close(); }
});

test('table proposal schema rejects invalid positions, widths, indices and extra keys', async () => {
  const { pptxProposalEditSchema } = await import('../src/pptx-schema');
  const edit = { type: 'insertTableColumn', slideId: '@slide', shapeId: '@shape', index: 0, position: 'before' };
  expect(pptxProposalEditSchema.safeParse(edit).success).toBe(true);
  expect(pptxProposalEditSchema.safeParse({ ...edit, width: 'copy', cells: ['A'] }).success).toBe(true);
  for (const extra of [{ index: -1 }, { index: 0.5 }, { index: 4294967296 }, { position: 'middle' }, { width: 'stretch' }, { cells: [1] }, { extra: true }]) {
    expect(pptxProposalEditSchema.safeParse({ ...edit, ...extra }).success).toBe(false);
  }
});

test.each(['front', 'back', 'forward', 'backward'] as const)('office_propose exposes reorderShape %s without changing the live order before acceptance', async to => {
  const fixture = await readFile(new URL('../../../crates/pptx-edit/tests/fixtures/proposal-shapes.pptx', import.meta.url));
  const c = await connect(false, fixture);
  try {
    const { document, slide } = c;
    const before = await c.call('pptx_read_slide', { document, slide });
    const shape = before.shapes[to === 'back' || to === 'backward' ? 1 : 0].id;
    const edit = { type: 'reorderShape', slideId: slide, shapeId: shape, to };
    const pending = await c.call('office_propose', { document, author: 'agent', edits: [edit] });
    expect(pending.changes[0]).toMatchObject({ ref: shape, slide: 1, kind: 'content' });
    expect(await c.call('pptx_read_slide', { document, slide })).toEqual(before);
    await c.call('office_reject', { document, proposal: pending.id });
    expect(await c.call('pptx_read_slide', { document, slide })).toEqual(before);
    const accepted = await c.call('office_propose', { document, author: 'agent', edits: [edit] });
    await c.call('office_accept', { document, proposal: accepted.id });
    const after = await c.call('pptx_read_slide', { document, slide });
    expect(after.shapes.map((entry: any) => entry.id)).toEqual(before.shapes.map((entry: any) => entry.id).reverse());
    for (const invalid of [{ ...edit, to: 'up' }, { ...edit, extra: true }, { ...edit, shapeId: 'missing' }]) {
      expect((await c.raw('office_propose', { document, author: 'agent', edits: [invalid] })).isError).toBe(true);
      expect(await c.call('pptx_read_slide', { document, slide })).toEqual(after);
    }
  } finally { await c.close(); }
});

test('reorder proposal schema accepts both refs and strictly validates direction and keys', async () => {
  const { pptxProposalEditSchema } = await import('../src/pptx-schema');
  const edit = { type: 'reorderShape', slideId: '@slide', shapeId: '@shape', to: 'front' };
  for (const to of ['front', 'back', 'forward', 'backward']) {
    expect(pptxProposalEditSchema.safeParse({ ...edit, to }).success).toBe(true);
  }
  for (const invalid of [{ ...edit, to: 'up' }, { ...edit, to: null }, { ...edit, to: undefined }, { ...edit, extra: true }]) {
    expect(pptxProposalEditSchema.safeParse(invalid).success).toBe(false);
  }
});
