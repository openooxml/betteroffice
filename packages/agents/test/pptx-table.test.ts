import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { openPresentation } from '@betteroffice/pptx';
import { createOfficeMcpServer } from '../src/mcp';
import { attachPptx, openPptx } from '../src/pptx';
import { tableFixture } from './pptx-fixture';

async function connect(merged = false) {
  const root = await mkdtemp(join(tmpdir(), 'pptx-table-'));
  const source = await tableFixture(merged);
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
  const shape = outline.slides[0].shapes.find((s: any) => s.name === 'Table 8').id;
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
