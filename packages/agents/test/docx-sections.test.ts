import { expect, test } from 'bun:test';
import { openDocx } from '../src';
import { docxBatchSchema } from '../src/docx-schema';
import { sectionPackage } from '../../docx/src/yrs/sectionTestPackage';

test('agents insert before a paragraph, set a versioned section and reopen exported sections', async () => {
  const doc = await openDocx(await sectionPackage());
  try {
    const outline = doc.list();
    await doc.edit({ version: outline.version, edits: [{ op: 'insert_section_break', before: outline.items[1].ref, kind: 'nextPage' }] });
    const fresh = doc.list({ sectionOffset: 1, sectionLimit: 1 });
    expect(fresh.sections.items[0]).toMatchObject({ number: 2, firstParagraph: outline.items[1].ref, lastParagraph: outline.items[2].ref });
    await doc.edit({ version: fresh.version, edits: [{ op: 'set_section', section: 2, orientation: 'landscape', margins: { top: 900, bottom: 900, left: 1000, right: 1000 } }] });
    await expect(doc.edit({ version: outline.version, edits: [{ op: 'set_section', section: 1, orientation: 'portrait' }] })).rejects.toMatchObject({ code: 'STALE_VERSION' });
    const reopened = await openDocx(await doc.export());
    try { expect(reopened.list().sections.items[1]).toMatchObject({ number: 2, orientation: 'landscape', pageWidth: 15840, pageHeight: 12240, margins: { top: 900, left: 1000 } }); }
    finally { reopened.close(); }
    const before = doc.session.encodeState();
    await expect(doc.edit({ version: doc.list().version, edits: [
      { op: 'set_section', section: 1, orientation: 'landscape' },
      { op: 'insert_section_break', before: outline.items[0].ref, kind: 'continuous' },
    ] })).rejects.toMatchObject({ code: 'UNSUPPORTED_TARGET' });
    expect(doc.session.encodeState()).toEqual(before);
  } finally { doc.close(); }
});

test('section break schema requires exactly one before or after ref', () => {
  const batch = (edit: unknown) => docxBatchSchema.safeParse({ version: 0, edits: [edit] }).success;
  expect(batch({ op: 'insert_section_break', before: 'p1', kind: 'oddPage' })).toBe(true);
  expect(batch({ op: 'insert_section_break', after: 'p1', kind: 'evenPage' })).toBe(true);
  expect(batch({ op: 'insert_section_break', kind: 'continuous' })).toBe(false);
  expect(batch({ op: 'insert_section_break', before: 'p1', after: 'p2', kind: 'continuous' })).toBe(false);
});

test('MCP discovers section edits and exports a section setup that reopens', async () => {
  const { mkdtemp, writeFile, readFile, rm } = await import('node:fs/promises');
  const { tmpdir } = await import('node:os');
  const { join } = await import('node:path');
  const { Client } = await import('@modelcontextprotocol/sdk/client/index.js');
  const { InMemoryTransport } = await import('@modelcontextprotocol/sdk/inMemory.js');
  const { createOfficeMcpServer } = await import('../src/mcp');
  const root = await mkdtemp(join(tmpdir(), 'section-edit-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'section-edit-test', version: '1' });
  const [left, right] = InMemoryTransport.createLinkedPair();
  await server.connect(right);
  await client.connect(left);
  const call = async (name: string, args: Record<string, unknown>) => {
    const output = await client.callTool({ name, arguments: args });
    expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
    return output.structuredContent as Record<string, any>;
  };
  try {
    const tools = await client.listTools();
    const schema = JSON.stringify(tools.tools.find(tool => tool.name === 'docx_edit')?.inputSchema);
    expect(schema).toContain('insert_section_break');
    expect(schema).toContain('set_section');
    await writeFile(join(root, 'input.docx'), await sectionPackage());
    const { document } = await call('office_open', { path: 'input.docx' });
    const outline = await call('office_outline', { document });
    await call('docx_edit', { document, version: outline.version, edits: [{ op: 'insert_section_break', before: outline.items[1].ref, kind: 'evenPage' }] });
    const updated = await call('office_outline', { document, headingsOnly: true, sectionOffset: 1, sectionLimit: 1 });
    expect(updated.items).toHaveLength(0);
    expect(updated.sections.items[0].number).toBe(2);
    await call('docx_edit', { document, version: updated.version, edits: [{ op: 'set_section', section: 2, orientation: 'landscape', margins: { left: 900, right: 900 } }] });
    const stale = await client.callTool({ name: 'docx_edit', arguments: { document, version: outline.version, edits: [{ op: 'set_section', section: 1, pageWidth: 12000 }] } });
    expect(stale.structuredContent).toMatchObject({ code: 'STALE_VERSION' });
    await call('office_export', { document, path: 'output.docx' });
    const reopened = await openDocx(await readFile(join(root, 'output.docx')));
    try { expect(reopened.list().sections.items[1]).toMatchObject({ orientation: 'landscape', margins: { left: 900, right: 900 } }); }
    finally { reopened.close(); }
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});

test('a table-cell section target refuses an entire agent batch', async () => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:body>', '<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'));
  const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const outline = doc.list({ limit: 100 });
    const target = outline.items.find(item => item.table)!;
    expect(target).toBeDefined();
    const before = doc.session.encodeState();
    await expect(doc.edit({ version: outline.version, edits: [
      { op: 'set_section', section: 1, orientation: 'landscape' },
      { op: 'insert_section_break', after: target.ref, kind: 'continuous' },
    ] })).rejects.toMatchObject({ code: 'UNSUPPORTED_TARGET' });
    expect(doc.session.encodeState()).toEqual(before);
  } finally { doc.close(); }
});

test('a break before a paragraph that follows a table refuses instead of landing ahead of the table', async () => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:p w14:paraId="10000002">', '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="10000002">'));
  const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const outline = doc.list({ limit: 100 });
    const second = outline.items.find(item => item.text.endsWith('Second'))!;
    expect(second).toBeDefined();
    const before = doc.session.encodeState();
    await expect(doc.edit({ version: outline.version, edits: [{ op: 'insert_section_break', before: second.ref, kind: 'nextPage' }] })).rejects.toMatchObject({ code: 'UNSUPPORTED_TARGET' });
    expect(doc.session.encodeState()).toEqual(before);
    expect(doc.list().sections.total).toBe(1);
  } finally { doc.close(); }
});
