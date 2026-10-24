import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { createOfficeMcpServer } from '../src/mcp';
import { editFixture } from './fixture';

const cell = (text: string, rPr = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${text ? `<w:r>${rPr && `<w:rPr>${rPr}</w:rPr>`}<w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const table = (columns: number, rows: string[][]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="${columns * 3000}" w:type="dxa"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(columns)}</w:tblGrid>${rows.map(row => `<w:tr>${row.join('')}</w:tr>`).join('')}</w:tbl>`;

async function mergeFixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await editFixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  const body = [
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>',
    '<w:p><w:r><w:t xml:space="preserve">Revenue grew </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>strongly</w:t></w:r></w:p>',
    table(3, [
      [cell('Region', '<w:b/>'), cell(''), cell('Total', '<w:b/>')],
      [cell('North'), cell('Units', '<w:u w:val="single"/>'), cell('5')],
      [cell('South'), cell('Note A', '<w:i/><w:color w:val="C00000"/>'), cell('7')],
      [cell('Group heading', '<w:b/>'), cell(''), cell('')],
      [cell('Sum'), cell('12'), cell('13')],
    ]),
    '<w:p><w:r><w:t>Between tables.</w:t></w:r></w:p>',
    table(2, [[cell('Name', '<w:b/>'), cell('')], [cell('Alpha'), cell('Beta')]]),
    '<w:p><w:r><w:t>Closing note.</w:t></w:r></w:p>',
  ].join('');
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${body}`));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const opened: DocxAgentDocument[] = [];
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await mergeFixture());
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const at = (doc: DocxAgentDocument, index: number, row: number, column: number) =>
  doc.list({ limit: 100 }).items.find(item => item.table?.index === index && item.table.row === row && item.table.column === column)!.ref;
const documentXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const tables = (xml: string) => (xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? []).map(found => found.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) ?? []);
const cells = (row: string) => row.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];

async function paragraphs(bytes: Uint8Array) {
  const doc = await open(bytes);
  return doc.list({ limit: 100 }).items.map(item => {
    const { text, style, runs, table } = doc.read(item.ref);
    return { text, style, runs, table };
  });
}
type Snapshot = Awaited<ReturnType<typeof paragraphs>>;
const content = (items: Snapshot) => items.map(({ text, style, runs }) => ({ text, style, runs }));

async function edited(doc: DocxAgentDocument, edits: DocxAgentEdit[]) {
  const baseline = await doc.export();
  const receipt = await doc.edit({ version: doc.list().version, edits });
  const bytes = await doc.export();
  return { receipt, bytes, xml: await documentXml(bytes), baselineXml: await documentXml(baseline), before: await paragraphs(baseline), after: await paragraphs(bytes) };
}

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the batch to be refused.');
}

test('merge_cells joins a header cell with the empty cell to its right, one merge per table', async () => {
  const doc = await open();
  const region = at(doc, 1, 1, 1);
  const blank = at(doc, 1, 1, 2);
  const total = at(doc, 1, 1, 3);
  const { receipt, before, after } = await edited(doc, [
    { op: 'merge_cells', from: region, to: blank },
    { op: 'merge_cells', from: at(doc, 2, 1, 1), to: at(doc, 2, 1, 2) },
  ]);
  const [first, second] = receipt.results;
  expect(first).toMatchObject({ index: 0, op: 'merge_cells', rowSpan: 1, colSpan: 2 });
  expect(first.refs![0]).toBe(region);
  expect(first.refs!.map(ref => doc.read(ref!).text)).toEqual(['Region', '']);
  expect(second.refs!.map(ref => doc.read(ref!).text)).toEqual(['Name', '']);
  expect(() => doc.read(blank)).toThrow('Paragraph missing');
  expect(doc.read(region).table).toEqual({ story: 'body', index: 1, row: 1, column: 1, rows: 5, columns: 3, colSpan: 2 });
  expect(doc.read(total)).toMatchObject({ text: 'Total', table: { story: 'body', index: 1, row: 1, column: 3, rows: 5, columns: 3 } });
  expect(doc.read(total).table).not.toHaveProperty('colSpan');
  expect(doc.read(second.refs![0]!).table).toEqual({ story: 'body', index: 2, row: 1, column: 1, rows: 2, columns: 2, colSpan: 2 });
  expect(content(after)).toEqual(content(before));
});

test('merge_cells merges two filled cells of a column, keeping both cells’ paragraphs in row order', async () => {
  const doc = await open();
  const units = at(doc, 1, 2, 2);
  const note = at(doc, 1, 3, 2);
  const { receipt, before, after } = await edited(doc, [{ op: 'merge_cells', from: note, to: units }]);
  const [result] = receipt.results;
  expect(result).toMatchObject({ index: 0, op: 'merge_cells', rowSpan: 2, colSpan: 1 });
  expect(result.refs!.map(ref => doc.read(ref!).text)).toEqual(['Units', 'Note A']);
  expect(result.refs![0]).toBe(units);
  expect(() => doc.read(note)).toThrow('Paragraph missing');
  const merged = after.filter(p => p.table?.index === 1 && p.table.row === 2 && p.table.column === 2);
  expect(merged.map(p => p.text)).toEqual(['Units', 'Note A']);
  expect(merged.map(p => p.table)).toEqual([0, 1].map(() => ({ story: 'body', index: 1, row: 2, column: 2, rows: 5, columns: 3, rowSpan: 2 })));
  expect(after.find(p => p.text === 'Note A')!.runs).toEqual(before.find(p => p.text === 'Note A')!.runs);
  expect(after.find(p => p.text === 'Note A')!.runs[0].formatting).toMatchObject({ italic: true });
  expect(content(after.filter(p => p.text !== 'Note A'))).toEqual(content(before.filter(p => p.text !== 'Note A')));
});

test('merge_cells turns a group-heading row into one full-width cell', async () => {
  const doc = await open();
  const heading = at(doc, 1, 4, 1);
  const { receipt, before, after } = await edited(doc, [{ op: 'merge_cells', from: heading, to: at(doc, 1, 4, 3) }]);
  const [result] = receipt.results;
  expect(result).toMatchObject({ index: 0, op: 'merge_cells', rowSpan: 1, colSpan: 3 });
  expect(result.refs!.map(ref => doc.read(ref!).text)).toEqual(['Group heading', '', '']);
  expect(doc.read(heading).table).toEqual({ story: 'body', index: 1, row: 4, column: 1, rows: 5, columns: 3, colSpan: 3 });
  expect(after.filter(p => p.table?.row === 4).map(p => p.table?.colSpan)).toEqual([3, 3, 3]);
  expect(content(after)).toEqual(content(before));
});

test('merged cells save as gridSpan and vMerge and reopen without touching other content', async () => {
  const doc = await open();
  const { xml, baselineXml, before, after } = await edited(doc, [
    { op: 'merge_cells', from: at(doc, 1, 1, 1), to: at(doc, 1, 1, 2) },
    { op: 'merge_cells', from: at(doc, 1, 2, 2), to: at(doc, 1, 3, 2) },
    { op: 'merge_cells', from: at(doc, 1, 4, 1), to: at(doc, 1, 4, 3) },
    { op: 'merge_cells', from: at(doc, 2, 1, 1), to: at(doc, 2, 1, 2) },
  ]);
  const [first, second] = tables(xml);
  const [firstBefore, secondBefore] = tables(baselineXml);
  expect(cells(first[0])).toHaveLength(2);
  expect(cells(first[0])[0]).toContain('<w:gridSpan w:val="2"/>');
  expect(cells(first[1])[1]).toMatch(/<w:vMerge w:val="restart"\/>/);
  expect(cells(first[2])).toHaveLength(3);
  expect(cells(first[2])[1]).toMatch(/<w:vMerge(?: w:val="continue")?\/>/);
  expect(cells(first[2])[1]).not.toContain('Note A');
  expect(cells(first[3])).toHaveLength(1);
  expect(cells(first[3])[0]).toContain('<w:gridSpan w:val="3"/>');
  expect(cells(second[0])[0]).toContain('<w:gridSpan w:val="2"/>');
  expect(first[4]).toBe(firstBefore[4]);
  expect(second[1]).toBe(secondBefore[1]);
  const outside = (value: string) => value.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, '').match(/<w:p\b[\s\S]*?<\/w:p>/g);
  expect(outside(xml)).toEqual(outside(baselineXml));
  const spans = after.filter(p => p.table?.rowSpan || p.table?.colSpan).map(p => [p.text, p.table!.row, p.table!.column, p.table!.rowSpan ?? 1, p.table!.colSpan ?? 1]);
  expect(spans).toEqual([
    ['Region', 1, 1, 1, 2], ['', 1, 1, 1, 2],
    ['Units', 2, 2, 2, 1], ['Note A', 2, 2, 2, 1],
    ['Group heading', 4, 1, 1, 3], ['', 4, 1, 1, 3], ['', 4, 1, 1, 3],
    ['Name', 1, 1, 1, 2], ['', 1, 1, 1, 2],
  ]);
  expect(after.find(p => p.text === 'Region')!.runs[0].formatting.bold).toBe(true);
  expect(after.find(p => p.text === 'Units')!.runs).toEqual(before.find(p => p.text === 'Units')!.runs);
  expect(after.find(p => p.text === 'Note A')!.runs).toEqual(before.find(p => p.text === 'Note A')!.runs);
  expect(content(after.filter(p => p.text !== 'Note A'))).toEqual(content(before.filter(p => p.text !== 'Note A')));
});

test('merge_cells refusals name the problem and change nothing', async () => {
  const doc = await open();
  const outline = doc.list({ limit: 100 });
  const baseline = await documentXml(await doc.export());
  const heading = outline.items.find(item => item.text === 'Quarterly report')!.ref;
  const region = at(doc, 1, 1, 1);
  const refused = async (edits: DocxAgentEdit[], code: string, failedIndex: number) => {
    const error = await refusal(doc.edit({ version: outline.version, edits }));
    expect(error.code).toBe(code);
    expect(error.message).toContain('No batch edits were applied.');
    expect(error.details).toMatchObject({ failedIndex });
    expect(doc.list({ limit: 100 })).toEqual(outline);
    expect(await documentXml(await doc.export())).toBe(baseline);
  };
  await refused([{ op: 'merge_cells', from: region, to: at(doc, 2, 1, 2) }], 'UNSUPPORTED_TARGET', 0);
  await refused([{ op: 'merge_cells', from: region, to: region }], 'INVALID_ARGUMENT', 0);
  await refused([{ op: 'set_paragraph_style', paragraph: heading, style: 'Heading2' }, { op: 'merge_cells', from: heading, to: region }], 'NOT_IN_TABLE', 1);
  await refused([{ op: 'merge_cells', from: region, to: 'p999' }], 'UNKNOWN_REF', 0);
  await refused([
    { op: 'merge_cells', from: region, to: at(doc, 1, 1, 2) },
    { op: 'merge_cells', from: at(doc, 1, 1, 2), to: at(doc, 1, 2, 2) },
  ], 'UNKNOWN_REF', 1);
  await refused([
    { op: 'merge_cells', from: region, to: at(doc, 1, 1, 2) },
    { op: 'merge_cells', from: at(doc, 1, 2, 2), to: at(doc, 1, 1, 3) },
  ], 'UNSUPPORTED_TARGET', 1);
});

test('merge_cells rolls back with a failing later edit in the batch', async () => {
  const doc = await open();
  const outline = doc.list({ limit: 100 });
  const baseline = await documentXml(await doc.export());
  const error = await refusal(doc.edit({ version: outline.version, edits: [
    { op: 'merge_cells', from: at(doc, 1, 1, 1), to: at(doc, 1, 1, 2) },
    { op: 'delete_paragraph', paragraph: at(doc, 1, 1, 3) },
  ] }));
  expect(error.code).toBe('LAST_PARAGRAPH');
  expect(error.details).toMatchObject({ failedIndex: 1 });
  expect(doc.list({ limit: 100 })).toEqual(outline);
  expect(await documentXml(await doc.export())).toBe(baseline);
});

test('docx_edit lists merge_cells and the outline shows the merged span', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-merge-mcp-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'docx-merge', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  const call = async (name: string, args: Record<string, unknown>) => {
    const output = await client.callTool({ name, arguments: args });
    expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
    return output.structuredContent as Record<string, any>;
  };
  try {
    const tools = (await client.listTools()).tools;
    const edit = tools.find(item => item.name === 'docx_edit')!;
    expect(JSON.stringify(edit.inputSchema)).toContain('merge_cells');
    expect(edit.description).toContain('merge_cells');
    expect(tools.find(item => item.name === 'office_outline')!.description).toContain('colSpan');
    await writeFile(join(root, 'merge.docx'), await mergeFixture());
    const { document } = await call('office_open', { path: 'merge.docx' });
    const outline = await call('office_outline', { document, limit: 100 });
    const position = (row: number, column: number) => outline.items.find((item: any) => item.table?.index === 1 && item.table.row === row && item.table.column === column).ref;
    const result = await call('docx_edit', { document, version: outline.version, edits: [{ op: 'merge_cells', from: position(4, 1), to: position(4, 3) }] });
    expect(result.results[0]).toMatchObject({ op: 'merge_cells', rowSpan: 1, colSpan: 3 });
    const items = (await call('office_outline', { document, limit: 100 })).items;
    expect(items.find((item: any) => item.text === 'Group heading').table).toMatchObject({ row: 4, column: 1, colSpan: 3 });
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});

for (const commentFirst of [true, false]) {
  test(`merge_cells keeps comment anchors in the merged area with comment first ${commentFirst}`, async () => {
    let doc = await open();
    const outside = doc.grep({ query: 'Closing note.' }).matches[0];
    await doc.edit({ version: doc.list().version, edits: [
      { op: 'add_comment', match: outside.match, author: 'Reviewer', text: 'Keep this note.' },
      { op: 'insert_section_break', after: doc.list().items.find(item => item.text === 'Quarterly report')!.ref, kind: 'nextPage' },
      { op: 'set_section', section: 1, orientation: 'landscape' },
    ] });
    const baseline = await doc.export();
    doc = await open(baseline);
    const baselineZip = await JSZip.loadAsync(baseline);
    const baselineXml = await documentXml(baseline);
    const comment = { op: 'add_comment' as const, match: doc.grep({ query: commentFirst ? 'Note A' : 'Units' }).matches[0].match, author: 'Reviewer', text: 'Check these units.' };
    const merge = { op: 'merge_cells' as const, from: at(doc, 1, 2, 2), to: at(doc, 1, 3, 2) };
    const receipt = await doc.edit({ version: doc.list().version, edits: commentFirst ? [comment, merge] : [merge, comment] });
    const id = receipt.results[commentFirst ? 0 : 1].comment!;
    expect(doc.session.resolveComment(id)).toHaveLength(1);
    const saved = await doc.export();
    const zip = await JSZip.loadAsync(saved);
    const xml = await documentXml(saved);
    expect(tables(xml)[1]).toEqual(tables(baselineXml)[1]);
    const sections = (value: string) => value.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g);
    expect(sections(xml)).toEqual(sections(baselineXml));
    expect(sections(xml)).toHaveLength(2);
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toMatch(/<w:vMerge(?: w:val="continue")?\/>/);
    const existing = (await baselineZip.file('word/comments.xml')!.async('string')).match(/<w:comment\b[\s\S]*?<\/w:comment>/)![0];
    expect(await zip.file('word/comments.xml')!.async('string')).toContain(existing);
    const reopened = await open(saved);
    const savedComment = reopened.list({ story: 'comments' }).items.find(item => item.text === 'Check these units.')!;
    expect(reopened.session.resolveComment(savedComment.id)).toEqual(doc.session.resolveComment(id));
    expect(reopened.list().sections.items).toEqual(doc.list().sections.items);
    expect(reopened.list({ story: 'comments' }).items.map(item => item.text).sort()).toEqual(['Check these units.', 'Keep this note.']);
  });
}
