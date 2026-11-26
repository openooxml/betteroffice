import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const GRID_ACCENT = '<w:style w:type="table" w:styleId="GridAccent"><w:name w:val="Grid Accent"/><w:tblPr><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>`).join('') +
  '</w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr></w:style>';

async function styledFixture() {
  const zip = await JSZip.loadAsync(await editFixture());
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${GRID_ACCENT}</w:styles>`));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const opened: DocxAgentDocument[] = [];
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const item = (doc: DocxAgentDocument, text: string) => doc.list({ limit: 100 }).items.find(entry => entry.text === text)!;
const documentXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const blocks = (xml: string) => xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>|<w:p\b[^>]*\/>/g) ?? [];
const textOf = (block: string) => [...block.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('|');
const tables = (xml: string) => xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? [];

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the edit to be refused.');
}

test('insert_table adds a bordered full-width table after a paragraph and returns its cell refs', async () => {
  const doc = await open(await editFixture());
  const receipt = await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'insert_table', after: item(doc, 'Revenue grew strongly in the third quarter.').ref, rows: [['Region', 'Q1', 'Q2'], ['North', '1.2M'], ['South', '', '0.8M']] }],
  });
  const refs = receipt.results[0].refs!;
  expect(refs.map(ref => doc.read(ref!).text)).toEqual(['Region', 'Q1', 'Q2', 'North', '1.2M', '', 'South', '', '0.8M']);
  expect(item(doc, 'North').table).toMatchObject({ row: 2, column: 1, rows: 3, columns: 3 });

  const bytes = await doc.export();
  const xml = await documentXml(bytes);
  expect(blocks(xml).map(textOf)).toEqual([
    'Quarterly report', 'Revenue grew strongly in the third quarter.', 'Region|Q1|Q2|North|1.2M|South|0.8M',
    'Obsolete remark.', 'Quarter|Revenue|Q1|1.0M', 'Closing note.',
  ]);
  const table = tables(xml)[0]!;
  expect(table).toContain('<w:tblW w:w="9360" w:type="dxa"/>');
  expect(table).toContain('<w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>');
  expect(table).not.toContain('<w:tblStyle');
  expect(table.match(/<w:tc>/g)).toHaveLength(9);
  expect(table.match(/<w:tcBorders>/g)).toHaveLength(9);
  expect(table).toContain('<w:top w:val="single" w:sz="4"');

  const verified = await doc.verify();
  expect(verified).toMatchObject({ reopened: true, checks: ['save', 'reopen', 'structure'] });
  const reopened = await open(bytes);
  expect(item(reopened, '0.8M').table).toMatchObject({ row: 3, column: 3, rows: 3, columns: 3 });
}, 30_000);

test('insert_table before a following table keeps a paragraph between the two tables', async () => {
  const doc = await open(await editFixture());
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_table', after: item(doc, 'Obsolete remark.').ref, rows: [['New']] }] });
  const bytes = await doc.export();
  expect(blocks(await documentXml(bytes)).map(textOf)).toEqual([
    'Quarterly report', 'Revenue grew strongly in the third quarter.', 'Obsolete remark.', 'New', '', 'Quarter|Revenue|Q1|1.0M', 'Closing note.',
  ]);
  const reopened = await open(bytes);
  expect(item(reopened, 'New').table).toMatchObject({ rows: 1, columns: 1 });
  expect(item(reopened, 'Q1').table).toMatchObject({ index: 2 });
}, 30_000);

test('insert_table applies a table style by name with its header row look', async () => {
  const doc = await open(await styledFixture());
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'insert_table', after: item(doc, 'Revenue grew strongly in the third quarter.').ref, rows: [['Name', 'Role'], ['Ada', 'Engineer']], style: 'Grid Accent' }],
  });
  const bytes = await doc.export();
  const table = tables(await documentXml(bytes))[0]!;
  expect(table).toContain('<w:tblStyle w:val="GridAccent"/>');
  expect(table).toContain('<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>');
  expect(table).not.toContain('<w:tcBorders>');
  expect((await doc.verify()).reopened).toBe(true);

  const reopened = await open(bytes);
  const header = reopened.session.storySegments(item(reopened, 'Name').story!)
    .find(segment => segment.kind === 'text' && segment.text === 'Name')!;
  expect(header.attributes).toMatchObject({ bold: true });
}, 30_000);

test('insert_table refuses a table style the document does not define and leaves the document unchanged', async () => {
  const doc = await open(await styledFixture());
  const version = doc.list().version;
  const error = await refusal(doc.edit({
    version,
    edits: [{ op: 'insert_table', after: item(doc, 'Revenue grew strongly in the third quarter.').ref, rows: [['x']], style: 'Missing Grid' }],
  }));
  expect(error.code).toBe('UNKNOWN_STYLE');
  expect(error.message).toContain('Table style "Missing Grid" is not defined');
  expect(error.details).toMatchObject({ validStyles: expect.arrayContaining(['GridAccent']) });
  expect(doc.list().version).toBe(version);
  expect(tables(await documentXml(await doc.export()))).toHaveLength(1);
}, 30_000);

test('insert_table refs can be targeted by a following insert_table_row', async () => {
  const doc = await open(await editFixture());
  const receipt = await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'insert_table', after: item(doc, 'Quarterly report').ref, rows: [['A', 'B'], ['1', '2']] }],
  });
  const last = receipt.results[0].refs!.at(-1)!;
  await doc.edit({ version: receipt.version, edits: [{ op: 'insert_table_row', row: last, position: 'after', cells: ['3', '4'] }] });
  const bytes = await doc.export();
  expect(textOf(tables(await documentXml(bytes))[0]!)).toBe('A|B|1|2|3|4');
  const reopened = await open(bytes);
  expect(item(reopened, '4').table).toMatchObject({ row: 3, column: 2, rows: 3, columns: 2 });
}, 30_000);

test('insert_table refuses an anchor inside a table cell', async () => {
  const doc = await open(await editFixture());
  const error = await refusal(doc.edit({ version: doc.list().version, edits: [{ op: 'insert_table', after: item(doc, 'Q1').ref, rows: [['x']] }] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
}, 30_000);

test('insert_table after a paragraph that follows a table keeps that paragraph whole', async () => {
  const doc = await open(await editFixture());
  const closing = doc.list({ limit: 100 }).items.find(entry => entry.text.endsWith('Closing note.'))!;
  const receipt = await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_table', after: closing.ref, rows: [['Last']] }] });
  expect(doc.read(closing.ref).text.endsWith('Closing note.')).toBe(true);
  expect(doc.read(receipt.results[0].refs![0]!).text).toBe('Last');
  const bytes = await doc.export();
  expect(blocks(await documentXml(bytes)).map(textOf).slice(-4)).toEqual(['Quarter|Revenue|Q1|1.0M', 'Closing note.', 'Last', '']);
  expect((await doc.verify()).reopened).toBe(true);
}, 30_000);

test('insert_table takes the text width a set_section earlier in the batch leaves', async () => {
  const doc = await open(await editFixture());
  await doc.edit({
    version: doc.list().version,
    edits: [
      { op: 'set_section', section: 1, margins: { left: 720, right: 720 } },
      { op: 'insert_table', after: item(doc, 'Quarterly report').ref, rows: [['a', 'b']] },
    ],
  });
  const table = tables(await documentXml(await doc.export()))[0]!;
  expect(table).toContain('<w:tblW w:w="10800" w:type="dxa"/>');
  expect(table).toContain('<w:gridCol w:w="5400"/><w:gridCol w:w="5400"/>');
}, 30_000);
