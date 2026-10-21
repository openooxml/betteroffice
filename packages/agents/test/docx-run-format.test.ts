import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const run = (text: string, properties = '') => `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
async function formattedFixture(content: string, properties = '', styles = '') {
  const zip = await JSZip.loadAsync(await editFixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/,
    `<w:body><w:p><w:pPr>${properties}</w:pPr>${content}</w:p><w:p>${run('Tail')}</w:p>`));
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${styles}</w:styles>`));
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
async function xmlOf(doc: DocxAgentDocument) {
  return (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
}
const paragraphs = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const paragraphOf = (xml: string, text: string) => paragraphs(xml).find(p => p.includes(`>${text}</w:t>`))!;
const propertiesOf = (paragraph: string) => [...paragraph.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/g)].map(match => match[0]).join('');
async function insert(doc: DocxAgentDocument, texts = ['Inserted'], style?: string) {
  return doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', after: doc.list().items[0].ref, runFormat: 'anchor',
    paragraphs: texts.map(text => ({ text, ...(style && { style }) })),
  }] });
}

test('runFormat anchor copies direct italic and italicCs through export and reopen', async () => {
  const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>')));
  await insert(doc);
  expect(doc.read(item(doc, 'Inserted').ref).runs[0].formatting.italic).toBe(true);
  const saved = await xmlOf(doc);
  const properties = propertiesOf(paragraphOf(saved, 'Inserted'));
  expect(properties).toContain('<w:i/>');
  expect(properties).toContain('<w:iCs/>');
  const reopened = await open(await doc.export());
  expect(doc.read(item(doc, 'Inserted').ref).runs).toEqual(reopened.read(item(reopened, 'Inserted').ref).runs);
  expect(propertiesOf(paragraphOf(await xmlOf(reopened), 'Inserted'))).toBe(properties);
});

test('runFormat anchor uses the first run instead of the last run', async () => {
  const doc = await open(await formattedFixture(run('First ', '<w:i/><w:iCs/>') + run('last', '<w:b/><w:bCs/>')));
  await insert(doc);
  const properties = propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'));
  expect(properties).toContain('<w:i/>');
  expect(properties).toContain('<w:iCs/>');
  expect(properties).not.toContain('<w:b/>');
  expect(properties).not.toContain('<w:bCs/>');
});

test('runFormat anchor keeps first-run complex-script formatting when live runs coalesce', async () => {
  const doc = await open(await formattedFixture(run('First ', '<w:i/><w:iCs/>') + run('last', '<w:i/><w:iCs w:val="0"/>')));
  await insert(doc);
  expect(propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'))).toContain('<w:iCs/>');
});

test('runFormat anchor skips a leading hyperlink', async () => {
  const linked = `<w:hyperlink w:anchor="target">${run('Linked ', '<w:b/><w:bCs/>')}</w:hyperlink>`;
  const doc = await open(await formattedFixture(linked + run('Plain', '<w:i/><w:iCs/>')));
  await insert(doc);
  const paragraph = paragraphOf(await xmlOf(doc), 'Inserted');
  expect(propertiesOf(paragraph)).toContain('<w:iCs/>');
  expect(propertiesOf(paragraph)).toContain('<w:i/>');
  expect(propertiesOf(paragraph)).not.toContain('<w:b/>');
  expect(paragraph).not.toContain('<w:hyperlink');
});

test('runFormat anchor skips revisions and fields before plain text', async () => {
  const content = `<w:ins w:id="1">${run('Revision ', '<w:b/>')}</w:ins>` +
    `<w:fldSimple w:instr="PAGE">${run('1', '<w:b/>')}</w:fldSimple>` + run('Plain', '<w:i/><w:iCs/>');
  const doc = await open(await formattedFixture(content));
  await insert(doc);
  const paragraph = paragraphOf(await xmlOf(doc), 'Inserted');
  expect(propertiesOf(paragraph)).toContain('<w:i/>');
  expect(propertiesOf(paragraph)).toContain('<w:iCs/>');
  expect(paragraph).not.toMatch(/<w:(ins|del|fldSimple)\b/);
});

test('runFormat anchor falls back to the selected style for an empty or hyperlink-only anchor', async () => {
  for (const content of ['', `<w:hyperlink w:anchor="target">${run('Linked', '<w:b/>')}</w:hyperlink>`]) {
    const doc = await open(await formattedFixture(content, '<w:rPr><w:b/></w:rPr>'));
    await insert(doc, ['Inserted'], 'Quote');
    const inserted = doc.read(item(doc, 'Inserted').ref);
    expect(inserted.style).toBe('Quote');
    expect(inserted.runs[0].formatting.italic).toBe(true);
    expect(inserted.runs[0].formatting.bold).toBeUndefined();
  }
});

test('runFormat anchor overrides an explicit paragraph style run formatting', async () => {
  const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>')));
  await insert(doc, ['Inserted'], 'Heading1');
  const inserted = doc.read(item(doc, 'Inserted').ref);
  expect(inserted.style).toBe('Heading1');
  expect(inserted.runs[0].formatting.italic).toBe(true);
  expect(inserted.runs[0].formatting.bold).toBeUndefined();
  const properties = propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'));
  expect(properties).toContain('<w:iCs/>');
  expect(properties).toMatch(/<w:b w:val="(?:0|false|off)"\/>/);
});

test('runFormat anchor resolves inherited complex-script marks and preserves direct false values', async () => {
  const styles = '<w:style w:type="paragraph" w:styleId="Script"><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:iCs/><w:bCs/></w:rPr></w:style>';
  for (const direct of ['', '<w:iCs w:val="0"/><w:bCs w:val="0"/>']) {
    const doc = await open(await formattedFixture(run('Anchor', direct), '<w:pStyle w:val="Script"/>', styles));
    await insert(doc, ['Inserted'], 'Script');
    const properties = propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'));
    if (direct) {
      expect(properties).toMatch(/<w:iCs w:val="(?:0|false|off)"\/>/);
      expect(properties).toMatch(/<w:bCs w:val="(?:0|false|off)"\/>/);
    } else {
      const inserted = await open(await doc.export());
      await insert(inserted, ['Inherited']);
      expect(propertiesOf(paragraphOf(await xmlOf(inserted), 'Inherited'))).toContain('<w:iCs/>');
    }
  }
});

test('runFormat anchor inserts multiple paragraphs as one undo step', async () => {
  const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>')));
  const receipt = await insert(doc, ['One', '', 'Two']);
  expect(receipt.results[0].refs?.map(ref => doc.read(ref!).text)).toEqual(['One', '', 'Two']);
  for (const text of ['One', 'Two']) expect(propertiesOf(paragraphOf(await xmlOf(doc), text))).toContain('<w:iCs/>');
  expect(doc.session.undo()).toBe(true);
  expect(doc.list().items.map(entry => entry.text)).toEqual(['Anchor', 'Tail']);
  expect(doc.session.redo()).toBe(true);
  expect(propertiesOf(paragraphOf(await xmlOf(doc), 'Two'))).toContain('<w:iCs/>');
});

test('runFormat anchor rolls back insertion when a later draft edit fails', async () => {
  const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>')));
  const version = doc.list().version;
  const anchor = item(doc, 'Anchor').ref;
  const baseline = await xmlOf(doc);
  await expect(doc.edit({ version, edits: [
    { op: 'insert_paragraphs', after: anchor, runFormat: 'anchor', paragraphs: [{ text: 'Inserted' }] },
    { op: 'delete_paragraph', paragraph: anchor },
    { op: 'set_paragraph_style', paragraph: anchor, style: 'Quote' },
  ] })).rejects.toMatchObject({ code: 'UNKNOWN_REF', details: { failedIndex: 2 } });
  expect(doc.list().version).toBe(version);
  expect(await xmlOf(doc)).toBe(baseline);
});

test('runFormat style and omitted runFormat keep default paragraph formatting', async () => {
  for (const runFormat of [undefined, 'style'] as const) {
    const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>'), '<w:pStyle w:val="Heading1"/>'));
    await doc.edit({ version: doc.list().version, edits: [{
      op: 'insert_paragraphs', after: item(doc, 'Anchor').ref, ...(runFormat && { runFormat }), paragraphs: [{ text: 'Inserted' }],
    }] });
    const inserted = doc.read(item(doc, 'Inserted').ref);
    expect([null, 'Normal']).toContain(inserted.style);
    expect(inserted.runs[0].formatting.italic).toBeUndefined();
    expect(inserted.runs[0].formatting.bold).toBeUndefined();
    expect(propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'))).not.toContain('<w:iCs/>');
  }
});

test('runFormat cannot be combined with list continue', async () => {
  const doc = await open(await formattedFixture(run('Anchor')));
  for (const runFormat of ['style', 'anchor'] as const) {
    await expect(doc.edit({ version: doc.list().version, edits: [{
      op: 'insert_paragraphs', after: item(doc, 'Anchor').ref, runFormat, list: 'continue', paragraphs: [{ text: 'Inserted' }],
    }] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  }
});

test('format_text sets and clears complex-script twins only for requested marks', async () => {
  const doc = await open(await formattedFixture(run('Anchor')));
  for (const value of [true, false]) {
    for (const mark of ['italic', 'bold'] as const) {
      const match = doc.grep({ query: 'Anchor' });
      await doc.edit({ version: match.version, edits: [{ op: 'format_text', match: match.matches[0].match, [mark]: value }] });
      const segment = doc.session.storySegments('body').find(entry => entry.kind === 'text' && entry.text === 'Anchor')!;
      expect(segment.attributes[`${mark}Cs`]).toBe(value);
      const properties = propertiesOf(paragraphOf(await xmlOf(doc), 'Anchor'));
      if (value) expect(properties).toContain(mark === 'italic' ? '<w:iCs/>' : '<w:bCs/>');
      else expect(properties).not.toContain(mark === 'italic' ? '<w:iCs/>' : '<w:bCs/>');
    }
  }
});

test('runFormat anchor reads formatting applied earlier in the same batch', async () => {
  const doc = await open(await formattedFixture(run('Anchor', '<w:i/><w:iCs/>')));
  const match = doc.grep({ query: 'Anchor' });
  await doc.edit({ version: match.version, edits: [
    { op: 'format_text', match: match.matches[0].match, italic: false, bold: true },
    { op: 'insert_paragraphs', after: match.matches[0].ref, runFormat: 'anchor', paragraphs: [{ text: 'Inserted' }] },
  ] });
  const properties = propertiesOf(paragraphOf(await xmlOf(doc), 'Inserted'));
  expect(properties).toContain('<w:b/>');
  expect(properties).toContain('<w:bCs/>');
  expect(properties).not.toContain('<w:i/>');
  expect(properties).not.toContain('<w:iCs/>');
});
