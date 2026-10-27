import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx/yrs';
import { openDocx, type DocxAgentDocument } from '../src';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const THEMED = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>';
const EXPLICIT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
const THEME_DEFAULTS = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:cstheme="minorHAnsi"/>';
const PARTIAL_DEFAULTS = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Times New Roman" w:cs="Arial"/>';
const PARTIAL = '<w:rFonts w:cs="Arial"/><w:lang w:eastAsia="en-GB"/>';
const run = (text: string, properties: string) => `<w:r><w:rPr>${properties}</w:rPr><w:t>${text}</w:t></w:r>`;
const paragraph = (text: string, properties: string, ppr = '') => `<w:p><w:pPr>${ppr}<w:rPr>${properties}</w:rPr></w:pPr>${run(text, properties)}</w:p>`;
const TAIL = '<w:p w:rsidR="12345678"><w:r><w:t>Untouched tail.</w:t></w:r></w:p>';

async function fixture(body: string, defaults = THEME_DEFAULTS) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'numbering', 'theme'].map(name => `<Relationship Id="${name}" Type="${R}/${name}" Target="${name === 'theme' ? 'theme/theme1' : name}.xml"/>`).join('')}</Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}${TAIL}<w:sectPr/></w:body></w:document>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${defaults}<w:lang w:val="en-US" w:eastAsia="en-US"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`);
  zip.file('word/theme/theme1.xml', `<a:theme xmlns:a="${A}"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  zip.file('word/numbering.xml', `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const opened: DocxAgentDocument[] = [];
afterEach(() => opened.splice(0).forEach(doc => doc.close()));
async function open(body: string, defaults = THEME_DEFAULTS) {
  const doc = await openDocx(await fixture(body, defaults));
  opened.push(doc);
  return doc;
}
async function xmlOf(doc: DocxAgentDocument) {
  return (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
}
function fonts(xml: string) {
  return (xml.match(/<w:rFonts\b[^>]*\/>/g) ?? []).map(tag =>
    Object.fromEntries([...tag.matchAll(/w:(\w+)="([^"]*)"/g)].map(match => [match[1], match[2]])));
}
async function comment(doc: DocxAgentDocument) {
  const found = doc.grep({ query: 'middle' });
  await doc.edit({ version: doc.list().version, edits: [{ op: 'add_comment', match: found.matches[0]!.match, author: 'Reader', text: 'Check' }] });
}
function expectFonts(xml: string, source: string) {
  const values = fonts(xml);
  expect(values.length).toBeGreaterThan(0);
  values.forEach(value => expect(value).toEqual(fonts(source)[0]));
  expect(xml).toContain(TAIL);
}

test('comment splits keep authored theme fonts on runs and paragraph mark', async () => {
  const doc = await open(paragraph('Before middle after', THEMED));
  await comment(doc);
  expectFonts(await xmlOf(doc), THEMED);
});

test('list continuation copies only the neighbour authored font slots', async () => {
  const doc = await open(paragraph('Before middle after', THEMED, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'));
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: doc.list().items[0]!.ref, list: 'continue', paragraphs: [{ text: 'Inserted' }] }] });
  expectFonts(await xmlOf(doc), THEMED);
});

test('list continuation copies authored paragraph mark fonts from an empty last item', async () => {
  const numbering = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
  const doc = await open(`${paragraph('First item', THEMED, numbering)}<w:p><w:pPr>${numbering}<w:rPr>${THEMED}</w:rPr></w:pPr></w:p>`);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: doc.list().items[1]!.ref, list: 'continue', paragraphs: [{ text: 'Inserted' }] }] });
  const xml = await xmlOf(doc);
  const inserted = (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(value => value.includes('Inserted'))!;
  expect(fonts(inserted).length).toBeGreaterThan(0);
  fonts(inserted).forEach(value => expect(value).toEqual(fonts(THEMED)[0]));
  expectFonts(xml, THEMED);
});

test('list continuation copies the plain run fonts before a complex hyperlink field', async () => {
  const plainFonts = '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>';
  const hyperlinkFonts = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>';
  const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('Example', `<w:rStyle w:val="Hyperlink"/>${hyperlinkFonts}`)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const doc = await open(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>${run('See ', plainFonts)}${field}</w:p>`);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: doc.list().items[0]!.ref, list: 'continue', paragraphs: [{ text: 'Inserted' }] }] });
  const xml = await xmlOf(doc);
  const inserted = (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(value => value.includes('Inserted'))!;
  expect(fonts(inserted).length).toBeGreaterThan(0);
  fonts(inserted).forEach(value => expect(value).toEqual(fonts(plainFonts)[0]));
  expect(xml).toContain(TAIL);
});

test.each([false, true])('merged cell runs keep explicit fonts without extra script slots with comment rewrite %s', async (rewrite) => {
  const cells = ['Left middle end', 'Right'].map(text => `<w:tc><w:tcPr/>${paragraph(text, EXPLICIT)}</w:tc>`).join('');
  const doc = await open(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`);
  const items = doc.list().items.filter(item => item.table);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'merge_cells', from: items[0]!.ref, to: items[1]!.ref }] });
  if (rewrite) await comment(doc);
  expectFonts(await xmlOf(doc), EXPLICIT);
});

test('section and comment edits keep partial fonts and source language', async () => {
  const doc = await open(paragraph('Before middle after', PARTIAL,
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'), PARTIAL_DEFAULTS);
  const added = await doc.edit({ version: doc.list().version, edits: [{
    op: 'add_comment', match: doc.grep({ query: 'middle' }).matches[0]!.match, author: 'Reader', text: 'Check',
  }] });
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: doc.list().items[0]!.ref, list: 'continue', paragraphs: [{ text: 'Inserted' }] }] });
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_section_break', after: doc.list().items.find(item => item.text === 'Inserted')!.ref, kind: 'nextPage' }] });
  await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: added.results[0]!.comment!, author: 'Reader', text: 'Follow up' }] });
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_section', section: 1, orientation: 'landscape', pageWidth: 16838, pageHeight: 11906 }] });
  const xml = await xmlOf(doc);
  expectFonts(xml, PARTIAL);
  const languages = xml.match(/<w:lang\b[^>]*\/>/g) ?? [];
  expect(languages.length).toBeGreaterThan(1);
  languages.forEach(language => expect(language).toBe('<w:lang w:eastAsia="en-GB"/>'));
});

test('list continuation encoded without authored boundaries recovers its source neighbour on save', async () => {
  const body = paragraph('Before middle after', THEMED, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
  const bytes = await fixture(body);
  const doc = await openDocx(bytes);
  opened.push(doc);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: doc.list().items[0]!.ref, list: 'continue', paragraphs: [{ text: 'Inserted' }] }] });
  const source = doc.session.paragraphs('body').find(item => item.text === 'Before middle after')!;
  const inserted = doc.session.paragraphs('body').find(item => item.text === 'Inserted')!;
  const anchor = doc.session.storySegments('body').find(segment => segment.kind === 'pilcrow' && segment.paraId === source.paraId)!;
  if (anchor.kind !== 'pilcrow') throw new Error('missing anchor');
  doc.session.setParagraphAttr(inserted.paraId, '_originalRunBoundaries', null);
  doc.session.setParagraphAttr(inserted.paraId, '_originalFormatting', anchor.properties._originalFormatting);
  const session = await createYrsSession();
  try {
    session.openDocx(bytes, false);
    session.loadState(doc.session.encodeState());
    const xml = (await JSZip.loadAsync((await saveYrsDocx(session)).bytes)).file('word/document.xml')!;
    expectFonts(await xml.async('string'), THEMED);
  } finally { session.destroy(); }
});
