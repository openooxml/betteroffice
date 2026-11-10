import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SMART_TAG = '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Place</w:t></w:r></w:smartTag>';
const TARGET = `<w:p><w:r><w:rPr><w:u w:val="single" w:color="FF0000"/></w:rPr><w:t>Target</w:t></w:r>${SMART_TAG}</w:p>`;

async function docx(body: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}
const opened: DocxAgentDocument[] = [];
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

test('verify reports a paragraph saved as it was because formatting it would drop a smart tag', async () => {
  const doc = await openDocx(await docx(`<w:p><w:r><w:t>Before</w:t></w:r></w:p>${TARGET}`));
  opened.push(doc);
  const match = doc.grep({ query: 'Target' });
  await doc.edit({ version: match.version, edits: [{ op: 'format_text', match: match.matches[0]!.match, underline: false }] });
  const kept = (await doc.verify()).keptUnchanged!;
  expect(kept.structure).toBe(1);
  expect(kept.paragraphs).toHaveLength(1);
  expect(kept.paragraphs[0]!.ref).toEqual(expect.any(String));
  expect(kept.paragraphs[0]!.scope).toBe('paragraph');
  expect(kept.paragraphs[0]!.dropped.length).toBeGreaterThan(0);
  const xml = await (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
  expect(xml).toContain(TARGET);
});

test('verify reports nothing kept when the formatting is written', async () => {
  const doc = await openDocx(await docx(`<w:p><w:r><w:t>Target</w:t></w:r>${SMART_TAG}</w:p>`));
  opened.push(doc);
  const match = doc.grep({ query: 'Target' });
  await doc.edit({ version: match.version, edits: [{ op: 'format_text', match: match.matches[0]!.match, bold: true }] });
  expect((await doc.verify()).keptUnchanged).toBeUndefined();
});

test('setting the language of a paragraph reports the runs it kept', async () => {
  const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const doc = await openDocx(await docx(`<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:hyperlink w:anchor="target"><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Target</w:t></w:r>${field}</w:hyperlink></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>`));
  opened.push(doc);
  const match = doc.grep({ query: 'Target' });
  const edited = await doc.edit({ version: match.version, edits: [{ op: 'set_language', paragraph: match.matches[0]!.ref, lang: 'de-DE' }] });
  expect(edited.results[0]).toMatchObject({ keptUnchanged: { field: 1 } });
});

