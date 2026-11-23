import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const NUMBERING = '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>';

const LIST_STYLES =
  '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>';

async function listFixture(numbering: string | null = NUMBERING) {
  const zip = await JSZip.loadAsync(await editFixture());
  if (numbering) {
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>',
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
      '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
    zip.file('word/numbering.xml', numbering);
  }
  const styles = (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${LIST_STYLES}</w:styles>`);
  zip.file('word/styles.xml', styles);
  return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), styles };
}

const opened: DocxAgentDocument[] = [];
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const item = (doc: DocxAgentDocument, text: string) =>
  doc.list({ limit: 100 }).items.find(entry => entry.text === text)!;
const byText = (xml: string, text: string) =>
  (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(block => block.includes(`>${text}<`))!;
const pilcrow = (doc: DocxAgentDocument, text: string) =>
  doc.session.paragraphs('body').find(entry => entry.text === text)!.properties;

test('insert_paragraphs with numbered styles adds bullet and numbered list items numbered by their style', async () => {
  const source = await listFixture();
  const doc = await open(source.bytes);
  await doc.edit({
    version: doc.list().version,
    edits: [{
      op: 'insert_paragraphs',
      after: item(doc, 'Revenue grew strongly in the third quarter.').ref,
      paragraphs: [
        { text: 'First bullet', style: 'ListBullet' },
        { text: 'Second bullet', style: 'List Bullet' },
        { text: 'First step', style: 'ListNumber' },
        { text: 'Second step', style: 'ListNumber' },
      ],
    }],
  });
  for (const text of ['First bullet', 'Second bullet', 'First step', 'Second step']) {
    expect(item(doc, text).list).toEqual({ level: 0 });
  }
  expect(pilcrow(doc, 'First bullet')).toMatchObject({ numPr: { numId: 1 }, listIsBullet: true });
  expect(pilcrow(doc, 'Second step')).toMatchObject({ numPr: { numId: 2 }, listNumFmt: 'decimal' });

  const bytes = await doc.export();
  const saved = await JSZip.loadAsync(bytes);
  const xml = await saved.file('word/document.xml')!.async('string');
  for (const [text, style] of [['First bullet', 'ListBullet'], ['Second step', 'ListNumber']]) {
    const paragraph = byText(xml, text);
    expect(paragraph).toContain(`<w:pStyle w:val="${style}"/>`);
    expect(paragraph).not.toContain('<w:numPr>');
  }
  expect(await saved.file('word/styles.xml')!.async('string')).toBe(source.styles);
  const numbering = await saved.file('word/numbering.xml')!.async('string');
  expect(numbering).toBe(NUMBERING);
  expect(numbering).toContain('<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>');
  expect(numbering).toContain('<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>');

  const reopened = await open(bytes);
  for (const text of ['First bullet', 'Second bullet', 'First step', 'Second step']) {
    expect(item(reopened, text).list).toEqual({ level: 0 });
  }
  expect(reopened.read(item(reopened, 'First bullet').ref).style).toBe('ListBullet');
  expect(pilcrow(reopened, 'Second bullet')).toMatchObject({ numPr: { numId: 1 }, listIsBullet: true });
  expect(pilcrow(reopened, 'Second step')).toMatchObject({ numPr: { numId: 2 }, listNumFmt: 'decimal' });
  expect(item(reopened, 'Obsolete remark.').list).toBeUndefined();
}, 30_000);

test('insert_paragraphs refuses a numbered style whose numbering the document does not define', async () => {
  const source = await listFixture(null);
  const doc = await open(source.bytes);
  const version = doc.list().version;
  let error: DocumentToolError | undefined;
  try {
    await doc.edit({
      version,
      edits: [{ op: 'insert_paragraphs', after: item(doc, 'Obsolete remark.').ref, paragraphs: [{ text: 'Orphan bullet', style: 'ListBullet' }] }],
    });
  } catch (caught) {
    error = caught as DocumentToolError;
  }
  expect(error?.code).toBe('UNSUPPORTED');
  expect(error?.message).toContain('style "ListBullet" names list numbering that the document\'s numbering part does not define');
  expect(doc.list().version).toBe(version);
  expect(item(doc, 'Orphan bullet')).toBeUndefined();
}, 30_000);
