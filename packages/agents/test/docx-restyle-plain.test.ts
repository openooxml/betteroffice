import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const STYLE_PROPERTIES = '<w:pPr><w:keepNext/>' +
  '<w:pBdr><w:bottom w:val="single" w:sz="8" w:color="112233"/></w:pBdr>' +
  '<w:spacing w:before="240" w:after="80" w:line="276" w:lineRule="auto"/>' +
  '<w:ind w:left="720"/><w:outlineLvl w:val="0"/></w:pPr>';
const DIRECT_SPACING = '<w:spacing w:before="10" w:after="160" w:line="259" w:lineRule="auto"/>';
const opened: DocxAgentDocument[] = [];

async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

async function restyleFixture(properties: string, runProperties: string) {
  const zip = await JSZip.loadAsync(await editFixture());
  const styles = await zip.file('word/styles.xml')!.async('string');
  zip.file('word/styles.xml', styles.replace('<w:pPr><w:outlineLvl w:val="0"/></w:pPr>', STYLE_PROPERTIES));
  const run = `<w:r>${runProperties}<w:t>Other</w:t></w:r>`;
  const paragraph = '<w:p w14:paraId="00004304">' + properties + run + '</w:p>';
  const document = (await zip.file('word/document.xml')!.async('string'))
    .replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ')
    .replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${paragraph}`);
  zip.file('word/document.xml', document);
  return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), document, paragraph, run };
}

function properties(doc: DocxAgentDocument) {
  return doc.session.paragraphs('body')[0]!.properties;
}

function expectInherited(doc: DocxAgentDocument, color: string, before: number, left: number, fontSize = 32) {
  const attrs = properties(doc);
  expect(attrs.borders).toMatchObject({ bottom: { color: { rgb: color } } });
  expect(attrs.spaceBefore).toBe(before);
  expect(attrs.spaceAfter).toBe(80);
  expect(attrs.lineSpacing).toBe(276);
  expect(attrs.indentLeft).toBe(left);
  expect(attrs.defaultTextFormatting).toMatchObject({ bold: true, fontSize });
  const direct = attrs._originalFormatting as Record<string, unknown>;
  expect(direct.styleId).toBe('Heading1');
  for (const key of ['borders', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'indentLeft', 'keepNext', 'runProperties']) {
    expect(direct[key]).toBeUndefined();
  }
}

for (const [name, sourceProperties, runProperties] of [
  ['without pPr', '', ''],
  ['with a style', '<w:pPr><w:pStyle w:val="Normal"/></w:pPr>', ''],
  ['with direct spacing', `<w:pPr><w:pStyle w:val="Normal"/>${DIRECT_SPACING}</w:pPr>`, ''],
  ['with direct run properties', '', '<w:rPr><w:lang w:val="de-DE"/></w:rPr>'],
] as const) {
  test(`restyling a plain paragraph ${name} keeps new style properties inherited`, async () => {
    const source = await restyleFixture(sourceProperties, runProperties);
    const doc = await open(source.bytes);
    await doc.edit({
      version: doc.list().version,
      edits: [{ op: 'set_paragraph_style', paragraph: doc.list().items[0]!.ref, style: 'heading 1' }],
    });
    expectInherited(doc, '112233', 240, 720);
    expect(properties(doc).keepNext).toBe(true);
    const part = doc.session.paragraphSavePlan().splicedParts!
      .find(entry => entry.part === 'word/document.xml')!;
    if (!sourceProperties.includes(DIRECT_SPACING) && !runProperties) {
      expect(part.paragraphPatches?.map(patch => patch.ordinal)).toEqual([0]);
      expect(part.legacyParagraphs ?? []).toEqual([]);
    } else {
      expect(part.paragraphPatches ?? []).toEqual([]);
      expect(part.legacyParagraphs?.map(paragraph => paragraph.ordinal)).toEqual([0]);
    }
    const exported = await doc.export();
    const saved = await JSZip.loadAsync(exported);
    const expected = '<w:p w14:paraId="00004304"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
      source.run + '</w:p>';
    expect(await saved.file('word/document.xml')!.async('string')).toBe(
      source.document.replace(source.paragraph, expected));
    const reopened = await open(exported);
    expectInherited(reopened, '112233', 240, 720);
    expect(properties(reopened).keepNext).toBe(true);
    const styles = await saved.file('word/styles.xml')!.async('string');
    saved.file('word/styles.xml', styles.replace('112233', '445566')
      .replace('w:before="240"', 'w:before="480"')
      .replace('w:left="720"', 'w:left="1440"').replace('<w:keepNext/>', '')
      .replace('<w:sz w:val="32"/>', '<w:sz w:val="40"/>'));
    const updated = await open(await saved.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    expectInherited(updated, '445566', 480, 1440, 40);
    expect(properties(updated).keepNext ?? null).toBeNull();
  }, 30_000);
}
