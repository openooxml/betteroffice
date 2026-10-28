import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const NUMBERING = '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3)"/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="43"><w:abstractNumId w:val="7"/>' +
  '<w:lvlOverride w:ilvl="2"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>';

const EXTRA_STYLES =
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListedStyle"><w:name w:val="Listed Style"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="43"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="NumberedHeading"><w:name w:val="Numbered Heading"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="43"/></w:numPr></w:pPr></w:style>';

const HEADING = '<w:p w:rsidR="00BA03AE" w14:paraId="00004301">' +
  '<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="43"/></w:numPr>' +
  '<w:spacing w:before="0" w:after="160" w:line="259" w:lineRule="auto"/>' +
  '<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>' +
  '<w:r w:rsidRPr="0092171E"><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:b/><w:sz w:val="24"/>' +
  '<w:szCs w:val="24"/></w:rPr><w:t>Some heading</w:t></w:r></w:p>';
const FOLLOWING = '<w:p w14:paraId="00004302"><w:pPr><w:pStyle w:val="ListParagraph"/>' +
  '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="43"/></w:numPr></w:pPr>' +
  '<w:r><w:t>Following item</w:t></w:r></w:p>';
const STYLE_NUMBERED = '<w:p w14:paraId="00004303"><w:pPr><w:pStyle w:val="ListedStyle"/></w:pPr>' +
  '<w:r><w:t>Style numbered heading</w:t></w:r></w:p>';
const PLAIN = '<w:p w14:paraId="00004304"><w:r><w:t>Plain paragraph</w:t></w:r></w:p>';

async function restyleFixture() {
  const zip = await JSZip.loadAsync(await editFixture());
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
  const styles = (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${EXTRA_STYLES}</w:styles>`);
  zip.file('word/styles.xml', styles);
  zip.file('word/numbering.xml', NUMBERING);
  const xml = await zip.file('word/document.xml')!.async('string');
  const document = xml
    .replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ')
    .replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${HEADING}${FOLLOWING}${STYLE_NUMBERED}${PLAIN}`);
  zip.file('word/document.xml', document);
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    document,
    styles,
  };
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
const blocks = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const byText = (xml: string, text: string) => blocks(xml).find(block => block.includes(`>${text}<`))!;

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the edit to be refused.');
}

test('restyling a directly numbered paragraph to Heading 1 keeps its numbering and direct formatting', async () => {
  const source = await restyleFixture();
  const doc = await open(source.bytes);
  const target = item(doc, 'Some heading');
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 1' }],
  });
  expect(doc.read(target.ref).style).toBe('Heading1');
  const pilcrow = doc.session.paragraphs('body').find(entry => entry.text === 'Some heading')!.properties;
  expect(pilcrow.indentLeft).not.toBe(720);
  expect(pilcrow.indentLeft ?? null).toBeNull();
  expect(pilcrow.contextualSpacing ?? null).toBeNull();
  expect(pilcrow.outlineLevel).toBe(0);
  expect(pilcrow.spaceBefore).toBe(0);
  expect(pilcrow.spaceAfter).toBe(160);
  expect(pilcrow.numPr).toMatchObject({ numId: 43, ilvl: 0 });

  const plan = doc.session.paragraphSavePlan().splicedParts!
    .find(part => part.part === 'word/document.xml')!;
  expect(plan.paragraphPatches?.map(patch => patch.ordinal)).toContain(0);
  expect((plan.legacyParagraphs ?? []).filter(entry => entry.ordinal === 0)).toEqual([]);

  const saved = await JSZip.loadAsync(await doc.export());
  const xml = await saved.file('word/document.xml')!.async('string');
  expect(byText(xml, 'Some heading')).toBe(
    HEADING.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
  expect(byText(xml, 'Following item')).toBe(FOLLOWING);
  expect(await saved.file('word/numbering.xml')!.async('string')).toBe(NUMBERING);
  expect(await saved.file('word/styles.xml')!.async('string')).toBe(source.styles);

  const reopened = await open(await doc.export());
  expect(reopened.read(item(reopened, 'Some heading').ref).style).toBe('Heading1');
  expect(item(reopened, 'Some heading').list).toEqual({ level: 0 });
  expect(item(reopened, 'Following item').list).toEqual({ level: 0 });
}, 30_000);

test('restyling a paragraph numbered only by its style to Heading 1 removes its numbering', async () => {
  const source = await restyleFixture();
  const doc = await open(source.bytes);
  expect(item(doc, 'Style numbered heading').list).toEqual({ level: 0 });
  const target = item(doc, 'Style numbered heading');
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 1' }],
  });
  expect(doc.read(target.ref).style).toBe('Heading1');
  const saved = await JSZip.loadAsync(await doc.export());
  const xml = await saved.file('word/document.xml')!.async('string');
  const paragraph = byText(xml, 'Style numbered heading');
  expect(paragraph).toContain('<w:pStyle w:val="Heading1"/>');
  expect(paragraph).not.toContain('<w:numPr>');
  const reopened = await open(await doc.export());
  expect(item(reopened, 'Style numbered heading').list).toBeUndefined();
  expect(reopened.read(item(reopened, 'Style numbered heading').ref).style).toBe('Heading1');
}, 30_000);

test('restyling keeps a direct run size equal to the old style in a paragraph patch', async () => {
  const source = await restyleFixture();
  const zip = await JSZip.loadAsync(source.bytes);
  const styles = source.styles.replace(
    '<w:contextualSpacing/></w:pPr></w:style>',
    '<w:contextualSpacing/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>');
  expect(styles).not.toBe(source.styles);
  zip.file('word/styles.xml', styles);
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const target = item(doc, 'Some heading');
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 1' }],
  });
  const part = doc.session.paragraphSavePlan().splicedParts!
    .find(entry => entry.part === 'word/document.xml')!;
  expect(part.paragraphPatches?.map(patch => patch.ordinal)).toContain(0);
  expect((part.legacyParagraphs ?? []).filter(entry => entry.ordinal === 0)).toEqual([]);
  const saved = await JSZip.loadAsync(await doc.export());
  const xml = await saved.file('word/document.xml')!.async('string');
  const paragraph = byText(xml, 'Some heading');
  expect(paragraph).toContain('<w:sz w:val="24"/>');
  expect(paragraph).toBe(
    HEADING.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
}, 30_000);

test('inherited run formatting stays inherited after saving and restyling back', async () => {
  const source = await restyleFixture();
  const zip = await JSZip.loadAsync(source.bytes);
  zip.file('word/styles.xml', source.styles.replace(
    '<w:contextualSpacing/></w:pPr></w:style>',
    '<w:contextualSpacing/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'));
  const inherited = HEADING.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '');
  zip.file('word/document.xml', source.document.replace(HEADING, inherited));
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: item(doc, 'Some heading').ref, style: 'heading 1' }],
  });
  const part = doc.session.paragraphSavePlan().splicedParts!
    .find(entry => entry.part === 'word/document.xml')!;
  expect(part.paragraphPatches?.map(patch => patch.ordinal)).toContain(0);
  expect((part.legacyParagraphs ?? []).filter(entry => entry.ordinal === 0)).toEqual([]);
  const restyled = doc.session.storySegments('body')
    .find(segment => segment.kind === 'text' && segment.text.includes('Some heading'))!;
  expect(restyled.attributes).toMatchObject({ fontSize: { size: 32 }, bold: true });
  const bytes = await doc.export();
  const saved = await JSZip.loadAsync(bytes);
  const xml = await saved.file('word/document.xml')!.async('string');
  expect(byText(xml, 'Some heading')).toBe(
    inherited.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
  const reopened = await open(bytes);
  await reopened.edit({
    version: reopened.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: item(reopened, 'Some heading').ref, style: 'List Paragraph' }],
  });
  const restored = await JSZip.loadAsync(await reopened.export());
  expect(byText(await restored.file('word/document.xml')!.async('string'), 'Some heading')).toBe(inherited);
  const segments = reopened.session.storySegments('body');
  const heading = segments.find(segment => segment.kind === 'text' && segment.text.includes('Some heading'))!;
  expect(heading.attributes.fontSize).toMatchObject({ size: 24 });
  expect(heading.attributes.bold ?? null).toBeNull();
}, 30_000);

test('inherited highlight stays inherited after saving and restyling back', async () => {
  const source = await restyleFixture();
  const zip = await JSZip.loadAsync(source.bytes);
  const styles = source.styles
    .replace('<w:contextualSpacing/></w:pPr></w:style>',
      '<w:contextualSpacing/></w:pPr><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>')
    .replace('<w:sz w:val="32"/></w:rPr></w:style>',
      '<w:sz w:val="32"/><w:highlight w:val="green"/></w:rPr></w:style>');
  expect(styles).toContain('<w:highlight w:val="yellow"/>');
  expect(styles).toContain('<w:highlight w:val="green"/>');
  zip.file('word/styles.xml', styles);
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: item(doc, 'Some heading').ref, style: 'heading 1' }],
  });
  const part = doc.session.paragraphSavePlan().splicedParts!
    .find(entry => entry.part === 'word/document.xml')!;
  expect(part.paragraphPatches?.map(patch => patch.ordinal)).toContain(0);
  expect((part.legacyParagraphs ?? []).filter(entry => entry.ordinal === 0)).toEqual([]);
  const restyled = doc.session.storySegments('body')
    .find(segment => segment.kind === 'text' && segment.text.includes('Some heading'))!;
  expect(restyled.attributes.highlight).toBe('green');
  const bytes = await doc.export();
  const saved = await JSZip.loadAsync(bytes);
  expect(byText(await saved.file('word/document.xml')!.async('string'), 'Some heading')).toBe(
    HEADING.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
  const reopened = await open(bytes);
  await reopened.edit({
    version: reopened.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: item(reopened, 'Some heading').ref, style: 'List Paragraph' }],
  });
  const restored = await JSZip.loadAsync(await reopened.export());
  expect(byText(await restored.file('word/document.xml')!.async('string'), 'Some heading')).toBe(HEADING);
  const heading = reopened.session.storySegments('body')
    .find(segment => segment.kind === 'text' && segment.text.includes('Some heading'))!;
  expect(heading.attributes.highlight).toBe('yellow');
}, 30_000);

test('restyling an unedited numbered heading with a footnote saves only its style', async () => {
  const source = await restyleFixture();
  const zip = await JSZip.loadAsync(source.bytes);
  const heading = HEADING.replace('<w:t>Some heading</w:t></w:r></w:p>',
    '<w:t>Chapter 1</w:t></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>');
  zip.file('word/document.xml', source.document.replace(HEADING, heading));
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>',
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="footnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
  const footnotes = '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>';
  zip.file('word/footnotes.xml', footnotes);
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const target = doc.list({ limit: 100 }).items.find(entry => entry.text.startsWith('Chapter 1'))!;
  await doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 1' }],
  });
  const part = doc.session.paragraphSavePlan().splicedParts!
    .find(entry => entry.part === 'word/document.xml')!;
  expect(part.paragraphPatches?.map(patch => patch.ordinal)).toContain(0);
  expect((part.legacyParagraphs ?? []).filter(entry => entry.ordinal === 0)).toEqual([]);
  const saved = await JSZip.loadAsync(await doc.export());
  expect(byText(await saved.file('word/document.xml')!.async('string'), 'Chapter 1')).toBe(
    heading.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
  expect(await saved.file('word/footnotes.xml')!.async('string')).toBe(footnotes);
}, 30_000);

test('restyling a text-edited numbered heading requires saving and reopening first', async () => {
  const source = await restyleFixture();
  const doc = await open(source.bytes);
  doc.session.insertText({ story: 'body', paraId: '00004301', offset: 0 }, 'Some ');
  const target = item(doc, 'Some Some heading');
  const version = doc.list().version;
  const before = doc.session.storySegments('body');
  const error = await refusal(doc.edit({
    version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'heading 1' }],
  }));
  expect(error.code).toBe('UNSUPPORTED');
  expect(error.message).toContain('restyle this numbered paragraph before editing its text, or save and reopen first');
  expect(doc.list().version).toBe(version);
  expect(doc.session.storySegments('body')).toEqual(before);
  const bytes = await doc.export();
  const saved = await JSZip.loadAsync(bytes);
  const edited = blocks(await saved.file('word/document.xml')!.async('string'))[0]!;
  const reopened = await open(bytes);
  await reopened.edit({
    version: reopened.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: item(reopened, 'Some Some heading').ref, style: 'heading 1' }],
  });
  const restored = await JSZip.loadAsync(await reopened.export());
  expect(blocks(await restored.file('word/document.xml')!.async('string'))[0]).toBe(
    edited.replace('<w:pStyle w:val="ListParagraph"/>', '<w:pStyle w:val="Heading1"/>'));
}, 30_000);

test('set_paragraph_style still refuses a style that defines list numbering', async () => {
  const source = await restyleFixture();
  const doc = await open(source.bytes);
  const target = item(doc, 'Plain paragraph');
  const error = await refusal(doc.edit({
    version: doc.list().version,
    edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'Numbered Heading' }],
  }));
  expect(error.code).toBe('UNSUPPORTED');
  expect(error.message).toContain('defines list numbering; set_paragraph_style cannot apply a numbered style yet');
  expect(doc.read(target.ref).style).not.toBe('NumberedHeading');
}, 30_000);
