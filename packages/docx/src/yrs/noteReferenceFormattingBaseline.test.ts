import { describe, expect, it } from 'bun:test';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { documentToYrs } from './documentToYrs';
import { DirtyProjectionStories, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, type YrsSession } from './index';
import { documentXml, setupStructuralSaves } from './structuralSaveTestSupport';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const DATE = '2025-03-04T05:06:07Z';

type Apply = (session: YrsSession) => void;

const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const INSERTED_MARK = `<w:ins w:id="21" w:author="Ada" w:date="${DATE}">${text('!')}</w:ins>`;
const STYLED_REFERENCE = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
const BARE_REFERENCE = '<w:r><w:footnoteReference w:id="1"/></w:r>';
const ownRun = (properties: string) => `<w:r><w:rPr>${properties}</w:rPr><w:footnoteReference w:id="1"/></w:r>`;
const paragraph = (content: string) => `<w:p w14:paraId="0E000001" w14:textId="0E000001">${content}</w:p>`;
const around = (reference: string, after = INSERTED_MARK) => paragraph(`${text('Head ')}${reference}${text(' tail')}${after}`);
const FIELD_PARAGRAPH = paragraph(
  `${text('Head ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "target" </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:hyperlink w:anchor="target">${STYLED_REFERENCE}${text('Link')}</w:hyperlink>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>${text(' tail')}${INSERTED_MARK}`
);
const DELETED_REFERENCE_PARAGRAPH = around(`<w:del w:id="12" w:author="Bo" w:date="${DATE}">${STYLED_REFERENCE}</w:del>`, '');

const DEFAULT_STYLE = '<w:docDefaults><w:rPrDefault><w:rPr><w:rStyle w:val="Quiet"/></w:rPr></w:rPrDefault></w:docDefaults>';
const quietStyle = (quiet: string) =>
  `<w:style w:type="character" w:styleId="Quiet"><w:name w:val="Quiet"/><w:rPr>${quiet}</w:rPr></w:style>`;
const DEFAULT_STYLE_ROWS = [
  { sets: 'a colour', quiet: '<w:color w:val="333333"/>', reset: '<w:color w:val="auto"/>' },
  { sets: 'an underline', quiet: '<w:u w:val="single"/>', reset: '<w:u w:val="none"/>' },
  { sets: 'bold', quiet: '<w:b/>', reset: '<w:b w:val="0"/>' },
];

const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>' +
  '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';

const FOOTNOTES = `<w:footnotes xmlns:w="${W}">` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
  `${text(' footnote body')}</w:p></w:footnote></w:footnotes>`;

function fixture(body: string, quiet?: string): Uint8Array {
  const styles = quiet === undefined ? STYLES : `${DEFAULT_STYLE}${STYLES}${quietStyle(quiet)}`;
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/></Types>`
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdFootnotes" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`
    )],
    ['word/styles.xml', toBytes(`<w:styles xmlns:w="${W}">${styles}</w:styles>`)],
    ['word/footnotes.xml', toBytes(FOOTNOTES)],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`
    )],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function firstParagraph(xml: string): string {
  return /<w:p\b[^>]*>[\s\S]*?<\/w:p>/.exec(xml)?.[0] ?? '';
}

function enclosingRun(xml: string, element: string): string {
  return new RegExp(`<w:r>(?:(?!<w:r>)[\\s\\S])*?${element}</w:r>`).exec(xml)?.[0] ?? '';
}

const plainText = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');
const fieldCharacterRun = (xml: string, type: string) => enclosingRun(xml, `<w:fldChar w:fldCharType="${type}"/>`);
const referenceRun = (xml: string) => enclosingRun(xml, '<w:footnoteReference w:id="1"/>');

function paraIdOf(session: YrsSession): string {
  const target = session.paragraphs('body')[0];
  if (!target) throw new Error('no body paragraph');
  return target.paraId;
}

function typeAt(session: YrsSession, offset: number): void {
  session.insertText({ story: 'body', paraId: paraIdOf(session), offset }, 'X');
}

const typeIntoHead: Apply = (session) => typeAt(session, 2);

function typeIntoLink(session: YrsSession): void {
  let offset = 0;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'pilcrow') break;
    if (segment.kind === 'text' && segment.text.includes('Link')) {
      typeAt(session, offset + segment.text.indexOf('Link') + 2);
      return;
    }
    offset += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('no Link text in the first paragraph');
}

function boldHead(session: YrsSession): void {
  const paraId = paraIdOf(session);
  session.toggleMark({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }, { type: 'bold' });
}

const structuralSave = setupStructuralSaves();

async function saveFromParsedDocument(source: Uint8Array, apply: Apply): Promise<string> {
  const parsed = await parseDocx(source.slice().buffer, { preloadFonts: false });
  const session = await createYrsSession();
  try {
    documentToYrs(session, parsed);
    apply(session);
    const dirty = new DirtyProjectionStories();
    dirty.add('body');
    const live = session.materializeDocx();
    const base = live ? mergeDocxHostMetadata(live, parsed) : parsed;
    const projected = yrsToDocument(session, base, dirty.projectionOptions());
    return documentXml(new Uint8Array(await saveEditorDocument(session, projected, [])));
  } finally {
    session.destroy();
  }
}

interface EditorPath {
  name: string;
  keepsSourceBytes: boolean;
  open(source: Uint8Array, apply: Apply): Promise<string>;
}

const PATHS: EditorPath[] = [
  {
    name: 'Rust loader session',
    keepsSourceBytes: true,
    open: (source, apply) =>
      structuralSave('main-thread (hybrid and in-thread)', source, (engine) => apply(engine as unknown as YrsSession)),
  },
  {
    name: 'worker engine',
    keepsSourceBytes: true,
    open: (source, apply) => structuralSave('worker-only', source, (engine) => apply(engine as unknown as YrsSession)),
  },
  { name: 'TypeScript loader', keepsSourceBytes: false, open: saveFromParsedDocument },
];

describe.each(PATHS)('$name', (path) => {
  it('writes the field characters and the footnote reference of a rebuilt field result bare', async () => {
    const saved = firstParagraph(await path.open(fixture(FIELD_PARAGRAPH), typeIntoLink));
    expect(plainText(saved)).toBe('Head LiXnk tail!');
    expect({
      begin: fieldCharacterRun(saved, 'begin'),
      separate: fieldCharacterRun(saved, 'separate'),
      end: fieldCharacterRun(saved, 'end'),
      reference: referenceRun(saved),
    }).toEqual({
      begin: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
      separate: '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
      end: '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
      reference: BARE_REFERENCE,
    });
  });

  for (const row of DEFAULT_STYLE_ROWS) {
    it(`writes a default character style that sets ${row.sets} with its reset on a rewritten text run and none on a bare footnote reference`, async () => {
      const saved = firstParagraph(await path.open(fixture(around(BARE_REFERENCE), row.quiet), typeIntoHead));
      expect({
        text: enclosingRun(saved, '<w:t xml:space="preserve">HeXad </w:t>'),
        reference: referenceRun(saved),
      }).toEqual({
        text: `<w:r><w:rPr><w:rStyle w:val="Quiet"/>${row.reset}</w:rPr><w:t xml:space="preserve">HeXad </w:t></w:r>`,
        reference: BARE_REFERENCE,
      });
    });
  }

  it('writes an untracked styled footnote reference bare when a tracked change in its paragraph makes a rewrite', async () => {
    const saved = firstParagraph(await path.open(fixture(around(STYLED_REFERENCE)), typeIntoHead));
    expect(plainText(saved)).toBe('HeXad  tail!');
    expect(referenceRun(saved)).toBe(BARE_REFERENCE);
  });

  it('writes an untracked styled footnote reference on a plain save as its source run, or bare where the paragraph is rewritten', async () => {
    const saved = firstParagraph(await path.open(fixture(around(STYLED_REFERENCE, '')), () => {}));
    expect(referenceRun(saved)).toBe(path.keepsSourceBytes ? STYLED_REFERENCE : BARE_REFERENCE);
  });

  it('writes a footnote reference whose own run has a language bare when its paragraph is rewritten', async () => {
    const saved = firstParagraph(await path.open(fixture(around(ownRun('<w:lang w:val="fr-FR"/>'))), typeIntoHead));
    expect(plainText(saved)).toBe('HeXad  tail!');
    expect(referenceRun(saved)).toBe(BARE_REFERENCE);
  });

  it('writes a footnote reference whose own run has a font hint, bCs, iCs and a language bare when its paragraph is rewritten', async () => {
    const own = ownRun('<w:rFonts w:hint="eastAsia"/><w:bCs/><w:iCs/><w:lang w:val="fr-FR"/>');
    const saved = firstParagraph(await path.open(fixture(around(own)), typeIntoHead));
    expect(plainText(saved)).toBe('HeXad  tail!');
    expect(referenceRun(saved)).toBe(BARE_REFERENCE);
  });
});

describe.each(PATHS.filter((path) => path.keepsSourceBytes))('$name', (path) => {
  it('keeps the source paragraph when text beside a deleted footnote reference is bolded', async () => {
    const saved = firstParagraph(await path.open(fixture(DELETED_REFERENCE_PARAGRAPH), boldHead));
    expect(saved).toBe(DELETED_REFERENCE_PARAGRAPH);
  });
});
