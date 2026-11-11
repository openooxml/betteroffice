import JSZip from 'jszip';
import type * as Yrs from '../index';
import type { DocxEditStep, DocxTextMarks } from '../edits';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** `00000001` plain, `00000002` directly bold, `00000003` a bold Heading1, `00000004` a run in the bold character style `Strong`, `00000005` in the superscript style `Raised`, `00000006` in the highlighting style `Marked`. */
export async function textRunFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`, { date });
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', { date });
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', { date });
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Raised"><w:name w:val="Raised"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Marked"><w:name w:val="Marked"/><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style></w:styles>`, { date });
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Plain text</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words</w:t></w:r></w:p><w:p w14:paraId="00000003"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p w14:paraId="00000004"><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>Strong run</w:t></w:r></w:p><w:p w14:paraId="00000005"><w:pPr><w:pStyle w:val="Raised"/></w:pPr><w:r><w:t>Up</w:t></w:r></w:p><w:p w14:paraId="00000006"><w:pPr><w:pStyle w:val="Marked"/></w:pPr><w:r><w:t>Lit</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const paragraph = (paraId: string) => ({ kind: 'paragraph', story: 'body', paraId }) as const;

/** Steps that differ only in plain `text` against one unmarked run; both must save identically. */
export const PLAIN_EQUIVALENT_STEPS: ReadonlyArray<{ name: string; text: DocxEditStep; runs: DocxEditStep }> = [
  ['plain', '00000001'], ['bold', '00000002'], ['heading', '00000003'], ['character style', '00000004'],
].flatMap(([context, paraId]) => [
  {
    name: `insert into ${context} text`,
    text: { op: 'insertText', target: paragraph(paraId), at: 'end', text: ' x\ty' },
    runs: { op: 'insertText', target: paragraph(paraId), at: 'end', runs: [{ text: ' x\ty' }] },
  },
  {
    name: `new paragraph after ${context} text`,
    text: { op: 'insertParagraphs', target: { story: 'body', paraId }, at: 'end', paragraphs: [{ text: 'New' }] },
    runs: { op: 'insertParagraphs', target: { story: 'body', paraId }, at: 'end', paragraphs: [{ runs: [{ text: 'New' }] }] },
  },
]);

/** One marked run, the run properties its saved `w:rPr` must hold, and those that must be absent or explicitly off. */
export const MARKED_RUN_CASES: ReadonlyArray<{
  marks: DocxTextMarks;
  saved: ReadonlyArray<[string, Record<string, string>]>;
  off?: readonly string[];
}> = [
  { marks: { bold: true }, saved: [['w:b', {}]] },
  { marks: { italic: true }, saved: [['w:i', {}]] },
  { marks: { underline: true }, saved: [['w:u', { 'w:val': 'single' }]] },
  { marks: { underline: 'double' }, saved: [['w:u', { 'w:val': 'double' }]] },
  { marks: { strike: true }, saved: [['w:strike', {}]] },
  { marks: { strike: 'dblStrike' }, saved: [['w:dstrike', {}]], off: ['w:strike'] },
  { marks: { verticalAlign: 'superscript' }, saved: [['w:vertAlign', { 'w:val': 'superscript' }]] },
  { marks: { verticalAlign: 'subscript' }, saved: [['w:vertAlign', { 'w:val': 'subscript' }]] },
  { marks: { color: 'C00000' }, saved: [['w:color', { 'w:val': 'C00000' }]] },
  { marks: { color: { themeColor: 'accent1' } }, saved: [['w:color', { 'w:themeColor': 'accent1' }]] },
  { marks: { highlight: 'yellow' }, saved: [['w:highlight', { 'w:val': 'yellow' }]] },
  { marks: { fontSizePt: 10.5 }, saved: [['w:sz', { 'w:val': '21' }], ['w:szCs', { 'w:val': '21' }]] },
  { marks: { fontFamily: 'Arial' }, saved: [['w:rFonts', { 'w:ascii': 'Arial', 'w:hAnsi': 'Arial' }]] },
];

/** The snippet from the request: three runs, one undo step. */
export const SNIPPET_STEP: DocxEditStep = {
  op: 'insertText',
  target: paragraph('00000001'),
  at: 'end',
  runs: [
    { text: ' The item was checked and approved. ' },
    { text: 'No changes needed.', marks: { bold: true } },
    { text: ' ' },
    { text: 'See the attached notes for details.', marks: { italic: true } },
  ],
};

/** The `w:rPr` of the saved run whose text holds `text`, or `null` for a run without one. */
export function savedRunProperties(documentXml: string, text: string): string | null {
  const escaped = text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = new RegExp(`<w:r>(?:<w:rPr>((?:(?!</w:rPr>).)*)</w:rPr>)?<w:t(?: [^>]*)?>[^<]*${escaped}[^<]*</w:t></w:r>`).exec(documentXml);
  if (!match) throw new Error(`no saved run reads ${JSON.stringify(text)}`);
  return match[1] ?? null;
}

/** Whether `properties` holds a `name` element carrying every attribute in `attributes`. */
export function holdsProperty(properties: string | null, name: string, attributes: Record<string, string>): boolean {
  return [...(properties ?? '').matchAll(new RegExp(`<${name}(\\s[^>]*)?/>`, 'g'))].some(([, found = '']) =>
    Object.entries(attributes).every(([key, value]) => found.includes(`${key}="${value}"`))
  );
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

/** Applies `steps` to a fresh session, saves, and reopens the saved bytes. */
export async function editSaveReopen(api: typeof Yrs, steps: readonly DocxEditStep[], clientId = 41) {
  const source = await textRunFixture();
  const session = await api.createYrsSession({ clientId });
  const reopened = await api.createYrsSession({ clientId: clientId + 1 });
  try {
    session.seedFromDocx(source);
    const result = session.applyEdits({ expectVersion: session.version(), steps });
    if (!result.ok) throw new Error(JSON.stringify(result.failure));
    const saved = await api.saveYrsDocx(session, { updateModifiedDate: false });
    reopened.seedFromDocx(saved.bytes);
    const resaved = await api.saveYrsDocx(reopened, { updateModifiedDate: false });
    return {
      result,
      bytes: saved.bytes,
      xml: await documentXml(saved.bytes),
      resavedXml: await documentXml(resaved.bytes),
      segments: session.storySegments('body'),
      reopened: reopened.storySegments('body'),
    };
  } finally {
    session.destroy();
    reopened.destroy();
  }
}

/** Whether `properties` turns `name` on: the element is present without an off `w:val`. */
export function turnsOn(properties: string | null, name: string): boolean {
  return [...(properties ?? '').matchAll(new RegExp(`<${name}(\\s[^>]*)?/>`, 'g'))].some(([, found = '']) =>
    !/w:val="(0|false|off)"/.test(found)
  );
}
