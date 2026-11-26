import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { editFixture } from './fixture';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const numbering = `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const HEADING = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';

let id = 0x3000;
const p = (pPr: string, text: string, rPr = '') => {
  id++;
  const hex = id.toString(16).toUpperCase().padStart(8, '0');
  return `<w:p w14:paraId="${hex}" w14:textId="7${hex.slice(1)}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
};

async function fixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await editFixture());
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
  zip.file('word/numbering.xml', numbering);
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${HEADING}</w:styles>`));
  const body = p('<w:pStyle w:val="Heading1"/>', 'Heading')
    + p('', 'Accent anchor', '<w:i/><w:color w:val="C00000"/>')
    + p('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/>', 'Bullet item')
    + p('', 'Tail');
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml
    .replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ')
    .replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${body}`));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const opened: DocxAgentDocument[] = [];
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await fixture());
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

type Edit = Extract<DocxAgentEdit, { op: 'insert_paragraphs' }>;
const item = (doc: DocxAgentDocument, text: string) => doc.list({ limit: 100 }).items.find(entry => entry.text === text)!;
async function insert(doc: DocxAgentDocument, anchor: string, edit: Omit<Edit, 'op'>) {
  const where = edit.before !== undefined ? { before: item(doc, anchor).ref } : { after: item(doc, anchor).ref };
  return doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', ...edit, ...where }] });
}
async function code(promise: Promise<unknown>) {
  try { await promise; return 'ok'; } catch (error) { return (error as { code: string }).code; }
}
async function documentXml(bytes: Uint8Array) {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}
const blocks = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const textOf = (paragraph: string) => [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
const on = (rPr: string, tag: string) => new RegExp(`<w:${tag}(?:\\s+w:val="(?:true|1|on)")?\\s*/>`).test(rPr);
/** Characters of one saved paragraph grouped by their bold, italic and underline. */
function marks(paragraph: string) {
  const spans: Array<{ text: string; bold: boolean; italic: boolean; underline: boolean }> = [];
  for (const run of paragraph.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []) {
    const rPr = run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
    const span = { text: textOf(run), bold: on(rPr, 'b'), italic: on(rPr, 'i'), underline: /<w:u w:val="single"\/>/.test(rPr) };
    const last = spans.at(-1);
    if (last && last.bold === span.bold && last.italic === span.italic && last.underline === span.underline) last.text += span.text;
    else if (span.text) spans.push(span);
  }
  return spans;
}
const runsOf = (paragraph: string) => (paragraph.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).map(run => run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '');
/** Saves, reopens and saves again, returning the reopened document.xml. */
async function roundTrip(doc: DocxAgentDocument) {
  const reopened = await open(await doc.export());
  return documentXml(await reopened.export());
}

const LABELLED = [
  { text: 'Label: ', bold: true }, { text: 'plain, ' }, { text: 'slanted', italic: true },
  { text: ', ' }, { text: 'both', bold: true, italic: true }, { text: ' and ' }, { text: 'lined', underline: true },
];
const EXPECTED = [
  { text: 'Label: ', bold: true, italic: false, underline: false },
  { text: 'plain, ', bold: false, italic: false, underline: false },
  { text: 'slanted', bold: false, italic: true, underline: false },
  { text: ', ', bold: false, italic: false, underline: false },
  { text: 'both', bold: true, italic: true, underline: false },
  { text: ' and ', bold: false, italic: false, underline: false },
  { text: 'lined', bold: false, italic: false, underline: true },
];
const JOINED = 'Label: plain, slanted, both and lined';

test('runs put bold, italic, both and underline on the right characters after save and reopen', async () => {
  for (const placement of [{ after: 'x' }, { before: 'x' }] as const) {
    const doc = await open();
    const receipt = await insert(doc, 'Tail', { ...placement, paragraphs: [{ runs: LABELLED }, { runs: [{ text: 'Second ' }, { text: 'bold', bold: true }] }] });
    const refs = receipt.results[0].refs as string[];
    expect(refs.map(ref => doc.read(ref).text)).toEqual([JOINED, 'Second bold']);
    const xml = await roundTrip(doc);
    expect(marks(blocks(xml).find(paragraph => textOf(paragraph) === JOINED)!)).toEqual(EXPECTED);
    expect(marks(blocks(xml).find(paragraph => textOf(paragraph) === 'Second bold')!)).toEqual([
      { text: 'Second ', bold: false, italic: false, underline: false },
      { text: 'bold', bold: true, italic: false, underline: false },
    ]);
  }
});

test('one unmarked run saves exactly like text in every placement', async () => {
  const cases: Array<[string, Omit<Edit, 'op' | 'paragraphs'>]> = [
    ['Tail', { after: 'x' }],
    ['Heading', { after: 'x' }],
    ['Accent anchor', { before: 'x' }],
    ['Accent anchor', { after: 'x', runFormat: 'anchor' }],
    ['Accent anchor', { before: 'x', runFormat: 'anchor' }],
    ['Bullet item', { after: 'x', list: 'continue' }],
  ];
  const bytes = await fixture();
  const withoutIds = (xml: string) => xml.replace(/ w14:(?:paraId|textId)="[^"]*"/g, '');
  for (const [anchor, edit] of cases) {
    const withText = await open(bytes);
    const withRuns = await open(bytes);
    await insert(withText, anchor, { ...edit, paragraphs: [{ text: 'Same words' }, { text: '' }] });
    await insert(withRuns, anchor, { ...edit, paragraphs: [{ runs: [{ text: 'Same words' }] }, { runs: [{ text: '' }] }] });
    expect(withoutIds(await documentXml(await withRuns.export()))).toBe(withoutIds(await documentXml(await withText.export())));
  }
});

test('text with runs, neither, empty runs and an empty run beside others are refused', async () => {
  const doc = await open();
  const refused = (paragraphs: unknown[]) => code(insert(doc, 'Tail', { after: 'x', paragraphs: paragraphs as Edit['paragraphs'] }));
  expect(await refused([{ text: 'Both', runs: [{ text: 'Both' }] }])).toBe('INVALID_ARGUMENT');
  expect(await refused([{ style: 'Heading1' }])).toBe('INVALID_ARGUMENT');
  expect(await refused([{ runs: [] }])).toBe('INVALID_ARGUMENT');
  expect(await refused([{ runs: [{ text: 'A', bold: true }, { text: '' }] }])).toBe('INVALID_ARGUMENT');
  expect(await refused([{ runs: [{ text: 'Tab\there' }] }])).toBe('INVALID_ARGUMENT');
  expect(await refused([{ runs: [{ text: 'A', strike: true }] }])).toBe('INVALID_ARGUMENT');
  expect(doc.list({ limit: 100 }).items.map(entry => entry.text)).toEqual(['Heading', 'Accent anchor', 'Bullet item', 'Tail']);
});

test('runFormat anchor copies the anchor run under every run and adds the run marks on top', async () => {
  for (const placement of [{ after: 'x' }, { before: 'x' }] as const) {
    const doc = await open();
    await insert(doc, 'Accent anchor', { ...placement, runFormat: 'anchor', paragraphs: [{ runs: [{ text: 'Strong', bold: true }, { text: ' rest' }, { text: ' upright', italic: false }] }] });
    const saved = blocks(await roundTrip(doc)).find(paragraph => textOf(paragraph) === 'Strong rest upright')!;
    expect(marks(saved)).toEqual([
      { text: 'Strong', bold: true, italic: true, underline: false },
      { text: ' rest', bold: false, italic: true, underline: false },
      { text: ' upright', bold: false, italic: false, underline: false },
    ]);
    expect(runsOf(saved).every(rPr => rPr.includes('<w:color w:val="C00000"/>'))).toBe(true);
  }
});

test('list continue joins the list and keeps the run marks', async () => {
  const doc = await open();
  await insert(doc, 'Bullet item', { after: 'x', list: 'continue', paragraphs: [{ runs: [{ text: 'Term: ', bold: true }, { text: 'definition' }] }] });
  expect(item(doc, 'Term: definition').list).toEqual(item(doc, 'Bullet item').list);
  const saved = blocks(await roundTrip(doc)).find(paragraph => textOf(paragraph) === 'Term: definition')!;
  expect(saved).toContain('<w:numId w:val="1"/>');
  expect(marks(saved)).toEqual([
    { text: 'Term: ', bold: true, italic: false, underline: false },
    { text: 'definition', bold: false, italic: false, underline: false },
  ]);
});

test('run marks sit on top of the paragraph style', async () => {
  const doc = await open();
  await insert(doc, 'Tail', { after: 'x', paragraphs: [{ style: 'Heading1', runs: [{ text: 'Styled ' }, { text: 'slanted', italic: true }] }] });
  const saved = blocks(await roundTrip(doc)).find(paragraph => textOf(paragraph) === 'Styled slanted')!;
  expect(saved).toContain('<w:pStyle w:val="Heading1"/>');
  expect(marks(saved).map(({ text, italic }) => ({ text, italic }))).toEqual([{ text: 'Styled ', italic: false }, { text: 'slanted', italic: true }]);
});
