import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { editFixture } from './fixture';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const level = (ilvl: number, format: string, text: string, left: number) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
const numbering = `<w:numbering xmlns:w="${w}">` +
  `<w:abstractNum w:abstractNumId="0">${level(0, 'bullet', '•', 720)}${level(1, 'bullet', 'o', 1440)}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="1">${level(0, 'decimal', '%1.', 720)}${level(1, 'lowerLetter', '%2)', 1440)}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="2">${level(0, 'decimal', '%1.', 720).replace('<w:lvlText', '<w:pStyle w:val="ListNumber"/><w:lvlText')}</w:abstractNum>` +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>';
const styles = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr><w:spacing w:after="60"/></w:pPr></w:style>';

let id = 0x1000;
const p = (pPr: string, ...runs: Array<[string, string]>) => {
  id++;
  return `<w:p w14:paraId="${id.toString(16).toUpperCase().padStart(8, '0')}" w14:textId="7${id.toString(16).toUpperCase().padStart(7, '0')}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}` +
    runs.map(([rPr, text]) => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`).join('') + '</w:p>';
};
const bullet = (ilvl: number) => `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="40"/><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/>`;
const step = '<w:keepNext/><w:pageBreakBefore/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:tabs><w:tab w:val="left" w:pos="1080"/></w:tabs><w:spacing w:before="20" w:after="20"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/><w:rPr><w:sz w:val="20"/></w:rPr>';

async function listFixture(anchor?: string, extraStyles = ''): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await editFixture());
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${styles}${extraStyles}</w:styles>`));
  zip.file('word/numbering.xml', numbering);
  const xml = await zip.file('word/document.xml')!.async('string');
  const body = [
    p('<w:pStyle w:val="Heading1"/>', ['', 'Checklist']),
    p(bullet(0), ['<w:b/>', 'Alpha: '], ['<w:color w:val="1F4E79"/>', 'first bullet']),
    p(bullet(1), ['<w:i/>', 'Alpha detail']),
    anchor ?? p(bullet(0), ['', 'Beta bullet']),
    p(bullet(0), ['<w:color w:val="1F4E79"/>', 'Gamma bullet']),
    p('<w:jc w:val="center"/>', ['', 'Between lists.']),
    p(step, ['<w:sz w:val="20"/>', 'First step']),
    p(step, ['<w:sz w:val="20"/>', 'Second step']),
    p('<w:pStyle w:val="ListNumber"/>', ['', 'Style one']),
    p('<w:pStyle w:val="ListNumber"/>', ['', 'Style two']),
    `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(bullet(0), ['', 'Cell bullet'])}</w:tc></w:tr></w:tbl>`,
    p('', ['', 'Closing note.']),
  ].join('');
  zip.file('word/document.xml', xml
    .replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ')
    .replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${body}`));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const opened: DocxAgentDocument[] = [];
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await listFixture());
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

async function parts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  return { document: await zip.file('word/document.xml')!.async('string'), numbering: await zip.file('word/numbering.xml')!.async('string') };
}
const blocks = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const textOf = (paragraph: string) => [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
const pPrOf = (paragraph: string) => paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
const rPrsOf = (paragraph: string) => (paragraph.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).map(run => run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '');
const item = (doc: DocxAgentDocument, text: string) => doc.list({ limit: 100 }).items.find(entry => entry.text === text)!;
const markers = (doc: DocxAgentDocument) => (doc.session.yrsBlocksForStory('body') as Array<{ kind: string; attrs?: { listMarker?: string } }>)
  .flatMap(block => block.kind === 'paragraph' && block.attrs?.listMarker ? [block.attrs.listMarker] : []);

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the batch to be refused.');
}

/** Continues the list after `anchorText`; every other paragraph and numbering.xml must save byte-identical. */
async function continued(doc: DocxAgentDocument, anchorText: string, texts: string[]) {
  const before = await parts(await doc.export());
  const anchor = item(doc, anchorText);
  const receipt = await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: anchor.ref, list: 'continue', paragraphs: texts.map(text => ({ text })) }] });
  const bytes = await doc.export();
  const after = await parts(bytes);
  const old = blocks(before.document);
  const now = blocks(after.document);
  const at = old.findIndex(paragraph => textOf(paragraph) === anchorText);
  const added = now.slice(at + 1, at + 1 + texts.length);
  expect(added.map(textOf)).toEqual(texts);
  expect([...now.slice(0, at + 1), ...now.slice(at + 1 + texts.length)]).toEqual(old);
  expect(after.numbering).toBe(before.numbering);
  for (const name of ['paraId', 'textId']) {
    const ids = [...after.document.matchAll(new RegExp(`w14:${name}="([^"]+)"`, 'g'))].map(match => match[1]);
    expect(new Set(ids).size).toBe(ids.length);
  }
  const refs = receipt.results[0].refs as string[];
  expect(refs.map(ref => doc.read(ref).text)).toEqual(texts);
  for (const ref of refs) expect(doc.list({ limit: 100 }).items.find(entry => entry.ref === ref)?.list).toEqual(anchor.list);
  const reopened = await open(bytes);
  for (const text of texts) expect(item(reopened, text).list).toEqual(anchor.list);
  return { anchor: old[at], added, anchorItem: anchor };
}

test('outline marks list paragraphs with their zero-based level', async () => {
  const doc = await open();
  const items = doc.list({ limit: 100 }).items;
  expect(items.filter(entry => entry.list).map(entry => [entry.text, entry.list!.level])).toEqual([
    ['Alpha: first bullet', 0], ['Alpha detail', 1], ['Beta bullet', 0], ['Gamma bullet', 0],
    ['First step', 0], ['Second step', 0], ['Style one', 0], ['Style two', 0], ['Cell bullet', 0],
  ]);
  expect(items.filter(entry => !entry.list)).toHaveLength(3);
  expect(Object.keys(item(doc, 'Checklist'))).not.toContain('list');
});

test('list continue adds a bullet after a middle and after the last item with the anchor pPr', async () => {
  const doc = await open();
  for (const [anchor, text] of [['Beta bullet', 'Beta and a half'], ['Gamma bullet', 'Delta bullet']]) {
    const result = await continued(doc, anchor, [text]);
    expect(result.anchorItem.list).toEqual({ level: 0 });
    expect(pPrOf(result.added[0])).toBe(pPrOf(result.anchor));
    expect(pPrOf(result.added[0])).toContain('<w:numId w:val="1"/>');
    expect(rPrsOf(result.added[0])).toEqual(rPrsOf(result.anchor).slice(-1));
  }
  const order = doc.list({ limit: 100 }).items.map(entry => entry.text);
  expect(markers(doc)).toEqual(['•', '○', '•', '•', '•', '•', '1.', '2.', '1.', '2.']);
  expect(order.slice(1, 7)).toEqual(['Alpha: first bullet', 'Alpha detail', 'Beta bullet', 'Beta and a half', 'Gamma bullet', 'Delta bullet']);
});

test('list continue keeps a nested item at its level and takes the last run formatting', async () => {
  const doc = await open();
  const { anchor, added } = await continued(doc, 'Alpha detail', ['Alpha second detail']);
  expect(pPrOf(added[0])).toBe(pPrOf(anchor));
  expect(pPrOf(added[0])).toContain('<w:ilvl w:val="1"/>');
  expect(item(doc, 'Alpha second detail').list).toEqual({ level: 1 });
  const first = await continued(doc, 'Alpha: first bullet', ['Alpha follow-up']);
  expect(rPrsOf(first.added[0])).toEqual([rPrsOf(first.anchor)[1]]);
  expect(rPrsOf(first.added[0])[0]).not.toContain('<w:b/>');
});

test('list continue follows a list paragraph style without writing direct numbering', async () => {
  const doc = await open();
  const { anchor, added } = await continued(doc, 'Style one', ['Style one and a half']);
  expect(pPrOf(added[0])).toBe(pPrOf(anchor));
  expect(added[0]).toContain('<w:pStyle w:val="ListNumber"/>');
  expect(added[0]).not.toContain('<w:numPr>');
  expect(item(doc, 'Style one and a half').list).toEqual({ level: 0 });
  expect(markers(doc).slice(-3)).toEqual(['1.', '2.', '3.']);
});

test('list continue takes bold paragraph mark formatting from an empty anchor', async () => {
  const doc = await open(await listFixture(p(`${bullet(0)}<w:rPr><w:b/></w:rPr>`)));
  const { anchor, added } = await continued(doc, '', ['After empty bullet']);
  const mark = pPrOf(anchor).match(/<w:rPr>[\s\S]*?<\/w:rPr>/)![0];
  expect(mark).toBe('<w:rPr><w:b/></w:rPr>');
  expect(rPrsOf(added[0])).toEqual([mark]);
  expect(added[0]).not.toMatch(/<w:b\b[^>]*w:val="(?:0|false|off)"/);
});

test('list continue keeps an italic list style after a hyperlink-only anchor', async () => {
  const style = '<w:style w:type="paragraph" w:styleId="ItalicList"><w:name w:val="Italic List"/><w:basedOn w:val="ListParagraph"/><w:rPr><w:i/></w:rPr></w:style>';
  const anchor = p(bullet(0).replace('ListParagraph', 'ItalicList')).replace('</w:p>', '<w:hyperlink w:anchor="Checklist"><w:r><w:rPr><w:b/></w:rPr><w:t>Linked bullet</w:t></w:r></w:hyperlink></w:p>');
  const doc = await open(await listFixture(anchor, style));
  const { added } = await continued(doc, 'Linked bullet', ['After linked bullet']);
  expect(pPrOf(added[0])).toContain('<w:pStyle w:val="ItalicList"/>');
  expect(added[0]).not.toMatch(/<w:i\b[^>]*w:val="(?:0|false|off)"/);
  expect(rPrsOf(added[0]).join('')).not.toContain('<w:b');
  expect(added[0]).not.toContain('<w:hyperlink');
});

test('list continue preserves adjustRightInd after export and reopen', async () => {
  const doc = await open(await listFixture(p(bullet(0).replace('<w:spacing', '<w:adjustRightInd w:val="0"/><w:spacing'), ['', 'Adjusted bullet'])));
  const { anchor, added } = await continued(doc, 'Adjusted bullet', ['Adjusted continuation']);
  expect(pPrOf(anchor)).toContain('<w:adjustRightInd w:val="0"/>');
  expect(pPrOf(added[0])).toBe(pPrOf(anchor));
  const reopened = await open(await doc.export());
  const segments = reopened.session.storySegments('body');
  const at = segments.findIndex(segment => segment.kind === 'text' && segment.text === 'Adjusted continuation');
  expect(at).toBeGreaterThanOrEqual(0);
  const mark = segments.slice(at + 1).find(segment => segment.kind === 'pilcrow');
  expect(mark).toMatchObject({ kind: 'pilcrow', properties: { adjustRightInd: false } });
});

test('list continue inserts several items in order and drops the anchor page break', async () => {
  const doc = await open();
  const { anchor, added } = await continued(doc, 'First step', ['Step one-a', 'Step one-b']);
  const expected = pPrOf(anchor).replace('<w:pageBreakBefore/>', '');
  expect(pPrOf(anchor)).toContain('<w:pageBreakBefore/>');
  expect(added.map(pPrOf)).toEqual([expected, expected]);
  expect(expected).toContain('<w:keepNext/>');
  expect(added.map(rPrsOf)).toEqual([rPrsOf(anchor), rPrsOf(anchor)]);
  expect(markers(doc).slice(4)).toEqual(['1.', '2.', '3.', '4.', '1.', '2.']);
  const cell = await continued(doc, 'Cell bullet', ['Cell bullet two']);
  expect(pPrOf(cell.added[0])).toBe(pPrOf(cell.anchor));
  expect(doc.read(item(doc, 'Cell bullet two').ref).table).toMatchObject({ row: 1, column: 1 });
});

test('list continue refuses a non-list anchor or a style and changes nothing; omitted list inserts plain paragraphs', async () => {
  const doc = await open();
  const version = doc.list().version;
  const baseline = await parts(await doc.export());
  const heading = item(doc, 'Checklist');
  const refused = async (edits: DocxAgentEdit[], code: string) => {
    const error = await refusal(doc.edit({ version, edits }));
    expect(error.code).toBe(code);
    expect(doc.list().version).toBe(version);
    expect(await parts(await doc.export())).toEqual(baseline);
    return error;
  };
  const plain = await refused([
    { op: 'set_paragraph_style', paragraph: heading.ref, style: 'Heading2' },
    { op: 'insert_paragraphs', after: item(doc, 'Between lists.').ref, list: 'continue', paragraphs: [{ text: 'Not a list' }] },
  ], 'UNSUPPORTED_TARGET');
  expect(plain.details).toMatchObject({ failedIndex: 1 });
  expect(plain.message).toContain('not a list item');
  const styled = await refused([
    { op: 'insert_paragraphs', after: item(doc, 'Beta bullet').ref, list: 'continue', paragraphs: [{ text: 'Styled', style: 'Quote' }] },
  ], 'INVALID_ARGUMENT');
  expect(styled.details).toMatchObject({ failedIndex: 0 });

  await doc.edit({ version, edits: [{ op: 'insert_paragraphs', after: item(doc, 'Beta bullet').ref, paragraphs: [{ text: 'Plain after bullet' }] }] });
  expect(item(doc, 'Plain after bullet').list).toBeUndefined();
  const saved = blocks((await parts(await doc.export())).document).find(paragraph => textOf(paragraph) === 'Plain after bullet')!;
  expect(saved).not.toContain('<w:numPr>');
});

test('list continue refuses list items inside content controls atomically', async () => {
  const zip = await JSZip.loadAsync(await listFixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:sectPr>', `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${p(bullet(0), ['', 'Control bullet'])}</w:sdtContent></w:sdt><w:sectPr>`));
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const version = doc.list().version;
  const baseline = await parts(await doc.export());
  const error = await refusal(doc.edit({ version, edits: [
    { op: 'insert_paragraphs', after: item(doc, 'Control bullet').ref, list: 'continue', paragraphs: [{ text: 'Another control bullet' }] },
  ] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.message).toContain('content controls');
  expect(error.details).toMatchObject({ failedIndex: 0 });
  expect(doc.list().version).toBe(version);
  expect(await parts(await doc.export())).toEqual(baseline);
});
