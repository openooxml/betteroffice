import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { hostSaveMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, yrsToDocument, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { RUN_WITH_ALTERNATE_CONTENT, TOC_ENTRY_WITH_PAGE_REFERENCE, WORD_ROOT } from './__fixtures__/unroundtrippedRuns';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OUTSIDE_BEFORE = '<w:p w:rsidR="00AA0000"><w:pPr><w:pPrChange w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:t>Outside before</w:t></w:r></w:p>';
const OUTSIDE_AFTER = '<w:p><w:r><w:rPr><w:specVanish/></w:rPr><w:t>Outside after</w:t></w:r></w:p>';
const TBL_PR = '<w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblCaption w:val="caption"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>';
const TC_PR = '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcMar><w:left w:w="99" w:type="dxa"/></w:tcMar></w:tcPr>';
const cell = (content: string) => `<w:tc>${TC_PR}${content}</w:tc>`;
const table = (cells: string) => `<w:tbl>${TBL_PR}<w:tr>${cells}</w:tr></w:tbl>`;
const text = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const TABLE = table(cell(text('Cell one')) + cell(text('Cell two')));
const NESTED = table(cell(`${table(cell(text('Nested cell')))}<w:p/>`));
const BLOCK_SDT = `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${text('Control text')}</w:sdtContent></w:sdt>`;
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const makeDocx = (content: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${content}${SECT}</w:body></w:document>`)],
])));
const partOf = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const runOf = (xml: string, value: string) => (xml.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).find(run => run.includes(`>${value}<`)) ?? '';

type Save = (session: YrsSession) => Promise<Uint8Array>;
const SAVES: Record<string, Save> = {
  saveYrsDocx: async session => (await saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  editor: async (session) => {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    return new Uint8Array(await saveEditorDocument(session, projected, projected.package.document.comments ?? []));
  },
};
const open: YrsSession[] = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });
async function seed(content: string) {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(makeDocx(content));
  return session;
}
type Segment = { kind: string; paraId?: string; text?: string; embedKind?: string; attributes?: Record<string, unknown> };
const segmentsOf = (session: YrsSession, story = 'body') => session.storySegments(story) as Segment[];
const paragraphOf = (segments: Segment[], value: string) => {
  const at = segments.findIndex(segment => segment.kind === 'text' && segment.text === value);
  return segments.slice(at).find(segment => segment.kind === 'pilcrow')!.paraId!;
};
const unitLength = (segments: Segment[], paraId: string) => {
  const end = segments.findIndex(segment => segment.kind === 'pilcrow' && segment.paraId === paraId);
  let start = end;
  while (start > 0 && segments[start - 1]!.kind !== 'pilcrow') start--;
  return segments.slice(start, end).reduce((length, segment) => length + (segment.kind === 'text' ? segment.text!.length : 1), 0);
};
const across = (session: YrsSession, from: string, to: string) => {
  const segments = segmentsOf(session);
  const last = paragraphOf(segments, to);
  return { story: 'body', start: { paraId: paragraphOf(segments, from), offset: 0 }, end: { paraId: last, offset: unitLength(segments, last) } };
};
const blockEmbeds = (session: YrsSession) => segmentsOf(session).filter(segment => segment.kind === 'embed' && ['table', 'blockSdt', 'pageBreak', 'columnBreak'].includes(segment.embedKind!));
const SELECTED = `${OUTSIDE_BEFORE}${text('First')}${TABLE}${text('Last')}${OUTSIDE_AFTER}`;

const DELTAS: Array<[string, Record<string, unknown>, string]> = [
  ['bold', { bold: true }, '<w:b/>'],
  ['italic', { italic: true }, '<w:i/>'],
  ['underline', { underline: true }, '<w:u '],
  ['strike', { strike: true }, '<w:strike'],
  ['font size', { fontSize: 14 }, '<w:sz w:val="28"/>'],
  ['colour', { color: { rgb: 'FF0000' } }, '<w:color w:val="FF0000"'],
  ['highlight', { highlight: 'yellow' }, '<w:highlight w:val="yellow"/>'],
  ['font family', { fontFamily: { ascii: 'Arial' } }, 'w:ascii="Arial"'],
];
for (const [saveName, save] of Object.entries(SAVES)) {
  for (const [name, delta, effect] of DELTAS) {
    test(`${saveName}: ${name} across a table formats its cells and keeps everything else`, async () => {
      const session = await seed(SELECTED);
      session.formatRange(across(session, 'First', 'Last'), delta as never);
      expect(blockEmbeds(session).flatMap(embed => Object.keys(embed.attributes ?? {}))).toEqual([]);
      const xml = partOf(await save(session));
      expect(xml).toContain(OUTSIDE_BEFORE);
      expect(xml).toContain(OUTSIDE_AFTER);
      expect(xml).toContain(TBL_PR);
      expect(xml.match(new RegExp(TC_PR.replace(/[/]/g, '\\/'), 'g'))).toHaveLength(2);
      for (const value of ['First', 'Cell one', 'Cell two', 'Last']) expect(runOf(xml, value)).toContain(effect);
      expect(runOf(xml, 'Outside before')).not.toContain(effect);
    });
  }
}

test('bold toggles on across a table, off again, and on for a mixed selection', async () => {
  const session = await seed(`${text('First')}${TABLE}${text('Last')}`);
  const range = across(session, 'First', 'Last');
  session.toggleMark(range, { type: 'bold' } as never);
  let xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Cell one', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
  session.toggleMark(range, { type: 'bold' } as never);
  xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Cell one', 'Last']) expect(runOf(xml, value)).not.toContain('<w:b/>');
  session.formatRange({ ...range, end: { paraId: range.start.paraId, offset: 5 } }, { bold: true });
  session.formatRange({ ...range, start: { paraId: range.end.paraId, offset: 1 } }, { bold: true });
  session.toggleMark(range, { type: 'bold' } as never);
  xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Cell one', 'Cell two', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
});

test('nested tables and block controls take the formatting too', async () => {
  const session = await seed(`${text('First')}${NESTED}${BLOCK_SDT}${text('Last')}`);
  expect(session.formatRange(across(session, 'First', 'Last'), { bold: true }).keptUnchanged).toEqual({});
  expect(blockEmbeds(session).flatMap(embed => Object.keys(embed.attributes ?? {}))).toEqual([]);
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Nested cell', 'Control text', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
});

test('toggling a mark across a table reports the paragraph it kept', async () => {
  for (const mark of [{ type: 'bold' }, { type: 'fontSize', value: 24 }] as const) {
    const session = await createYrsSession();
    open.push(session);
    session.seedFromDocx(rooted(`${text('First')}${table(cell(WITH_TEXT + text('Nested plain')))}${text('Last')}`));
    expect(session.toggleMark(across(session, 'First', 'Last'), mark).keptUnchanged).toEqual({ alternateContent: 1 });
    expect(partOf(await SAVES.saveYrsDocx!(session))).toContain(WITH_TEXT);
  }
});

test('a paragraph with alternate content in a nested table or a block control is kept and reported', async () => {
  const copy = WITH_TEXT.replace(/w14:paraId="[0-9A-F]+"/, 'w14:paraId="0FEDCBA9"').replace(/w14:textId="[0-9A-F]+"/, 'w14:textId="0FEDCBA9"');
  const nested = table(cell(`${table(cell(WITH_TEXT + text('Nested plain')))}<w:p/>`));
  const control = `<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>${copy}${text('Control plain')}</w:sdtContent></w:sdt>`;
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(rooted(`${text('First')}${nested}${control}${text('Last')}`));
  expect(session.formatRange(across(session, 'First', 'Last'), { bold: true }).keptUnchanged).toEqual({ alternateContent: 2 });
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  expect(xml).toContain(WITH_TEXT);
  expect(xml).toContain(copy);
  for (const value of ['First', 'Nested plain', 'Control plain', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
});

const WITH_TEXT = RUN_WITH_ALTERNATE_CONTENT.replace(/<\/w:p>$/, '<w:r><w:t>Beside the drawing</w:t></w:r></w:p>');
const rooted = (content: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`${WORD_ROOT}<w:body>${content}${SECT}</w:body></w:document>`)],
])));
for (const [saveName, save] of Object.entries(SAVES)) {
  test(`${saveName}: bold across a table keeps a real cell paragraph with alternate content byte for byte`, async () => {
    const session = await createYrsSession();
    open.push(session);
    session.seedFromDocx(rooted(`${text('First')}${table(cell(WITH_TEXT) + cell(text('Plain cell')))}${text('Last')}`));
    const receipt = session.formatRange(across(session, 'First', 'Last'), { bold: true });
    expect(receipt.keptUnchanged).toEqual({ alternateContent: 1 });
    const xml = partOf(await save(session));
    expect(xml).toContain(WITH_TEXT);
    for (const value of ['First', 'Plain cell', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
  });
}

test('clearing across a table keeps a cell wrapped in a content control and reports it', async () => {
  const bold = (value: string) => `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${value}</w:t></w:r></w:p>`;
  const wrapped = `<w:sdt><w:sdtPr><w:id w:val="9"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtEndPr/><w:sdtContent>${cell(bold('Box'))}</w:sdtContent></w:sdt>`;
  const session = await seed(`${bold('First')}${table(cell(bold('Label')) + wrapped)}${bold('Last')}`);
  expect(session.clearFormatting(across(session, 'First', 'Last')).keptUnchanged).toEqual({ contentControl: 1 });
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  expect(xml).toContain(wrapped);
  for (const value of ['First', 'Label', 'Last']) expect(runOf(xml, value)).not.toContain('<w:b/>');
});

const TOC_END_PARAGRAPH = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
for (const [saveName, save] of Object.entries(SAVES)) {
  test(`${saveName}: clearing across the body keeps a real table of contents entry and reports it`, async () => {
    const session = await createYrsSession();
    open.push(session);
    session.seedFromDocx(rooted(`${text('First')}${TOC_ENTRY_WITH_PAGE_REFERENCE}${TOC_END_PARAGRAPH}${text('Last')}`));
    expect(session.clearFormatting(across(session, 'First', 'Last')).keptUnchanged).toEqual({ field: 1 });
    const xml = partOf(await save(session));
    expect(xml).toContain(TOC_ENTRY_WITH_PAGE_REFERENCE);
    expect(xml.match(/PAGEREF/g)?.length).toBe(1);
  });
}

for (const [label, run] of [
  ['clearing', (session: YrsSession, range: ReturnType<typeof across>) => session.clearFormatting(range)],
  ['bolding', (session: YrsSession, range: ReturnType<typeof across>) => session.formatRange(range, { bold: true })],
] as const) {
  test(`${label} inside one table of contents entry keeps it and reports it`, async () => {
    const session = await createYrsSession();
    open.push(session);
    session.seedFromDocx(rooted(`${text('First')}${TOC_ENTRY_WITH_PAGE_REFERENCE}${TOC_END_PARAGRAPH}${text('Last')}`));
    const entry = (session.paragraphs('body') as Array<{ paraId: string; text: string }>)[1]!;
    const range = { story: 'body', start: { paraId: entry.paraId, offset: 0 }, end: { paraId: entry.paraId, offset: entry.text.length } };
    expect(run(session, range).keptUnchanged).toEqual({ field: 1 });
    const xml = partOf(await SAVES.saveYrsDocx!(session));
    expect(xml).toContain(TOC_ENTRY_WITH_PAGE_REFERENCE);
    expect(xml.match(/PAGEREF/g)?.length).toBe(1);
  });
}

test('formatting with nothing kept reports an empty receipt', async () => {
  const session = await seed(SELECTED);
  expect(session.formatRange(across(session, 'First', 'Last'), { bold: true }).keptUnchanged).toEqual({});
  expect(session.clearFormatting(across(session, 'First', 'Last')).keptUnchanged).toEqual({});
});

test('a page break in the selection is not marked', async () => {
  const session = await seed(`${text('First')}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${text('Last')}`);
  session.formatRange(across(session, 'First', 'Last'), { bold: true });
  expect(blockEmbeds(session).flatMap(embed => Object.keys(embed.attributes ?? {}))).toEqual([]);
  expect(partOf(await SAVES.saveYrsDocx!(session))).toContain('<w:br w:type="page"/>');
});

test('clearing across a table clears its cells and keeps its structure', async () => {
  const bold = (value: string) => `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${value}</w:t></w:r></w:p>`;
  const session = await seed(`${bold('First')}${table(cell(bold('Cell one')))}${bold('Last')}`);
  session.clearFormatting(across(session, 'First', 'Last'));
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Cell one', 'Last']) expect(runOf(xml, value)).not.toContain('<w:b/>');
  expect(xml).toContain(TBL_PR);
  expect(xml).toContain(TC_PR);
});

test('one undo restores the body and the cells', async () => {
  const session = await seed(`${text('First')}${TABLE}${text('Last')}`);
  const before = partOf(await SAVES.saveYrsDocx!(session));
  session.formatRange(across(session, 'First', 'Last'), { bold: true });
  expect(partOf(await SAVES.saveYrsDocx!(session))).not.toBe(before);
  session.undo();
  expect(partOf(await SAVES.saveYrsDocx!(session))).toBe(before);
});

test('a selection inside one paragraph beside a table leaves the table alone', async () => {
  const session = await seed(`${text('First')}${TABLE}${text('Last')}`);
  const range = across(session, 'First', 'First');
  session.formatRange(range, { bold: true });
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  expect(runOf(xml, 'First')).toContain('<w:b/>');
  expect(xml).toContain(TABLE);
});

test('the worker save takes the formatting into cells', async () => {
  const bytes = makeDocx(SELECTED);
  const worker = startWorker(1);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const { hostJson } = await client.open(bytes);
    const host = decodeDocxHostJson(hostJson, bytes).document;
    const resident = worker.sessions[0]!;
    const segments = resident.storySegments('body') as Segment[];
    resident.formatRange({ story: 'body', start: { paraId: paragraphOf(segments, 'First'), offset: 0 }, end: { paraId: paragraphOf(segments, 'Last'), offset: unitLength(segments, paragraphOf(segments, 'Last')) } }, { bold: true });
    const saved = await client.save({ comments: host.package.document.comments ?? [], host: hostSaveMetadata(host) });
    const xml = partOf(new Uint8Array(saved.bytes));
    expect(xml).toContain(OUTSIDE_BEFORE);
    expect(xml).toContain(OUTSIDE_AFTER);
    expect(xml).toContain(TBL_PR);
    for (const value of ['First', 'Cell one', 'Cell two', 'Last']) expect(runOf(xml, value)).toContain('<w:b/>');
  } finally { client.destroy(); }
}, 60_000);

test('suggesting bold across a table tracks the change in the cells too', async () => {
  const session = await seed(`${text('First')}${TABLE}${text('Last')}`);
  session.formatRange(across(session, 'First', 'Last'), { bold: true }, { name: 'Reviewer', date: '2026-01-01T00:00:00Z' } as never);
  expect(blockEmbeds(session).flatMap(embed => Object.keys(embed.attributes ?? {}))).toEqual([]);
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  for (const value of ['First', 'Cell one', 'Cell two', 'Last']) {
    expect(runOf(xml, value)).toContain('<w:b/>');
    expect(runOf(xml, value)).toContain('<w:rPrChange');
  }
  expect(xml).toContain(TBL_PR);
});

const FIELD_CELL = '<w:p><w:bookmarkStart w:id="0" w:name="target"/><w:r><w:t>Anchor</w:t></w:r><w:bookmarkEnd w:id="0"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "target" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="target"><w:r><w:rPr><w:b/></w:rPr><w:t>Link</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
const fieldResultStories = (session: YrsSession) => (session.listStories().stories as Array<{ story: string }>)
  .filter(({ story }) => (session.storySegments(story) as Segment[]).some(segment => segment.kind === 'text' && segment.attributes?.fieldResult !== undefined))
  .map(({ story }) => story);

test('a refused suggestion across a table with a field in a cell changes nothing', async () => {
  const session = await seed(`${text('First')}${table(cell(FIELD_CELL))}${text('Last')}`);
  const before = partOf(await SAVES.saveYrsDocx!(session));
  const state = session.encodeState();
  expect(() => session.formatRange(across(session, 'First', 'Last'), { bold: true }, { name: 'Reviewer', date: '2026-01-01T00:00:00Z' } as never)).toThrow();
  expect(session.encodeState()).toEqual(state);
  expect(partOf(await SAVES.saveYrsDocx!(session))).toBe(before);
});

test('clearing across a table keeps a field in a cell whole', async () => {
  const session = await seed(`${text('First')}${table(cell(FIELD_CELL))}${text('Last')}`);
  const projected = fieldResultStories(session);
  expect(projected.length).toBeGreaterThan(0);
  session.clearFormatting(across(session, 'First', 'Last'));
  expect(fieldResultStories(session)).toEqual(projected);
  const xml = partOf(await SAVES.saveYrsDocx!(session));
  const begin = xml.indexOf('w:fldCharType="begin"');
  const separate = xml.indexOf('w:fldCharType="separate"');
  const link = xml.indexOf('>Link<');
  const end = xml.indexOf('w:fldCharType="end"');
  expect(begin).toBeGreaterThan(-1);
  expect(begin < separate && separate < link && link < end).toBe(true);
  expect(xml).toContain('HYPERLINK');
  const hyperlink = xml.indexOf('<w:hyperlink');
  expect(separate < hyperlink && hyperlink < link && link < end).toBe(true);
  expect(runOf(xml, 'Link')).not.toContain('<w:b/>');
});

const CELL_CONTENT: Array<[string, string, string[]]> = [
  ['an opaque complex field', '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>', ['w:fldCharType="begin"', 'w:fldCharType="separate"', 'w:fldCharType="end"', ' PAGE ']],
  ['a simple field', '<w:p><w:r><w:t xml:space="preserve">Date </w:t></w:r><w:fldSimple w:instr=" DATE "><w:r><w:t>1/1/2026</w:t></w:r></w:fldSimple></w:p>', ['<w:fldSimple w:instr=" DATE "']],
  ['a hyperlink', '<w:p><w:hyperlink w:anchor="target"><w:r><w:rPr><w:b/></w:rPr><w:t>Linked</w:t></w:r></w:hyperlink><w:bookmarkStart w:id="3" w:name="target"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>', ['<w:hyperlink w:anchor="target"', '<w:bookmarkStart w:id="3"', '<w:bookmarkEnd w:id="3"']],
  ['an inline content control', '<w:p><w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:i/></w:rPr><w:t>Control</w:t></w:r></w:sdtContent></w:sdt></w:p>', ['<w:sdt>', '<w:id w:val="9"/>', '</w:sdtContent>']],
  ['a tab', '<w:p><w:r><w:t>Left</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Right</w:t></w:r></w:p>', ['<w:tab/>']],
  ['a text box', '<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="box" style="width:100pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>', ['<w:pict>', '<w:txbxContent>', '>Boxed<']],
];
for (const [name, content, kept] of CELL_CONTENT) {
  for (const [label, run] of [
    ['clearing', (session: YrsSession, range: ReturnType<typeof across>) => session.clearFormatting(range)],
    ['bolding', (session: YrsSession, range: ReturnType<typeof across>) => session.formatRange(range, { bold: true })],
  ] as const) {
    test(`${label} across a table keeps ${name} in a cell`, async () => {
      const source = `${text('First')}${table(cell(content))}${text('Last')}`;
      const session = await seed(source);
      run(session, across(session, 'First', 'Last'));
      const xml = partOf(await SAVES.saveYrsDocx!(session));
      for (const token of kept) expect(xml).toContain(token);
      expect(xml.match(/<w:p\b/g)?.length).toBe(source.match(/<w:p\b/g)?.length);
      expect(xml.match(/<w:tc>/g)?.length).toBe(source.match(/<w:tc>/g)?.length);
    });
  }
}

const BOLD = '<w:rPr><w:b/></w:rPr>';
const run = (content: string) => `<w:r>${BOLD}${content}</w:r>`;
const bold = (value: string) => run(`<w:t>${value}</w:t>`);
const fieldChar = (kind: string) => run(`<w:fldChar w:fldCharType="${kind}"/>`);
const code = (value: string) => run(`<w:instrText xml:space="preserve"> ${value} </w:instrText>`);
const SMART = 'w:uri="urn:schemas-microsoft-com:office:smarttags"';
const TOC_ENTRY = `<w:p>${fieldChar('begin')}${code('TOC \\o "1-3"')}${fieldChar('separate')}<w:hyperlink w:anchor="_Toc1" w:history="1">${bold('Entry')}${run('<w:tab/>')}${fieldChar('begin')}${code('PAGEREF _Toc1')}${fieldChar('separate')}${bold('2')}${fieldChar('end')}</w:hyperlink></w:p><w:p>${fieldChar('end')}</w:p>`;
const FIELD_END_IN_CELL = table(cell(`<w:p><w:bookmarkStart w:id="9" w:name="b"/>${fieldChar('begin')}${code('REF b')}${fieldChar('separate')}${bold('Value')}</w:p><w:p>${bold('Result')}${fieldChar('end')}<w:bookmarkEnd w:id="9"/></w:p>`));
const SMART_TAG_IN_CELL = table(cell(`<w:p>${bold('Near ')}<w:smartTag ${SMART} w:element="place"><w:smartTag ${SMART} w:element="PlaceName">${bold('Ancient')}</w:smartTag></w:smartTag>${bold(' Monument')}</w:p>`));
const TWO_PAGE_BREAKS = `<w:p>${run('<w:br w:type="page"/>')}<w:r>${BOLD}<w:br w:type="page"/><w:t>A</w:t></w:r></w:p>`;
const CLEAR_SHAPES: Array<[string, string]> = [
  ['a table of contents entry', TOC_ENTRY],
  ['a field ending in a table cell', FIELD_END_IN_CELL],
  ['smart tags in a table cell', SMART_TAG_IN_CELL],
  ['alternate content in a table cell', table(cell(WITH_TEXT))],
  ['two page breaks', TWO_PAGE_BREAKS],
];
const STRUCTURE_TAGS = ['w:hyperlink', 'w:fldChar', 'w:instrText', 'w:smartTag', 'mc:AlternateContent', 'mc:Fallback', 'w:pict', 'w:txbxContent', 'w:tab', 'w:br', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:p'];
const tagCounts = (xml: string) => Object.fromEntries(STRUCTURE_TAGS.map(tag => [tag, xml.match(new RegExp(`<${tag}[\\s/>]`, 'g'))?.length ?? 0]));
for (const [name, shape] of [...CLEAR_SHAPES, ['all of them together', CLEAR_SHAPES.map(([, content]) => content).join('')] as [string, string]]) {
  test(`clearing formatting around ${name} loses no structure`, async () => {
    const session = await createYrsSession();
    open.push(session);
    const source = `<w:p>${bold('First')}</w:p>${shape}<w:p>${bold('Last')}</w:p>`;
    session.seedFromDocx(rooted(source));
    session.clearFormatting(across(session, 'First', 'Last'));
    const plan = session.paragraphSavePlan() as { splicedParts?: Array<{ part: string }>; wholeParts?: Array<{ part: string }> };
    expect([...(plan.splicedParts ?? []), ...(plan.wholeParts ?? [])].map(({ part }) => part)).toContain('word/document.xml');
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    const was = tagCounts(`${WORD_ROOT}<w:body>${source}${SECT}</w:body></w:document>`), now = tagCounts(xml);
    for (const tag of STRUCTURE_TAGS) expect(`${tag} ${now[tag]! >= was[tag]!}`).toBe(`${tag} true`);
    if (!saved.keptUnchanged) expect(runOf(xml, 'First')).not.toContain('<w:b/>');
  });
}

test('formatting a part the parser does not read as written keeps it and says so when writing it would drop structure', async () => {
  const session = await createYrsSession();
  open.push(session);
  const source = `<w:p>${bold('First')}</w:p>${SMART_TAG_IN_CELL}<w:p>${bold('Last')}</w:p>`;
  session.seedFromDocx(rooted(source));
  session.clearFormatting(across(session, 'First', 'Last'));
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  expect(new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml'])).toContain(source);
  expect(saved.keptUnchanged?.paragraphs.every(paragraph => paragraph.scope === 'part' && paragraph.dropped.includes('w:smartTag'))).toBe(true);
  expect(saved.keptUnchanged?.structure).toBeGreaterThan(0);
  expect(saved.structureDropped).toBeUndefined();
});

test('formatting a part the parser does not read as written is saved when writing it loses nothing', async () => {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(rooted(`<w:p>${bold('First')}</w:p><w:p>${bold('Near ')}${bold('Monument')}</w:p><w:p>${bold('Last')}</w:p>`));
  session.clearFormatting(across(session, 'First', 'Last'));
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  expect(new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml'])).not.toContain('<w:b/>');
  expect(saved.keptUnchanged).toBeUndefined();
});

const BORDERED = `<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tblPrEx><w:tblBorders><w:insideV w:val="single"/></w:tblBorders></w:tblPrEx><w:tr><w:tblPrEx><w:tblBorders><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPrEx><w:tc><w:tcPr><w:shd w:val="clear" w:fill="112233"/></w:tcPr><w:p>${bold('Near ')}<w:smartTag ${SMART} w:element="place">${bold('Ancient')}</w:smartTag></w:p></w:tc></w:tr></w:tbl>`.replace('<w:tblPrEx><w:tblBorders><w:insideV w:val="single"/></w:tblBorders></w:tblPrEx><w:tr>', '<w:tr>');
const ALL_ELEMENTS = (xml: string) => { const counts: Record<string, number> = {}; for (const [, name] of xml.matchAll(/<([a-zA-Z0-9]+:[a-zA-Z0-9]+)[\s/>]/g)) counts[name!] = (counts[name!] ?? 0) + 1; return counts; };
for (const [name, shape] of [['a bordered table with smart tags', BORDERED], ...CLEAR_SHAPES] as Array<[string, string]>) {
  test(`clearing formatting around ${name} beside a section edit loses no structure and applies the margin`, async () => {
    const session = await createYrsSession();
    open.push(session);
    const source = `<w:p>${bold('First')}</w:p>${shape}<w:p>${bold('Last')}</w:p>`;
    session.seedFromDocx(rooted(source));
    session.clearFormatting(across(session, 'First', 'Last'));
    session.setSectionProperties('final', { margins: { left: 1800 } });
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    expect(xml).toContain('w:left="1800"');
    const was = tagCounts(`${WORD_ROOT}<w:body>${source}${SECT}</w:body></w:document>`), now = tagCounts(xml);
    for (const tag of STRUCTURE_TAGS) expect(`${tag} ${now[tag]! >= was[tag]!}`).toBe(`${tag} true`);
    const before = ALL_ELEMENTS(shape.replace(/<w:rPr>.*?<\/w:rPr>/g, '')), after = ALL_ELEMENTS(xml);
    for (const tag of ['w:tblBorders', 'w:tblPrEx', 'w:insideH', 'w:insideV', 'w:shd']) expect(`${tag} ${(after[tag] ?? 0) >= (before[tag] ?? 0)}`).toBe(`${tag} true`);
    expect(saved.structureDropped).toBeUndefined();
  });
}

