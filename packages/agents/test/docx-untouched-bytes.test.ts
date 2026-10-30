import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { ALTERNATE_TEXTBOX_RUN, bodyMarkerBody, fieldLabelBody, offsetTableBody, orphanTableMarkersBody, paragraphsOf, plainTableBody, textNodes, tocSdtBody, V, WORD_2007_ROOTS, word2007Body, WPS } from '../../docx/src/yrs/__fixtures__/untouchedBytes';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT = `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" xmlns:r="${R}" xmlns:v="${V}" xmlns:wps="${WPS}">`;

const rich = (id: string, text: string) =>
  `<w:p w14:paraId="${id}" w:rsidR="00752A6D" w:rsidRDefault="00593C67"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidR="00761B10"><w:rPr><w:color w:val="auto"/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:br/></w:r><w:proofErr w:type="spellEnd"/></w:p>`;

async function docx(body: string, extra: Record<string, string> = {}, rels = '', root = ROOT): Promise<{ bytes: Uint8Array; xml: string }> {
  const zip = new JSZip();
  const overrides = Object.keys(extra).map(part => {
    const kind = part.includes('footnotes') ? 'footnotes+xml' : part.includes('comments') ? 'comments+xml' : 'styles+xml';
    return `<Override PartName="/${part}" ContentType="${OFFICE}.${kind}"/>`;
  }).join('');
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>${overrides}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  if (rels) zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  const xml = `${root}<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  zip.file('word/document.xml', xml);
  for (const [part, content] of Object.entries(extra)) zip.file(part, content);
  return { bytes: await zip.generateAsync({ type: 'uint8array' }), xml };
}

const styles = (normalColor = '') => `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>${normalColor}</w:rPr></w:style></w:styles>`;

async function exportedXml(bytes: Uint8Array, part = 'word/document.xml'): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file(part)!.async('string');
}

function expectUntouched(sourceXml: string, outputXml: string, edited: string[] = []) {
  for (const paragraph of paragraphsOf(sourceXml)) {
    if (edited.includes(paragraph)) continue;
    expect(outputXml).toContain(paragraph);
  }
}

test.each(WORD_2007_ROOTS)('Word 2007 $name text replacement keeps every untouched paragraph', async ({ root }) => {
  const source = await docx(word2007Body(), {}, '', root);
  const doc = await openDocx(source.bytes);
  try {
    const match = doc.grep({ query: 'target' }).matches[0]!;
    const proposal = await doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'changed' }] });
    await doc.accept(proposal.id, { tracked: false });
    const bytes = await doc.export();
    const xml = await exportedXml(bytes);
    expectUntouched(source.xml, xml, [paragraphsOf(source.xml)[0]!]);
    expect(paragraphsOf(xml)).toHaveLength(2);
    expect(paragraphsOf(xml).map(paragraph => textNodes(paragraph).join(''))).toEqual(['Replace changed words', 'Retained paragraph words']);
    const writtenRoot = xml.match(/<w:document\b[^>]*>/)![0];
    for (const declaration of root.matchAll(/ xmlns(?::[^=]+)?="[^"]*"/g)) expect(writtenRoot).toContain(declaration[0]);
    const reopened = await openDocx(bytes);
    try { expect(reopened.grep({ query: 'Replace changed words' }).matches).toHaveLength(1); } finally { reopened.close(); }
  } finally { doc.close(); }
});

test.each(WORD_2007_ROOTS)('Word 2007 $name paragraph insertion keeps every untouched paragraph across saves', async ({ root }) => {
  const source = await docx(word2007Body(), {}, '', root);
  let bytes = source.bytes;
  let before = source.xml;
  for (const text of ['Inserted paragraph.', 'Inserted again.']) {
    const doc = await openDocx(bytes);
    try {
      const item = doc.list().items.find(entry => entry.text.startsWith('Replace target words'))!;
      await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: item.ref, paragraphs: [{ text }] }] });
      bytes = await doc.export();
      const xml = await exportedXml(bytes);
      expectUntouched(before, xml);
      expect(paragraphsOf(xml)).toHaveLength(paragraphsOf(before).length + 1);
      expect(textNodes(xml)).toEqual(['Replace target words', text, ...textNodes(before).slice(1)]);
      const writtenRoot = xml.match(/<w:document\b[^>]*>/)![0];
      for (const declaration of root.matchAll(/ xmlns(?::[^=]+)?="[^"]*"/g)) expect(writtenRoot).toContain(declaration[0]);
      expect(writtenRoot).toContain(root.includes('xmlns:mc=') ? 'mc:Ignorable="w14"' : root.includes('ve:Ignorable=') ? 've:Ignorable="draw legacy w14"' : 've:Ignorable="w14"');
      before = xml;
    } finally { doc.close(); }
  }
});

test.each([1, 3])('editing %i TOC labels keeps PAGEREF fields and hyperlinks byte for byte', async (count) => {
  const source = await docx(fieldLabelBody());
  const doc = await openDocx(source.bytes);
  try {
    const matches = doc.grep({ query: 'Background' }).matches.slice(0, count);
    expect(matches).toHaveLength(count);
    const proposal = await doc.propose({ author: 'Test', edits: matches.map(match => ({ match: match.match, newText: 'Context' })) });
    await doc.accept(proposal.id, { tracked: false });
    const part = doc.session.paragraphSavePlan().splicedParts!.find(part => part.part === 'word/document.xml')!;
    expect(part.legacyParagraphs ?? []).toEqual([]);
    expect(part.paragraphPatches).toHaveLength(count);
    const bytes = await doc.export();
    const xml = await exportedXml(bytes);
    let expected = source.xml;
    for (let index = 1; index <= count; index++) expected = expected.replace(`Background ${index}`, `Context ${index}`);
    expect(xml).toBe(expected);
    for (const [tag, total] of [['w:fldChar', 12], ['w:instrText', 4], ['w:hyperlink', 3]] as const) {
      expect(xml.match(new RegExp(`<${tag}\\b`, 'g'))).toHaveLength(total);
    }
    const reopened = await openDocx(bytes);
    try { expect(reopened.grep({ query: 'Context' }).matches).toHaveLength(count); } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('editing beside AlternateContent keeps the complete text box run and VML fallback', async () => {
  const edited = `<w:p w14:paraId="40000001"><w:r><w:t xml:space="preserve">Caption </w:t></w:r>${ALTERNATE_TEXTBOX_RUN}</w:p>`;
  const source = await docx(`${edited}${rich('40000002', 'Untouched paragraph')}`);
  const doc = await openDocx(source.bytes);
  try {
    const match = doc.grep({ query: 'Caption' }).matches[0]!;
    const proposal = await doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'Title' }] });
    await doc.accept(proposal.id, { tracked: false });
    const part = doc.session.paragraphSavePlan().splicedParts!.find(part => part.part === 'word/document.xml')!;
    expect(part.legacyParagraphs ?? []).toEqual([]);
    expect(part.paragraphPatches).toHaveLength(1);
    const bytes = await doc.export();
    const xml = await exportedXml(bytes);
    expect(xml).toBe(source.xml.replace('Caption', 'Title'));
    expect(xml).toContain(ALTERNATE_TEXTBOX_RUN);
    for (const [tag, total] of [['mc:AlternateContent', 1], ['mc:Fallback', 1], ['w:pict', 1], ['w:txbxContent', 2]] as const) {
      expect(xml.match(new RegExp(`<${tag}\\b`, 'g'))).toHaveLength(total);
    }
    const reopened = await openDocx(bytes);
    try { expect(reopened.grep({ query: 'Title' }).matches).toHaveLength(1); } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('ve root keeps ligature extensions in edited and untouched paragraphs', async () => {
  const rpr = '<w:rPr><w14:ligatures w14:val="standardContextual"/></w:rPr>';
  const untouched = `<w:p w14:paraId="50000001"><w:r>${rpr}<w:t>Untouched ligature paragraph.</w:t></w:r></w:p>`;
  const edited = `<w:p w14:paraId="50000002"><w:r>${rpr}<w:t xml:space="preserve">Edited ligature </w:t></w:r><w:r>${rpr}<w:t>paragraph here.</w:t></w:r></w:p>`;
  const root = `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:ve="${MC}" ve:Ignorable="w14">`;
  const source = await docx(`${untouched}${edited}`, {}, '', root);
  const doc = await openDocx(source.bytes);
  try {
    const match = doc.grep({ query: 'Edited' }).matches[0]!;
    const proposal = await doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'Changed' }] });
    await doc.accept(proposal.id, { tracked: false });
    const xml = await exportedXml(await doc.export());
    expect(xml).toBe(source.xml.replace('Edited', 'Changed'));
    expect(xml).toContain(untouched);
    expect(xml.match(/<w14:ligatures\b/g)).toHaveLength(3);
  } finally { doc.close(); }
});

test('comment range spanning table cells keeps untouched paragraphs byte for byte', async () => {
  const heading = rich('10000001', 'Title words here');
  const second = rich('10000002', 'Second heading words');
  const cellA = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>A</w:t></w:r></w:p>';
  const cellB = '<w:p><w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>';
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cellA}</w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cellB}</w:tc></w:tr></w:tbl>`;
  const comments = `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Ada"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>`;
  const source = await docx(`${heading}${second}${table}`, { 'word/comments.xml': comments }, `<Relationship Id="comments" Type="${R}/comments" Target="comments.xml"/>`);
  const doc = await openDocx(source.bytes);
  try {
    const item = doc.list().items.find(entry => entry.text.includes('Title words'))!;
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: item.ref, paragraphs: [{ text: 'Inserted paragraph.' }] }] });
    const xml = await exportedXml(await doc.export());
    expectUntouched(source.xml, xml);
    expect(xml).toContain(table);
    const reopened = await openDocx(await doc.export());
    try { expect(reopened.list().items.map(entry => entry.text)).toContain('Inserted paragraph.'); } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('merge_cells keeps untouched paragraphs byte for byte', async () => {
  const title = rich('10000001', 'Merged cells title');
  const head = '<w:p><w:r><w:t>Parameter</w:t></w:r></w:p>';
  const cellA = '<w:p><w:r><w:t>a</w:t></w:r></w:p>';
  const cellB = '<w:p><w:r><w:t>b</w:t></w:r></w:p>';
  const after = rich('10000002', 'After the table');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${head}</w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cellA}</w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cellB}</w:tc></w:tr></w:tbl>`;
  const source = await docx(`${title}${table}${after}`);
  const doc = await openDocx(source.bytes);
  try {
    const items = doc.list().items;
    const from = items.find(entry => entry.text === 'a')!;
    const to = items.find(entry => entry.text === 'b')!;
    await doc.edit({ version: doc.list().version, edits: [{ op: 'merge_cells', from: from.ref, to: to.ref }] });
    const bytes = await doc.export();
    const xml = await exportedXml(bytes);
    expectUntouched(source.xml, xml, [cellA, cellB]);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    const reopened = await openDocx(bytes);
    try {
      const texts = reopened.list().items.map(entry => entry.text);
      expect(texts).toContain('a');
      expect(texts).toContain('b');
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

async function expectInsertedBody(body: string, anchorText: string | null) {
  const source = await docx(body);
  const doc = await openDocx(source.bytes);
  try {
    const item = doc.list().items.find(entry => anchorText === null
      ? entry.story === 'body' && entry.text === '\ufffc'
      : anchorText === '' ? entry.story === 'body' && entry.text === '' : entry.text.startsWith(anchorText))!;
    expect(item).toBeDefined();
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: item.ref, paragraphs: [{ text: 'Inserted paragraph.' }] }] });
    const bytes = await doc.export();
    const xml = await exportedXml(bytes);
    expectUntouched(source.xml, xml);
    const sourceParagraphs = paragraphsOf(source.xml);
    expect(paragraphsOf(xml)).toHaveLength(sourceParagraphs.length + (anchorText === null ? 2 : 1));
    if (anchorText === null) {
      const added = paragraphsOf(xml).filter(paragraph => !sourceParagraphs.includes(paragraph));
      expect(added).toHaveLength(2);
      const empty = added.filter(paragraph => textNodes(paragraph).length === 0);
      expect(empty).toHaveLength(1);
      expect(empty[0]).toMatch(/^<w:p\b[^>]*w14:paraId=/);
    }
    for (const tag of ['w:txbxContent', 'mc:AlternateContent', 'mc:Fallback', 'w:bookmarkStart', 'w:bookmarkEnd']) {
      expect(xml.match(new RegExp(`<${tag}\\b`, 'g'))?.length ?? 0).toBe(source.xml.match(new RegExp(`<${tag}\\b`, 'g'))?.length ?? 0);
    }
    expect(textNodes(xml).filter(text => text !== 'Inserted paragraph.')).toEqual(textNodes(source.xml));
    expect(textNodes(xml).filter(text => text === 'Inserted paragraph.')).toHaveLength(1);
    const inserted = paragraphsOf(xml).find(paragraph => textNodes(paragraph).includes('Inserted paragraph.'))!;
    expect(inserted).not.toMatch(/<w:(?:br|drawing|pict)\b/);
    const reopened = await openDocx(bytes);
    try { expect(reopened.list().items.map(entry => entry.text.replace(/^\ufffc+/, ''))).toContain('Inserted paragraph.'); } finally { reopened.close(); }
    return xml;
  } finally { doc.close(); }
}

test('block sdt with a trailing page break and TOC fields keeps every source paragraph', async () => {
  const body = tocSdtBody();
  const xml = await expectInsertedBody(body, 'Contents anchor');
  expect(xml).toContain(body.slice(body.indexOf('<w:sdt>')));
  expect(xml.match(/<w:br w:type="page"\/>/g)).toHaveLength(1);
});

test('body bookmark end beside an empty anchor keeps every source paragraph', async () => {
  await expectInsertedBody(bodyMarkerBody(), '');
});

test('orphan bookmark ends between rows and cells keep every source paragraph', async () => {
  const body = orphanTableMarkersBody();
  const xml = await expectInsertedBody(body, 'Table anchor');
  expect(xml).toContain(body.slice(body.indexOf('<w:tbl>'), body.indexOf('</w:tbl>') + 8));
});

test('gridAfter and spanned rows keep every source paragraph on table and cell inserts', async () => {
  const body = offsetTableBody();
  const xml = await expectInsertedBody(body, null);
  expect(xml).toContain(body.slice(body.indexOf('<w:tbl>')));
  const tableEnd = xml.indexOf('</w:tbl>');
  expect(xml.indexOf('Inserted paragraph.')).toBeGreaterThan(tableEnd);
  const cellXml = await expectInsertedBody(body, 'sample');
  const sample = cellXml.indexOf('<w:t>sample</w:t>');
  expect(cellXml.indexOf('Inserted paragraph.')).toBeGreaterThan(sample);
  expect(cellXml.indexOf('Inserted paragraph.')).toBeLessThan(cellXml.indexOf('</w:tc>', sample));
});

test.each([
  { name: 'plain', body: plainTableBody },
  { name: 'gridAfter/gridSpan', body: offsetTableBody },
])('insert directly after a $name table keeps every source byte', async ({ body }) => {
  const source = body();
  const xml = await expectInsertedBody(source, null);
  const end = xml.indexOf('</w:tbl>') + '</w:tbl>'.length;
  expect(xml).toContain(source.slice(source.indexOf('<w:tbl>')));
  expect(xml.slice(end)).toMatch(/^<w:p\b[^>]*>/);
  expect(xml.indexOf('Inserted paragraph.')).toBeGreaterThan(end);
});

test('edited footnote keeps direct color auto and untouched runs byte for byte', async () => {
  const body = '<w:p><w:r><w:t>Body text with a note.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>';
  const runA = '<w:r><w:rPr><w:color w:val="auto"/><w:lang w:val="en-CA"/></w:rPr><w:t xml:space="preserve">A fiscal year runs from </w:t></w:r>';
  const runB = '<w:r><w:rPr><w:color w:val="auto"/><w:lang w:val="en-CA"/></w:rPr><w:t>April 1</w:t></w:r>';
  const runC = '<w:r><w:rPr><w:lang w:val="en-CA"/></w:rPr><w:t xml:space="preserve"> to March 31.</w:t></w:r>';
  const footnote = `<w:p>${runA}${runB}${runC}</w:p>`;
  const footnotes = `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1">${footnote}</w:footnote></w:footnotes>`;
  const source = await docx(body, { 'word/footnotes.xml': footnotes, 'word/styles.xml': styles('<w:color w:val="54575A"/>') },
    `<Relationship Id="footnotes" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>`);
  const doc = await openDocx(source.bytes);
  try {
    const match = doc.grep({ query: 'April 1' }).matches[0]!;
    const proposal = await doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'April 2' }] });
    await doc.accept(proposal.id);
    const bytes = await doc.export();
    const xml = await exportedXml(bytes, 'word/footnotes.xml');
    expectUntouched(footnotes, xml, [footnote]);
    expect(xml).toContain('<w:color w:val="auto"/>');
    expect(xml).toContain(runA);
    expect(xml).toContain(runC);
    expect(xml).toContain('April 2');
    const edited = paragraphsOf(xml).find(paragraph => paragraph.includes('April 2'))!;
    const donorRPr = runB.match(/<w:rPr>.*?<\/w:rPr>/)![0];
    const runs = edited.match(/<w:r[ >][\s\S]*?<\/w:r>/g)!;
    expect(runs.length).toBeGreaterThanOrEqual(footnote.match(/<w:r[ >]/g)!.length);
    for (const run of runs) {
      const rPr = run.match(/<w:rPr>.*?<\/w:rPr>/)?.[0] ?? '';
      expect([runA, runB, runC].map(source => source.match(/<w:rPr>.*?<\/w:rPr>/)?.[0] ?? '')).toContain(rPr);
    }
    expect(runs.find(run => run.includes('April 2'))).toContain(donorRPr);
    expect(runs.map(run => run.replace(/<[^>]+>/g, '')).join('')).toBe('A fiscal year runs from April 2 to March 31.');
    const reopened = await openDocx(bytes);
    try { expect(reopened.grep({ query: 'April 2' }).matches.length).toBeGreaterThan(0); } finally { reopened.close(); }
  } finally { doc.close(); }
});
