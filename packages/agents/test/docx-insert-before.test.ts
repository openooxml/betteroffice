import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const opened: DocxAgentDocument[] = [];
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const canonical = (xml: string) => xml.replace(/<([a-zA-Z0-9:]+)((?:\s+[a-zA-Z0-9:]+="[^"]*")+)\s*(\/?)>/g, (_all, tag: string, attrs: string, self: string) => {
  const sorted = [...attrs.matchAll(/([a-zA-Z0-9:]+)="([^"]*)"/g)].sort((a, b) => a[1].localeCompare(b[1])).map(match => ` ${match[1]}="${match[2]}"`).join('');
  return `<${tag}${sorted}${self}>`;
});
async function xmlOf(doc: DocxAgentDocument) {
  return (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
}
const blocks = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const textOf = (paragraph: string) => [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
const pPrOf = (paragraph: string) => paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
const firstRPrOf = (paragraph: string) => paragraph.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
const blockOf = (xml: string, text: string) => blocks(xml).find(paragraph => textOf(paragraph) === text)!;
const item = (doc: DocxAgentDocument, text: string) => doc.list({ limit: 100 }).items.find(entry => entry.text === text)!;
const itemAfterBlock = (doc: DocxAgentDocument, text: string) => doc.list({ limit: 100 }).items.find(entry => entry.text.endsWith(text))!;

let nextId = 0x2000;
function para(pPr: string, text: string, rPr = ''): string {
  nextId++;
  const id = nextId.toString(16).toUpperCase().padStart(8, '0');
  return `<w:p w14:paraId="${id}" w14:textId="7${id.slice(1)}" w:rsidR="00AB0001">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

async function fixtureWith(body: string, styles = '', extra?: (zip: JSZip) => Promise<void>, wholeBody = false): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await editFixture());
  if (styles) zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${styles}</w:styles>`));
  const xml = await zip.file('word/document.xml')!.async('string');
  const headed = xml.replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ');
  zip.file('word/document.xml', wholeBody
    ? headed.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}</w:body>`)
    : headed.replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${body}`));
  if (extra) await extra(zip);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const BODY_STYLES = '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="40"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>';
const BODY_PPR = '<w:pStyle w:val="BodyText"/><w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="4472C4"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:right="360" w:firstLine="360"/><w:jc w:val="both"/><w:outlineLvl w:val="2"/><w:rPr><w:i/><w:sz w:val="18"/></w:rPr>';
const BODY_RPR = '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:bCs w:val="0"/><w:i/><w:iCs/><w:color w:val="C00000"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-GB"/>';

function mainFixture(): Promise<Uint8Array> {
  const body = para('<w:pStyle w:val="Heading1"/>', 'Heading', '<w:b/><w:sz w:val="32"/>')
    + para(BODY_PPR, 'Body text', BODY_RPR)
    + para('', 'Tail');
  return fixtureWith(body, BODY_STYLES);
}

test('insert before copies paragraph and first run formatting with runFormat anchor', async () => {
  const doc = await open(await mainFixture());
  const baseline = blocks(await xmlOf(doc)).map(canonical);
  const receipt = await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Body text').ref, runFormat: 'anchor',
    paragraphs: [{ text: 'Inserted one' }, { text: 'Inserted two' }],
  }] });
  const saved = await xmlOf(doc);
  const now = blocks(saved);
  expect(now.map(textOf)).toEqual(['Heading', 'Inserted one', 'Inserted two', 'Body text', 'Tail']);
  expect([...now.slice(0, 1), ...now.slice(3)].map(canonical)).toEqual(baseline);
  const body = blockOf(saved, 'Body text');
  for (const text of ['Inserted one', 'Inserted two']) {
    const added = blockOf(saved, text);
    expect(canonical(pPrOf(added))).toBe(canonical(pPrOf(body)));
    expect(canonical(firstRPrOf(added))).toBe(canonical(firstRPrOf(body)));
  }
  expect(receipt.results[0].refs?.map(ref => doc.read(ref!).text)).toEqual(['Inserted one', 'Inserted two']);
  expect(doc.list({ limit: 100 }).items.slice(1, 3).map(entry => entry.ref)).toEqual(receipt.results[0].refs as string[]);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_paragraph_style', paragraph: receipt.results[0].refs![0]!, style: 'Quote' }] });
  expect(doc.read(receipt.results[0].refs![0]!).style).toBe('Quote');
});

test('insert before without runFormat leaves runs unformatted and style replaces only pStyle', async () => {
  const doc = await open(await mainFixture());
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Body text').ref, paragraphs: [{ text: 'Plain inserted' }],
  }] });
  const saved = await xmlOf(doc);
  const added = blockOf(saved, 'Plain inserted');
  expect(canonical(pPrOf(added))).toBe(canonical(pPrOf(blockOf(saved, 'Body text'))));
  expect(firstRPrOf(added)).not.toContain('<w:i');
  expect(firstRPrOf(added)).not.toContain('Georgia');
  expect(firstRPrOf(added)).not.toContain('C00000');
  expect(doc.read(item(doc, 'Plain inserted').ref).runs[0].formatting.italic).toBeUndefined();

  const styled = await open(await mainFixture());
  await styled.edit({ version: styled.list().version, edits: [{
    op: 'insert_paragraphs', before: item(styled, 'Body text').ref, paragraphs: [{ text: 'Styled inserted', style: 'Heading2' }],
  }] });
  const styledSaved = await xmlOf(styled);
  const styledAdded = blockOf(styledSaved, 'Styled inserted');
  const expected = canonical(pPrOf(blockOf(styledSaved, 'Body text'))).replace('BodyText', 'Heading2');
  expect(canonical(pPrOf(styledAdded))).toBe(expected);
  expect(styled.read(item(styled, 'Styled inserted').ref).style).toBe('Heading2');
});

test('insert before drops page break and frame and never copies section properties', async () => {
  const framedBody = para('', 'Prev')
    + para('<w:pStyle w:val="BodyText"/><w:pageBreakBefore/><w:framePr w:w="2880" w:h="1440" w:hRule="atLeast" w:x="720" w:y="720" w:hAnchor="page" w:vAnchor="page" w:wrap="around"/><w:spacing w:after="120"/>', 'Framed')
    + para('', 'Tail');
  const doc = await open(await fixtureWith(framedBody, BODY_STYLES));
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Framed').ref, paragraphs: [{ text: 'Before framed' }],
  }] });
  const saved = await xmlOf(doc);
  expect(pPrOf(blockOf(saved, 'Before framed'))).not.toContain('pageBreakBefore');
  expect(pPrOf(blockOf(saved, 'Before framed'))).not.toContain('framePr');
  expect(pPrOf(blockOf(saved, 'Framed'))).toContain('<w:pageBreakBefore/>');
  expect(pPrOf(blockOf(saved, 'Framed'))).toContain('<w:framePr');

  const sectionBody = para('<w:pStyle w:val="BodyText"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720"/></w:sectPr>', 'Section end')
    + para('', 'Next section')
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720"/></w:sectPr>';
  const sectioned = await open(await fixtureWith(sectionBody, BODY_STYLES, undefined, true));
  const sectionsBefore = sectioned.list().sections.total;
  const sectCountBefore = (await xmlOf(sectioned)).match(/<w:sectPr/g)?.length;
  await sectioned.edit({ version: sectioned.list().version, edits: [{
    op: 'insert_paragraphs', before: item(sectioned, 'Section end').ref, paragraphs: [{ text: 'Before section end' }],
  }] });
  const sectionSaved = await xmlOf(sectioned);
  expect(pPrOf(blockOf(sectionSaved, 'Before section end'))).not.toContain('sectPr');
  expect(sectionSaved.match(/<w:sectPr/g)?.length).toBe(sectCountBefore);
  expect(sectioned.list().sections.total).toBe(sectionsBefore);
});

test('insert before a numbered list item carries its numbering', async () => {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const numbering = `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const body = para('<w:pStyle w:val="Heading1"/>', 'Heading')
    + para('<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr><w:spacing w:after="40"/>', 'Numbered')
    + para('', 'Tail');
  const styles = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>';
  const doc = await open(await fixtureWith(body, styles, async zip => {
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
    zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
    zip.file('word/numbering.xml', numbering);
  }));
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Numbered').ref, paragraphs: [{ text: 'Before numbered' }],
  }] });
  const saved = await xmlOf(doc);
  expect(pPrOf(blockOf(saved, 'Before numbered'))).toContain('<w:numPr>');
  expect(pPrOf(blockOf(saved, 'Before numbered'))).toContain('<w:ilvl w:val="1"/>');
  expect(pPrOf(blockOf(saved, 'Before numbered'))).toContain('<w:numId w:val="2"/>');
  expect(item(doc, 'Before numbered').list).toEqual({ level: 1 });
});

test('insert before the first cell paragraph and the first body paragraph stays in place', async () => {
  const cellFirst = para('', 'Cell first');
  const cellSecond = para('', 'Cell second');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cellFirst}${cellSecond}</w:tc></w:tr></w:tbl>`;
  const doc = await open(await fixtureWith(para('', 'Prev') + table + para('', 'After')));
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Cell first').ref, paragraphs: [{ text: 'Inserted in cell' }],
  }] });
  const inserted = item(doc, 'Inserted in cell');
  expect(inserted.table).toEqual(item(doc, 'Cell first').table);
  expect(blocks(await xmlOf(doc)).map(textOf)).toEqual(['Prev', 'Inserted in cell', 'Cell first', 'Cell second', 'After']);

  const top = await open(await mainFixture());
  await top.edit({ version: top.list().version, edits: [{
    op: 'insert_paragraphs', before: item(top, 'Heading').ref, paragraphs: [{ text: 'At the top' }],
  }] });
  expect(blocks(await xmlOf(top)).map(textOf)).toEqual(['At the top', 'Heading', 'Body text', 'Tail']);
  expect(top.read(item(top, 'At the top').ref).style).toBe('Heading1');
});

test('insert before validation refuses bad anchors and list continue', async () => {
  const doc = await open(await mainFixture());
  const body = item(doc, 'Body text').ref;
  const heading = item(doc, 'Heading').ref;
  const version = () => doc.list().version;
  await expect(doc.edit({ version: version(), edits: [{ op: 'insert_paragraphs', before: body, after: heading, paragraphs: [{ text: 'X' }] }] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  await expect(doc.edit({ version: version(), edits: [{ op: 'insert_paragraphs', paragraphs: [{ text: 'X' }] }] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  await expect(doc.edit({ version: version(), edits: [{ op: 'insert_paragraphs', before: body, list: 'continue', paragraphs: [{ text: 'X' }] }] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  await expect(doc.edit({ version: version(), edits: [{ op: 'insert_paragraphs', before: 'p9999', paragraphs: [{ text: 'X' }] }] })).rejects.toMatchObject({ code: 'UNKNOWN_REF' });
});

test('insert before keeps style inherited run formatting for read and later edits', async () => {
  const doc = await open(await mainFixture());
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Heading').ref, paragraphs: [{ text: 'At the top' }],
  }] });
  const inserted = doc.read(item(doc, 'At the top').ref);
  expect(inserted.runs).toHaveLength(1);
  expect(inserted.runs[0].formatting.bold).toBe(true);
  const match = doc.grep({ query: 'At the top' });
  await doc.edit({ version: match.version, edits: [{ op: 'format_text', match: match.matches[0].match, italic: true }] });
  const saved = await xmlOf(doc);
  const added = blockOf(saved, 'At the top');
  expect(firstRPrOf(added)).toContain('<w:i/>');
  expect(firstRPrOf(added)).not.toMatch(/<w:b w:val="(?:0|false|off)"\/>/);
  expect(doc.read(item(doc, 'At the top').ref).runs[0].formatting.bold).toBe(true);
});

test('insert before a paragraph after a table formats the whole new text', async () => {
  const build = () => {
    const cell = para('', 'Cell text');
    const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;
    return fixtureWith(para('', 'Prev') + table + para(BODY_PPR, 'After table', BODY_RPR) + para('', 'Tail'), BODY_STYLES);
  };
  const plain = await open(await build());
  await plain.edit({ version: plain.list().version, edits: [{
    op: 'insert_paragraphs', before: itemAfterBlock(plain, 'After table').ref, paragraphs: [{ text: 'Inserted plain' }],
  }] });
  expect(plain.read(itemAfterBlock(plain, 'Inserted plain').ref).runs).toHaveLength(1);
  expect(blocks(await xmlOf(plain)).map(textOf)).toContain('Inserted plain');
  expect(firstRPrOf(blockOf(await xmlOf(plain), 'Inserted plain'))).not.toContain('C00000');

  const anchored = await open(await build());
  await anchored.edit({ version: anchored.list().version, edits: [{
    op: 'insert_paragraphs', before: itemAfterBlock(anchored, 'After table').ref, runFormat: 'anchor', paragraphs: [{ text: 'Inserted anchored' }],
  }] });
  const saved = await xmlOf(anchored);
  expect(anchored.read(itemAfterBlock(anchored, 'Inserted anchored').ref).runs).toHaveLength(1);
  expect(canonical(firstRPrOf(blockOf(saved, 'Inserted anchored')))).toBe(canonical(firstRPrOf(blockOf(saved, 'After table'))));
});

test('explicit style on insert before resolves the new style instead of copying the anchor style', async () => {
  const doc = await open(await mainFixture());
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Heading').ref,
    paragraphs: [{ text: 'Lead-in', style: 'Normal' }, { text: 'Sub', style: 'Heading2' }],
  }] });
  const items = doc.list({ limit: 100 }).items;
  expect(items.find(entry => entry.text === 'Lead-in')?.headingLevel).toBeNull();
  expect(items.find(entry => entry.text === 'Sub')?.headingLevel).toBe(2);

  const styles = '<w:style w:type="paragraph" w:styleId="SpacedA"><w:name w:val="Spaced A"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="111"/><w:jc w:val="center"/></w:pPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="SpacedB"><w:name w:val="Spaced B"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="222"/></w:pPr></w:style>';
  const leaked = await open(await fixtureWith(para('<w:pStyle w:val="SpacedA"/>', 'Anchor A') + para('', 'Tail'), styles));
  await leaked.edit({ version: leaked.list().version, edits: [{
    op: 'insert_paragraphs', before: item(leaked, 'Anchor A').ref, paragraphs: [{ text: 'With B', style: 'SpacedB' }],
  }] });
  const added = blockOf(await xmlOf(leaked), 'With B');
  expect(canonical(pPrOf(added))).toBe('<w:pPr><w:pStyle w:val="SpacedB"/></w:pPr>');
});

async function numberedStyleFixture(anchorText = 'Style one', noDefault = false): Promise<Uint8Array> {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const level = `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="ListNumber"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;
  const numbering = `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="2">${level}</w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
  const styles = '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr><w:spacing w:after="60"/></w:pPr></w:style>';
  const body = para('<w:pStyle w:val="Heading1"/>', 'Heading') + para('<w:pStyle w:val="ListNumber"/>', anchorText) + para('', 'Plain') + para('', 'Tail');
  return fixtureWith(body, styles, async zip => {
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
    zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
    zip.file('word/numbering.xml', numbering);
    if (noDefault) zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace(' w:default="1"', ''));
  });
}

test('insert before a style numbered paragraph keeps the list only without an explicit style', async () => {
  const doc = await open(await numberedStyleFixture());
  await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: item(doc, 'Style one').ref, paragraphs: [{ text: 'Before style list' }],
  }] });
  expect(item(doc, 'Before style list').list).toEqual({ level: 0 });
  const added = blockOf(await xmlOf(doc), 'Before style list');
  expect(pPrOf(added)).toContain('<w:pStyle w:val="ListNumber"/>');
  expect(pPrOf(added)).not.toContain('<w:numPr>');

  const styled = await open(await numberedStyleFixture());
  await styled.edit({ version: styled.list().version, edits: [{
    op: 'insert_paragraphs', before: item(styled, 'Style one').ref, paragraphs: [{ text: 'Not a list', style: 'Normal' }],
  }] });
  expect(item(styled, 'Not a list').list).toBeUndefined();
  expect(pPrOf(blockOf(await xmlOf(styled), 'Not a list'))).not.toContain('<w:numPr>');
});

test('explicit numbered style is refused like after and a missing default style does not break before', async () => {
  const doc = await open(await numberedStyleFixture());
  const plain = item(doc, 'Plain').ref;
  const codeOf = async (edit: Parameters<DocxAgentDocument['edit']>[0]['edits'][number]) => {
    try { await doc.edit({ version: doc.list().version, edits: [edit] }); return 'ok'; }
    catch (error) { return (error as { code: string }).code; }
  };
  const afterCode = await codeOf({ op: 'insert_paragraphs', after: plain, paragraphs: [{ text: 'X', style: 'ListNumber' }] });
  const beforeCode = await codeOf({ op: 'insert_paragraphs', before: plain, paragraphs: [{ text: 'X', style: 'ListNumber' }] });
  expect(afterCode).not.toBe('ok');
  expect(beforeCode).toBe(afterCode);

  const noDefault = await open(await numberedStyleFixture('Style one', true));
  await noDefault.edit({ version: noDefault.list().version, edits: [{
    op: 'insert_paragraphs', before: item(noDefault, 'Heading').ref, paragraphs: [{ text: 'No default top' }],
  }] });
  expect(blocks(await xmlOf(noDefault)).map(textOf)[0]).toBe('No default top');
});

test('insert after keeps its existing plain behaviour', async () => {
  const doc = await open(await mainFixture());
  const receipt = await doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', after: item(doc, 'Heading').ref, paragraphs: [{ text: 'Added after' }],
  }] });
  expect(blocks(await xmlOf(doc)).map(textOf)).toEqual(['Heading', 'Added after', 'Body text', 'Tail']);
  expect(receipt.results[0].refs?.map(ref => doc.read(ref!).text)).toEqual(['Added after']);
  expect(pPrOf(blockOf(await xmlOf(doc), 'Added after'))).not.toContain('BodyText');
});
