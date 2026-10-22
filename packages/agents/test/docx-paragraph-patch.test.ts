import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const FONT = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>';
const PROPERTIES = `<w:rPr>${FONT}<w:noProof/><w:color w:val="auto"/></w:rPr>`;
const DOCUMENT = 'word/document.xml';
const HEADER = 'word/header1.xml';
const BODY_RUNS = ['A0000001', 'A0000002', 'A0000003', 'A0000004', 'A0000005'];
const HEADER_RUNS = ['A0000006'];

function run(id: string, text: string, properties = PROPERTIES): string {
  return `<w:r w:rsidR="${id}">${properties}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function paragraph(id: string, content: string, properties = ''): string {
  return `<w:p w14:paraId="${id}" w:rsidR="01020304">${properties}${content}</w:p>`;
}

async function syntheticPackage() {
  const heading = paragraph('00000001',
    run('A0000001', 'The ochre fox rests.') +
    '<w:hyperlink r:id="link" w:history="1" w:tooltip="Keep &amp; retain">' +
    run('B0000001', ' Linked text.', '<w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr>') +
    '</w:hyperlink>' + run('B0000002', ' Untouched tail.', '<w:rPr><w:b/><w:bCs/></w:rPr>'),
    '<w:pPr><w:pStyle w:val="Heading1"/><w:divId w:val="17"/></w:pPr>');
  const note = paragraph('00000003',
    run('A0000002', 'Please review this.') +
    '<w:r w:rsidR="B0000003"><w:rPr><w:rStyle w:val="FootnoteReference"/>' +
    '<w:vertAlign w:val="superscript"/><w:noProof/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
    run('B0000004', ' Keep this &amp; that.', '<w:rPr><w:i/><w:color w:val="2468AC"/></w:rPr>'),
    '<w:pPr><w:divId w:val="23"/></w:pPr>');
  const mixed = paragraph('00000004',
    run('B0000005', 'Mixed ') +
    run('A0000004', 'bold', '<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>') +
    run('A0000005', 'plain') + run('B0000006', ' suffix.'));
  const cell = paragraph('00000005', run('A0000003', 'Cell amber value.'));
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>' +
    '<w:top w:val="single" w:sz="8" w:space="2" w:color="123456"/>' +
    '<w:bottom w:val="double" w:sz="12" w:space="3" w:color="ABCDEF"/>' +
    '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/>' +
    '<w:gridCol w:w="2000"/></w:tblGrid><w:tr w:rsidR="11223344" w14:paraId="11000001">' +
    '<w:trPr><w:gridAfter w:val="1"/><w:trHeight w:val="360"/></w:trPr>' +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr>` +
    '<w:tr><w:tc>' + paragraph('00000006', run('B0000007', 'Untouched cell.')) +
    '</w:tc><w:tc>' + paragraph('00000007', run('B0000008', 'Other cell.')) +
    '</w:tc></w:tr></w:tbl>';
  const parts = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      '</Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>` +
      `<Relationship Id="header" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="notes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="theme" Type="${R}/theme" Target="theme/theme1.xml"/>` +
      `<Relationship Id="link" Type="${R}/hyperlink" Target="https://example.invalid/link" TargetMode="External"/>` +
      '</Relationships>'],
    ['word/styles.xml', `<w:styles ${NS}><w:docDefaults><w:rPrDefault>${PROPERTIES}` +
      '</w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
      '<w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1">' +
      '<w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/>' +
      '</w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>' +
      '<w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="FootnoteReference">' +
      '<w:name w:val="Footnote Reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' +
      '</w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
      '<w:rPr><w:u w:val="single"/></w:rPr></w:style></w:styles>'],
    ['word/theme/theme1.xml', `<a:theme xmlns:a="${A}" name="Theme"><a:themeElements>` +
      '<a:fontScheme name="Fonts"><a:majorFont><a:latin typeface="Georgia"/>' +
      '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont>' +
      '<a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
      '</a:fontScheme></a:themeElements></a:theme>'],
    [HEADER, `<w:hdr ${NS}>` + paragraph('00000009', run('A0000006', 'Header amber.')) +
      paragraph('0000000A', run('B0000009', 'Header stays.')) + '</w:hdr>'],
    ['word/footnotes.xml', `<w:footnotes ${NS}><w:footnote w:id="1">` +
      paragraph('0000000B', '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
        '<w:footnoteRef/></w:r>' + run('B000000A', 'Footnote stays.')) +
      '</w:footnote></w:footnotes>'],
    ['customXml/preserved.xml', '<payload data="unchanged">Keep &#x61;ll source bytes.</payload>'],
    [DOCUMENT, `<w:document ${NS}><w:body>${heading}` +
      paragraph('00000002', run('B000000B', 'Untouched middle.')) + note + mixed + table +
      paragraph('00000008', run('B000000C', 'Untouched final.')) +
      '<w:sectPr><w:headerReference w:type="default" r:id="header"/>' +
      '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'],
  ]);
  const zip = new JSZip();
  for (const [name, xml] of parts) zip.file(name, xml);
  return { parts, bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) };
}

function runs(xml: string): string[] {
  return xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? [];
}

function hasId(xml: string, id: string): boolean {
  return xml.startsWith(`<w:r w:rsidR="${id}"`);
}

function withoutTouchedRuns(xml: string, ids: string[]): string {
  return xml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g,
    value => ids.some(id => hasId(value, id)) ? '' : value);
}

function properties(xml: string): string {
  return xml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
}

async function replace(doc: DocxAgentDocument, query: string, newText: string) {
  const matches = doc.grep({ query }).matches;
  expect(matches).toHaveLength(1);
  const proposal = doc.propose({ author: 'test', edits: [{ match: matches[0].match, newText }] });
  expect(await doc.accept(proposal.id)).toMatchObject({ status: 'accepted', tracked: false });
}

async function paragraphMarkPackage(mark: string, value = 'abc') {
  const source = await syntheticPackage();
  const properties = `<w:pPr><w:pStyle w:val="Normal"/><w:rPr>${mark}</w:rPr></w:pPr>`;
  const content = value ? `<w:r><w:t>${value}</w:t></w:r>` : '';
  const kept = paragraph('00000002', '<w:r><w:t>keep</w:t></w:r>', properties);
  source.parts.set(DOCUMENT, `<w:document ${NS}><w:body>` +
    paragraph('00000001', content, properties) + kept +
    '<w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr>' +
    '</w:body></w:document>');
  const zip = new JSZip();
  for (const [name, xml] of source.parts) zip.file(name, xml);
  return {
    parts: source.parts,
    kept,
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
  };
}

test('agent export preserves paragraph mark formatting when only text changes', async () => {
  for (const mark of [
    '<w:sz w:val="32"/>',
    '<w:b/><w:sz w:val="32"/><w:rFonts w:ascii="Serif" w:hAnsi="Keep"/>',
  ]) {
    const source = await paragraphMarkPackage(mark);
    const doc = await openDocx(source.bytes);
    try {
      const unchanged = await JSZip.loadAsync(await doc.export());
      expect(await unchanged.file(DOCUMENT)!.async('string')).toBe(source.parts.get(DOCUMENT)!);
      await replace(doc, 'abc', 'aXc');
      const part = doc.session.paragraphSavePlan().splicedParts!
        .find(part => part.part === DOCUMENT)!;
      expect(part.changed).toEqual([0]);
      expect(part.legacyParagraphs ?? []).toEqual([]);
      expect(part.paragraphPatches).toHaveLength(1);
      const saved = await JSZip.loadAsync(await doc.export());
      for (const [name, original] of source.parts) {
        const expected = name === DOCUMENT ? original.replace('>abc<', '>aXc<') : original;
        expect(await saved.file(name)!.async('uint8array')).toEqual(new TextEncoder().encode(expected));
      }
    } finally { doc.close(); }
  }
}, 30_000);

test('agent export counts paragraph mark formatting changes as paragraph replacements', async () => {
  for (const value of ['abc', '']) {
    const source = await paragraphMarkPackage('<w:sz w:val="32"/><w:b/>', value);
    const doc = await openDocx(source.bytes);
    const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
    try {
      const { paraId, properties } = doc.session.paragraphs('body')[0]!;
      const defaults = properties.defaultTextFormatting as Record<string, unknown>;
      doc.session.setParagraphAttrs({
        story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: value.length },
      }, { defaultTextFormatting: { ...defaults, fontSize: 40 } });
      const part = doc.session.paragraphSavePlan().splicedParts!
        .find(part => part.part === DOCUMENT)!;
      expect(part.changed).toEqual([0]);
      expect(part.paragraphPatches ?? []).toEqual([]);
      expect(part.legacyParagraphs).toEqual([expect.objectContaining({
        ordinal: 0, code: 'UNSUPPORTED_KEY', detail: 'key:defaultTextFormatting.fontSize',
      })]);
      const saved = await JSZip.loadAsync(await doc.export());
      const xml = await saved.file(DOCUMENT)!.async('string');
      expect(xml).toContain(source.kept);
      expect(diagnostic.mock.calls.filter(([warning]) =>
        typeof warning === 'string' && warning.startsWith('DOCX_MINIMAL_PATCH replace counts')
      )).toEqual([['DOCX_MINIMAL_PATCH replace counts {"key:defaultTextFormatting.fontSize":1}']]);
    } finally {
      diagnostic.mockRestore();
      doc.close();
    }
  }
}, 30_000);

test.todo('agent export writes a changed paragraph mark font size (deferred: main keeps the source size too)', async () => {
  const source = await paragraphMarkPackage('<w:sz w:val="32"/><w:b/>', 'abc');
  const doc = await openDocx(source.bytes);
  try {
    const { paraId, properties } = doc.session.paragraphs('body')[0]!;
    const defaults = properties.defaultTextFormatting as Record<string, unknown>;
    doc.session.setParagraphAttrs({
      story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 3 },
    }, { defaultTextFormatting: { ...defaults, fontSize: 40 } });
    const saved = await JSZip.loadAsync(await doc.export());
    expect(await saved.file(DOCUMENT)!.async('string')).toContain('<w:sz w:val="40"/>');
  } finally {
    doc.close();
  }
}, 30_000);

test('agent export patches body, note-adjacent, cell and header runs without changing other bytes', async () => {
  const source = await syntheticPackage();
  const doc = await openDocx(source.bytes);
  try {
    await replace(doc, 'ochre', 'violet');
    const review = doc.grep({ query: 'review' });
    expect(review.matches).toHaveLength(1);
    await doc.edit({ version: review.version, edits: [
      { op: 'format_text', match: review.matches[0].match, italic: true },
    ] });
    await replace(doc, 'Cell amber', 'Cell violet');
    await replace(doc, 'boldplain', 'XYZ');
    const header = doc.grep({ query: 'Header amber' }).matches[0];
    expect(header.story).not.toBe('body');
    await replace(doc, 'Header amber', 'Header violet');

    const plan = doc.session.paragraphSavePlan();
    const spliced = plan.splicedParts ?? [];
    expect(spliced.flatMap(part => part.legacyParagraphs ?? [])).toEqual([]);
    expect(spliced.flatMap(part => part.paragraphPatches ?? [])).toHaveLength(5);
    for (const part of spliced) {
      expect(part.paragraphPatches?.map(patch => patch.ordinal)).toEqual(part.changed);
    }

    const saved = await doc.export();
    const zip = await JSZip.loadAsync(saved);
    const savedNames = Object.values(zip.files).filter(file => !file.dir).map(file => file.name);
    expect(savedNames.sort()).toEqual([...source.parts.keys()].sort());
    for (const [name, original] of source.parts) {
      const bytes = await zip.file(name)!.async('uint8array');
      const ids = name === DOCUMENT ? BODY_RUNS : name === HEADER ? HEADER_RUNS : [];
      if (!ids.length) {
        expect(bytes).toEqual(new TextEncoder().encode(original));
        continue;
      }
      const xml = new TextDecoder().decode(bytes);
      expect(withoutTouchedRuns(xml, ids)).toBe(withoutTouchedRuns(original, ids));
      const before = original.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
      const after = xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
      expect(after).toHaveLength(before.length);
      before.forEach((paragraph, index) => {
        expect(withoutTouchedRuns(after[index], ids)).toBe(withoutTouchedRuns(paragraph, ids));
      });
      for (const id of ids) {
        const retained = runs(xml).filter(value => hasId(value, id));
        expect(retained).toHaveLength(id === 'A0000002' ? 3 : 1);
        const originalRun = runs(original).find(value => hasId(value, id))!;
        for (const value of retained) {
          expect(properties(value).replace('<w:i/>', '')).toBe(properties(originalRun));
        }
      }
    }

    const body = await zip.file(DOCUMENT)!.async('string');
    expect(body).toContain(run('A0000001', 'The violet fox rests.'));
    expect(body).toContain(run('A0000003', 'Cell violet value.'));
    expect(body).toContain(run('A0000004', 'XYZ', '<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>'));
    expect(body).toContain(run('A0000005', ''));
    expect(body).toContain(run('A0000002', 'Please '));
    expect(body).toContain(run('A0000002', 'review', PROPERTIES.replace('<w:noProof/>', '<w:i/><w:noProof/>')));
    expect(body).toContain(run('A0000002', ' this.'));
    expect(await zip.file(HEADER)!.async('string')).toContain(run('A0000006', 'Header violet.'));

    const reopened = await openDocx(saved);
    try {
      for (const query of ['The violet fox rests.', 'Cell violet value.', 'Mixed XYZ suffix.', 'Header violet.']) {
        expect(reopened.grep({ query }).matches).toHaveLength(1);
      }
      const italic = reopened.grep({ query: 'review' }).matches[0];
      expect(reopened.read(italic.ref).runs.filter(run => run.formatting.italic &&
        run.start < italic.end && run.end > italic.start).map(run => [run.start, run.end]))
        .toEqual([[italic.start, italic.end]]);
    } finally { reopened.close(); }
  } finally { doc.close(); }
}, 30_000);

test('set_paragraph_style uses counted restyle Replace and preserves untouched paragraphs', async () => {
  const source = await syntheticPackage();
  const styles = `<w:styles ${NS}>` + ['A', 'B'].map((id, index) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>` +
    `<w:rPr><w:rFonts w:ascii="${index ? 'Times New Roman' : 'Arial'}" ` +
    `w:hAnsi="${index ? 'Times New Roman' : 'Arial'}"/></w:rPr></w:style>`).join('') +
    '</w:styles>';
  const direct = '<w:rPr><w:rFonts w:ascii="Courier New"/><w:noProof/></w:rPr>';
  const complete = '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Arial"/><w:noProof/></w:rPr>';
  const keptRun = run('B0000001', ' untouched', complete);
  const changed = paragraph('00000001', run('A0000001', 'Retained', direct) + keptRun,
    '<w:pPr><w:pStyle w:val="A"/><w:divId w:val="17"/></w:pPr>');
  const before = paragraph('00000003', run('B0000003', 'Before every byte.'));
  const kept = paragraph('00000002', run('B0000002', 'Keep every byte.', complete));
  source.parts.set('word/styles.xml', styles);
  source.parts.set(DOCUMENT, `<w:document ${NS}><w:body>${before}${changed}${kept}` +
    '<w:sectPr/></w:body></w:document>');
  const zip = new JSZip();
  for (const [name, xml] of source.parts) zip.file(name, xml);
  const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
  const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
  try {
    const selected = doc.list().items.find(item => item.text === 'Retained untouched')!;
    await doc.edit({ version: doc.list().version, edits: [
      { op: 'set_paragraph_style', paragraph: selected.ref, style: 'B' },
    ] });
    const expected = { ascii: 'Times New Roman', hAnsi: 'Times New Roman' };
    expect(doc.read(selected.ref).runs[0].formatting.fontFamily).toMatchObject(expected);
    const plan = doc.session.paragraphSavePlan().splicedParts!.find(part => part.part === DOCUMENT)!;
    expect(plan.changed).toEqual([1]);
    expect(plan.legacyParagraphs).toEqual([expect.objectContaining({
      ordinal: 1, code: 'UNSUPPORTED_STRUCTURE', detail: expect.stringMatching(/^restyle:/),
    })]);
    expect(plan.paragraphPatches ?? []).toEqual([]);
    const bytes = await doc.export();
    const saved = await JSZip.loadAsync(bytes);
    for (const [name, original] of source.parts) {
      if (name !== DOCUMENT) {
        expect(await saved.file(name)!.async('uint8array')).toEqual(new TextEncoder().encode(original));
      }
    }
    const xml = await saved.file(DOCUMENT)!.async('string');
    expect(xml).toContain('<w:pStyle w:val="B"/>');
    expect(xml).toContain(before);
    expect(xml).toContain(kept);
    expect(diagnostic.mock.calls.filter(([warning]) => typeof warning === 'string' &&
      warning.startsWith('DOCX_MINIMAL_PATCH replace counts'))).toEqual([
        ['DOCX_MINIMAL_PATCH replace counts {"restyle":1}'],
      ]);
    const reopened = await openDocx(bytes);
    try {
      const selected = reopened.list().items.find(item => item.text === 'Retained untouched')!;
      expect(reopened.read(selected.ref).runs[0].formatting.fontFamily).toMatchObject(expected);
    } finally { reopened.close(); }
  } finally {
    diagnostic.mockRestore();
    doc.close();
  }
}, 30_000);
