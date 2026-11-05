import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RPR = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>' +
  '<w:b/><w:bCs/><w:color w:val="0066AA"/><w:sz w:val="22"/><w:szCs w:val="22"/>' +
  '<w:lang w:val="de-DE" w:eastAsia="de-DE" w:bidi="ar-SA"/>';
const RPR_ALT = RPR.replace('<w:b/><w:bCs/>', '<w:i/><w:iCs/>').replace('0066AA', 'C00000');
const run = (id: string, text: string, properties = RPR) =>
  `<w:r w:rsidR="${id}" w:rsidRPr="${id}"><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (content: string) => `<w:p>${content}</w:p>`;
const kept = '<w:p w:rsidR="98765432"><w:r><w:t>Keep &#x61;ll bytes.</w:t></w:r></w:p>';
const paragraphsOf = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const runCount = (xml: string) => xml.match(/<w:r\b/g)?.length ?? 0;
const textOf = (xml: string) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');

async function fixture(content: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>` +
    `<Relationship Id="comments" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
  zip.file('word/comments.xml', `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="test">` +
    '<w:p><w:r><w:t>Review</w:t></w:r></w:p></w:comment></w:comments>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${kept}` +
    paragraph(content) + kept + '<w:sectPr/></w:body></w:document>');
  return zip.generateAsync({ type: 'uint8array' });
}

async function replace(doc: DocxAgentDocument, query: string, newText: string, tracked = false) {
  const matches = doc.grep({ query }).matches;
  expect(matches).toHaveLength(1);
  const proposal = doc.propose({ author: 'test', edits: [{ match: matches[0]!.match, newText }] });
  expect(await doc.accept(proposal.id, tracked ? { tracked, date: '2026-01-01T00:00:00Z' } : { tracked }))
    .toMatchObject({ status: 'accepted', tracked });
}

function forceLegacy(doc: DocxAgentDocument) {
  const { paraId } = doc.session.paragraphs('body')[1]!;
  doc.session.setParagraphAttrs({
    story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 },
  }, { defaultTextFormatting: { fontSize: 40 } });
  const parts = doc.session.paragraphSavePlan().splicedParts ?? [];
  const part = parts.find(part => part.part === 'word/document.xml');
  if (part) expect(part.legacyParagraphs?.some(entry => entry.ordinal === 1)).toBe(true);
  else expect(parts).toHaveLength(0);
}

async function savedParagraph(doc: DocxAgentDocument, preserveBytes = true) {
  const zip = await JSZip.loadAsync(await doc.export());
  const xml = await zip.file('word/document.xml')!.async('string');
  if (preserveBytes) expect(xml.split(kept)).toHaveLength(3);
  return paragraphsOf(xml)[1]!;
}

const shapes = [
  { name: 'inside one run', source: run('000000A1', 'Alpha bravo charlie delta.'), query: 'bravo',
    replacement: 'BRAVO', expected: 'Alpha BRAVO charlie delta.', runs: 1 },
  { name: 'ends mid-run', source: run('000000B1', 'Echo foxtrot golf hotel.'), query: 'Echo fox',
    replacement: 'ECHO FOX', expected: 'ECHO FOXtrot golf hotel.', runs: 1 },
  { name: 'starts mid-run', source: run('000000C1', 'India juliet kilo lima.'), query: 'kilo lima.',
    replacement: 'KILO LIMA!', expected: 'India juliet KILO LIMA!', runs: 1 },
  { name: 'whole run', source: run('000000D1', 'Quebec ') + run('000000D2', 'romeo') + run('000000D3', ' sierra.'),
    query: 'romeo', replacement: 'ROMEO', expected: 'Quebec ROMEO sierra.', runs: 3 },
  { name: 'crosses identical properties', source: run('000000E1', 'Tango uniform ') + run('000000E2', 'victor whiskey.'),
    query: 'uniform victor', replacement: 'UNIFORM VICTOR', expected: 'Tango UNIFORM VICTOR whiskey.', runs: 1 },
  { name: 'crosses different properties', source: run('000000F1', 'Xray yankee ') + run('000000F2', 'zulu alpha.', RPR_ALT),
    query: 'yankee zulu', replacement: 'YANKEE ZULU', expected: 'Xray YANKEE ZULU alpha.', runs: 3 },
  { name: 'untouched source split', source: run('00000061', 'Guard first ') + run('00000062', 'second part.'),
    query: 'first', replacement: 'FIRST', expected: 'Guard FIRST second part.', runs: 2 },
];

for (const legacy of [false, true]) {
  for (const properties of [
    { name: 'language', first: '<w:lang w:val="en-US"/>', second: '<w:lang w:val="de-DE"/>' },
    { name: 'unmodeled properties', first: '', second: '<w:noProof/>' },
  ]) {
    test(`accepted replacements retain ${properties.name} after clearing bold${legacy ? ' through legacy save' : ''}`, async () => {
      const doc = await openDocx(await fixture(
        run('00000091', 'abc ', properties.first) + run('00000092', 'def', `<w:b/>${properties.second}`)
      ));
      try {
        const [hit] = doc.grep({ query: 'abc def' }).matches;
        const source = doc.read(hit!.ref);
        const proposal = doc.propose({ author: 'test', edits: [
          { ...source, start: 0, oldText: 'a', newText: 'A' },
          { ...source, start: 4, oldText: 'd', newText: 'D' },
        ] });
        expect(await doc.accept(proposal.id, { tracked: false })).toMatchObject({ status: 'accepted' });
        const { paraId } = doc.session.paragraphs('body')[1]!;
        doc.session.formatRange({
          story: 'body', start: { paraId, offset: 4 }, end: { paraId, offset: 7 },
        }, { bold: null });
        if (legacy) forceLegacy(doc);
        else {
          const part = doc.session.paragraphSavePlan().splicedParts?.find(part => part.part === 'word/document.xml');
          expect(part?.paragraphPatches?.some(patch => patch.ordinal === 1)).toBe(true);
        }
        const saved = await savedParagraph(doc);
        const runs = saved.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? [];
        expect(runs.map(textOf), saved).toEqual(legacy && properties.name === 'unmodeled properties'
          ? ['Abc Def'] : ['Abc ', 'Def']);
        expect(saved).not.toMatch(/<w:b(?:\s|\/|>)/);
        if (properties.name === 'language') {
          expect(runs[0]).toContain(properties.first);
          expect(runs[1]).toContain(properties.second);
          expect(runs[0]).not.toContain('de-DE');
          expect(runs[1]).not.toContain('en-US');
        } else if (legacy) {
          expect(saved).not.toContain('<w:noProof');
        } else {
          expect(runs[1]).toContain(properties.second);
          expect(runs[0]).not.toContain('<w:noProof');
        }
      } finally { doc.close(); }
    }, 30_000);
  }
}

for (const legacy of [false, true]) {
  for (const shape of shapes) {
    test(`accepted replacement ${shape.name}${legacy ? ' through legacy save' : ''}`, async () => {
      const doc = await openDocx(await fixture(shape.source));
      try {
        await replace(doc, shape.query, shape.replacement);
        if (legacy) forceLegacy(doc);
        const saved = await savedParagraph(doc);
        expect(textOf(saved)).toBe(shape.expected);
        expect(runCount(saved), saved).toBe(shape.runs);
        if (shape.name === 'crosses identical properties' && !legacy) {
          expect(saved).toContain('w:rsidR="000000E1" w:rsidRPr="000000E1"');
          expect(saved).toContain(`<w:rPr>${RPR}</w:rPr>`);
          expect(saved).not.toContain('000000E2');
        }
        if (shape.name === 'crosses different properties') {
          const runs = saved.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)!;
          expect(runs.map(textOf)).toEqual(['Xray YANKEE ', 'ZULU', ' alpha.']);
          expect(runs[0]).toContain('0066AA');
          expect(runs[0]).toContain('<w:bCs/>');
          expect(runs[1]).toContain('0066AA');
          expect(runs[1]).toContain('<w:iCs/>');
          expect(runs[1]).not.toContain('<w:bCs');
          expect(runs[2]).toContain('<w:i/>');
          expect(runs[2]).toContain('C00000');
        }
        if (shape.name === 'untouched source split' && !legacy) {
          expect(saved).toContain(run('00000062', 'second part.'));
        }
      } finally { doc.close(); }
    }, 30_000);
  }
}

const barriers = [
  { name: 'proofErr', before: '', middle: '<w:proofErr w:type="spellStart"/>', after: '', marker: '<w:proofErr' },
  { name: 'bookmarkStart', before: '', middle: '<w:bookmarkStart w:id="1" w:name="kept"/>',
    after: '<w:bookmarkEnd w:id="1"/>', marker: '<w:bookmarkStart' },
  { name: 'bookmarkEnd', before: '<w:bookmarkStart w:id="1" w:name="kept"/>', middle: '<w:bookmarkEnd w:id="1"/>',
    after: '', marker: '<w:bookmarkEnd' },
  { name: 'commentRangeStart', before: '', middle: '<w:commentRangeStart w:id="0"/>',
    after: '<w:commentRangeEnd w:id="0"/>', marker: '<w:commentRangeStart' },
  { name: 'commentRangeEnd', before: '<w:commentRangeStart w:id="0"/>', middle: '<w:commentRangeEnd w:id="0"/>',
    after: '', marker: '<w:commentRangeEnd' },
  { name: 'commentReference', before: '', middle: '<w:r><w:commentReference w:id="0"/></w:r>',
    after: '', marker: '<w:commentReference' },
  { name: 'fldChar', before: '', middle: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', after: '', marker: '<w:fldChar' },
  { name: 'instrText', before: '', middle: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>', after: '', marker: '<w:instrText' },
  { name: 'fldSimple', before: '', middle: '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>',
    after: '', marker: '<w:fldSimple' },
  { name: 'ins', before: '', middle: '<w:ins w:id="8" w:author="test" w:date="2026-01-01T00:00:00Z">' +
    run('00000088', 'revision ') + '</w:ins>', after: '', marker: '<w:ins' },
  { name: 'del', before: '', middle: '<w:del w:id="9" w:author="test" w:date="2026-01-01T00:00:00Z">' +
    '<w:r><w:delText>revision </w:delText></w:r></w:del>', after: '', marker: '<w:del' },
  { name: 'hyperlink scope', before: '', middle: '<w:hyperlink w:anchor="kept">' + run('00000089', 'linked ') +
    '</w:hyperlink>', after: '', marker: '<w:hyperlink' },
];

for (const legacy of [false, true]) {
  for (const barrier of barriers) {
    const name = legacy && barrier.name === 'proofErr'
      ? 'accepted replacements retain canonical proofErr behavior'
      : `accepted replacements retain ${barrier.name} boundaries${legacy ? ' through legacy save' : ''}`;
    test(name, async () => {
      const doc = await openDocx(await fixture(barrier.before + run('00000011', 'Left alpha ') +
        barrier.middle + run('00000012', 'beta right.') + barrier.after));
      try {
        await replace(doc, 'alpha', 'ALPHA');
        await replace(doc, 'beta', 'BETA');
        if (legacy) forceLegacy(doc);
        const saved = await savedParagraph(doc, barrier.name !== 'fldChar' && barrier.name !== 'del');
        if (legacy && barrier.name === 'proofErr') {
          expect(saved).not.toContain(barrier.marker);
          expect(textOf(saved)).toBe('Left ALPHA BETA right.');
          expect(runCount(saved)).toBe(1);
          return;
        }
        expect(saved).toContain(barrier.marker);
        const runs = saved.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? [];
        expect(runs.some(value => textOf(value).includes('ALPHA') && textOf(value).includes('BETA'))).toBe(false);
        expect(textOf(saved)).toContain('Left ALPHA ');
        expect(textOf(saved)).toContain('BETA right.');
        expect(saved.indexOf(barrier.marker)).toBeGreaterThan(saved.indexOf('ALPHA'));
        expect(saved.indexOf(barrier.marker)).toBeLessThan(saved.indexOf('BETA'));
      } finally { doc.close(); }
    }, 30_000);
  }
}

for (const legacy of [false, true]) {
  test(`tracked accept keeps insertion and deletion runs${legacy ? ' through legacy save' : ''}`, async () => {
    const doc = await openDocx(await fixture(run('00000081', 'Tracked one ') + run('00000082', 'two end.')));
    try {
      await replace(doc, 'one two', 'ONE TWO', true);
      if (legacy) forceLegacy(doc);
      const saved = await savedParagraph(doc);
      expect(saved).toMatch(/<w:ins\b/);
      expect(saved).toMatch(/<w:del\b/);
      expect(runCount(saved)).toBeGreaterThan(1);
      expect(textOf(saved)).toContain('ONE TWO');
    } finally { doc.close(); }
  }, 30_000);
}
