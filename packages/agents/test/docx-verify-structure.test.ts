import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { openDocx, type DocxAgentDocument } from '../src';
import { compareDocxStructure, docxModelCounts, normalizeInstruction, type DocxModelCounts, type DocxPartCounts } from '../src/docx-structure';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { createOfficeMcpServer } from '../src/mcp';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), c => c.charCodeAt(0));

const drawing = (id: string) => `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const BODY = [
  '<w:p><w:r><w:t>Introduction text that an agent translates.</w:t></w:r></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r><w:r><w:instrText xml:space="preserve">_Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Table 1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  `<w:p><w:bookmarkStart w:id="0" w:name="_Ref1"/><w:r><w:t>Picture paragraph</w:t></w:r>${drawing('rIdImage')}<w:bookmarkEnd w:id="0"/></w:p>`,
  '<w:p><w:r><w:t>Removable paragraph with a field </w:t></w:r><w:fldSimple w:instr=" DATE "><w:r><w:t>2026</w:t></w:r></w:fldSimple></w:p>',
  '<w:p><w:r><w:t>Closing sentence stays as it is.</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>Spare paragraph the agent deletes.</w:t></w:r></w:p>',
];

async function source(body = BODY.join('')): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdImage" Type="${R}/image" Target="media/image1.png"/></Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/></w:style></w:styles>`);
  zip.file('word/media/image1.png', PNG);
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function mutate(bytes: Uint8Array, change: (zip: JSZip) => Promise<void>) {
  const zip = await JSZip.loadAsync(bytes);
  await change(zip);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
const rewrite = (part: string, edit: (xml: string) => string) => async (zip: JSZip) => {
  zip.file(part, edit(await zip.file(part)!.async('string')));
};

const part = (overrides: Partial<DocxPartCounts> = {}): DocxPartCounts =>
  ({ paragraphs: 0, tables: 0, rows: 0, cells: 0, drawings: 0, notes: 0, controls: 0, fields: {}, ...overrides });
const counts = (fields: Record<string, number> = { PAGE: 1, 'REF _Ref1 \\h': 1, DATE: 1 }, overrides: Partial<DocxPartCounts> = {}, parts: Record<string, DocxPartCounts> = {}): DocxModelCounts =>
  ({ comments: 0, parts: { 'word/document.xml': part({ paragraphs: 7, drawings: 1, fields, ...overrides }), ...parts } });

const opened: DocxAgentDocument[] = [];
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}

test('an unchanged package compares clean', async () => {
  const bytes = await source();
  const report = await compareDocxStructure(bytes, bytes, counts(), counts());
  expect(report.summary).toBe('Structure intact: every drop is explained by the edits.');
  expect(report).toMatchObject({ ok: true, lost: [], unattributed: [], fields: { source: 3, saved: 3 }, drawings: { source: 1, saved: 1 }, relationships: { source: 3, saved: 3 } });
});

test('lost simple and complex field instructions are reported', async () => {
  const bytes = await source();
  const simple = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>', '<w:r><w:t>1</w:t></w:r>')));
  const report = await compareDocxStructure(bytes, simple, counts(), counts());
  expect(report.ok).toBe(false);
  expect(report.summary).toContain('the edits do not explain');
  expect(report.lost).toContainEqual(expect.objectContaining({ kind: 'field', name: 'PAGE', source: 1, saved: 0 }));
  const changed = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('_Ref1 \\h ', '_Ref2 \\h ')));
  const complex = await compareDocxStructure(bytes, changed, counts(), counts());
  expect(complex.ok).toBe(false);
  expect(complex.lost).toContainEqual(expect.objectContaining({ kind: 'field', name: 'REF _Ref1 \\h', source: 1, saved: 0 }));
});

test('a field the edits removed is explained by the model', async () => {
  const bytes = await source();
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(BODY[4], '')));
  const report = await compareDocxStructure(bytes, saved, counts(), counts({ PAGE: 1, 'REF _Ref1 \\h': 1 }, { paragraphs: 6 }));
  expect(report.ok).toBe(true);
  expect(report.lost).toEqual([]);
  expect(report.fields).toEqual({ source: 3, saved: 2 });
});

test('lost drawings, relationships and dangling references are reported', async () => {
  const bytes = await source();
  const noDrawing = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(drawing('rIdImage'), '')));
  const drawings = await compareDocxStructure(bytes, noDrawing, counts(), counts());
  expect(drawings.ok).toBe(false);
  expect(drawings.lost).toContainEqual(expect.objectContaining({ kind: 'element', name: 'w:drawing|w:pict|w:object', source: 1, saved: 0, explained: 0 }));
  expect((await compareDocxStructure(bytes, noDrawing, counts(), counts(undefined, { drawings: 0 }))).lost).toEqual([]);

  const noRelationship = await mutate(bytes, rewrite('word/_rels/document.xml.rels', xml => xml.replace(/<Relationship Id="rIdImage"[^>]*\/>/, '')));
  const relationships = await compareDocxStructure(bytes, noRelationship, counts(), counts());
  expect(relationships.ok).toBe(false);
  expect(relationships.lost).toContainEqual(expect.objectContaining({ kind: 'relationship', part: 'word/document.xml' }));
  expect(relationships.lost).toContainEqual(expect.objectContaining({ kind: 'reference', part: 'word/document.xml', name: 'rIdImage' }));

  const noMedia = await mutate(bytes, async zip => { zip.remove('word/media/image1.png'); });
  const media = await compareDocxStructure(bytes, noMedia, counts(), counts());
  expect(media.lost).toContainEqual(expect.objectContaining({ kind: 'part', name: 'word/media/image1.png' }));
  expect(media.lost).toContainEqual(expect.objectContaining({ kind: 'relationship', name: 'rIdImage target media/image1.png missing' }));
});

test('protected elements lost without a structural edit are losses, after one they are listed for review', async () => {
  const bytes = await source();
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('<w:bookmarkStart w:id="0" w:name="_Ref1"/>', '').replace('<w:bookmarkEnd w:id="0"/>', '')));
  const plain = await compareDocxStructure(bytes, saved, counts(), counts());
  expect(plain.ok).toBe(false);
  expect(plain.lost).toContainEqual(expect.objectContaining({ part: 'word/document.xml', kind: 'element', name: 'w:bookmarkStart', source: 1, saved: 0 }));
  const structural = await compareDocxStructure(bytes, saved, counts(), counts(undefined, { paragraphs: 6 }));
  expect(structural.ok).toBe(true);
  expect(structural.unattributed).toContainEqual(expect.objectContaining({ name: 'w:bookmarkStart' }));
});

test('namespace prefixes do not matter, other element drops are informational', async () => {
  const bytes = await source();
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:ww="${W}"`).replace('<w:p><w:r><w:t>Closing sentence stays as it is.</w:t></w:r></w:p>', '<ww:p><ww:r><ww:t>Closing sentence stays as it is.</ww:t></ww:r></ww:p>').replace('<w:r><w:t xml:space="preserve">See </w:t></w:r>', '')));
  const report = await compareDocxStructure(bytes, saved, counts(), counts());
  expect(report.lost).toEqual([]);
  expect(report.counts).toContainEqual(expect.objectContaining({ part: 'word/document.xml', element: 'w:r' }));
});

test('an addition in one part does not hide a loss in another', async () => {
  const bytes = await mutate(await source(), async zip => {
    zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header one</w:t></w:r></w:p><w:p><w:r><w:t>Header two</w:t></w:r></w:p></w:hdr>`);
  });
  const saved = await mutate(bytes, async zip => {
    zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header one</w:t></w:r></w:p></w:hdr>`);
    await rewrite('word/document.xml', xml => xml.replace('<w:sectPr>', '<w:p><w:r><w:t>Inserted</w:t></w:r></w:p><w:sectPr>'))(zip);
  });
  const header = { 'word/header1.xml': part({ paragraphs: 2 }) };
  const report = await compareDocxStructure(bytes, saved, counts(undefined, {}, header), counts(undefined, { paragraphs: 8 }, header));
  expect(report.ok).toBe(false);
  expect(report.lost).toContainEqual(expect.objectContaining({ part: 'word/header1.xml', kind: 'element', name: 'w:p', source: 2, saved: 1, explained: 0 }));
  const kept = await mutate(bytes, async zip => {
    zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header one</w:t></w:r></w:p></w:hdr>`);
  });
  const deletion = await compareDocxStructure(bytes, kept, counts(undefined, {}, header), counts(undefined, { paragraphs: 6 }, header));
  expect(deletion.lost).toContainEqual(expect.objectContaining({ part: 'word/header1.xml', name: 'w:p', source: 2, saved: 1 }));
});

test('a field lost next to an unrelated deletion is still a loss', async () => {
  const bytes = await source();
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(BODY[6], '').replace('<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>', '<w:r><w:t>1</w:t></w:r>')));
  const report = await compareDocxStructure(bytes, saved, counts(), counts(undefined, { paragraphs: 6 }));
  expect(report.ok).toBe(false);
  expect(report.lost).toContainEqual(expect.objectContaining({ kind: 'field', name: 'PAGE', source: 1, saved: 0, explained: 0 }));
});

test('content nested in a removed row or merged cell is listed for review, not failed', async () => {
  const cell = (inner: string) => `<w:tc><w:tcPr/>${inner}</w:tc>`;
  const row = (inner: string) => `<w:tr>${cell(inner)}</w:tr>`;
  const textBox = '<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>';
  const nested = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  const table = (rows: string) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>`;
  const kept = row('<w:p><w:r><w:t>Kept</w:t></w:r></w:p>');
  const bytes = await source(table(row(textBox + nested) + kept));
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(row(textBox + nested), '')));
  const shape = { tables: 1, drawings: 1, fields: {} };
  const report = await compareDocxStructure(bytes, saved, counts({}, { ...shape, paragraphs: 3, rows: 2, cells: 2 }), counts({}, { ...shape, paragraphs: 1, rows: 1, cells: 1, drawings: 0 }));
  expect(report).toMatchObject({ ok: true, lost: [] });
  expect(report.summary).toContain('could not attribute');
  expect(report.summary).toContain(`${report.unattributed.length} structural element`);
  expect(report.unattributed).toContainEqual(expect.objectContaining({ kind: 'field', name: 'SEQ Figure' }));
  expect(report.unattributed).toContainEqual(expect.objectContaining({ kind: 'element', name: 'w:p', source: 4, saved: 1, explained: 2 }));
  const sameCounts = await compareDocxStructure(bytes, saved, counts({}, { ...shape, paragraphs: 3, rows: 2, cells: 2 }), counts({}, { ...shape, paragraphs: 2, rows: 2, cells: 2 }));
  expect(sameCounts.ok).toBe(false);
});

test('model counts follow an alternate main part', async () => {
  const bytes = await mutate(await source(), async zip => {
    zip.file('word/document2.xml', await zip.file('word/document.xml')!.async('string'));
    zip.remove('word/document.xml');
    zip.file('word/_rels/document2.xml.rels', await zip.file('word/_rels/document.xml.rels')!.async('string'));
    zip.remove('word/_rels/document.xml.rels');
    await rewrite('_rels/.rels', xml => xml.replace('word/document.xml', 'word/document2.xml'))(zip);
    await rewrite('[Content_Types].xml', xml => xml.replace('/word/document.xml', '/word/document2.xml'))(zip);
  });
  const saved = await mutate(bytes, rewrite('word/document2.xml', xml => xml.replace(BODY[6], '')));
  const report = await compareDocxStructure(bytes, saved, counts(), counts(undefined, { paragraphs: 6 }));
  expect(report.lost).toEqual([]);
  expect(report.ok).toBe(true);
});

test('a glossary footnotes part keeps its own strict allowance', async () => {
  const notes = (field: string) => `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p>${field}</w:p></w:footnote></w:footnotes>`;
  const page = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
  const bytes = await mutate(await source(), async zip => {
    zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:footnote></w:footnotes>`);
    zip.file('word/glossary/document.xml', `<w:glossaryDocument xmlns:w="${W}"/>`);
    zip.file('word/glossary/footnotes.xml', notes(page));
    zip.file('word/glossary/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="n" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`);
    await rewrite('word/_rels/document.xml.rels', xml => xml.replace('</Relationships>', `<Relationship Id="n" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="g" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`))(zip);
  });
  const saved = await mutate(bytes, async zip => {
    zip.file('word/glossary/footnotes.xml', notes('<w:r><w:t>1</w:t></w:r>'));
    await rewrite('word/footnotes.xml', xml => xml.replace('<w:tr><w:tc><w:p/></w:tc></w:tr>', ''))(zip);
  });
  const note = (rows: number) => ({ 'word/footnotes.xml': part({ paragraphs: rows, tables: 1, rows, cells: rows }) });
  const report = await compareDocxStructure(bytes, saved, counts(undefined, {}, note(2)), counts(undefined, {}, note(1)));
  expect(report.ok).toBe(false);
  expect(report.lost).toContainEqual(expect.objectContaining({ part: 'word/glossary/footnotes.xml', kind: 'field', name: 'PAGE' }));
});

test('strict namespaces are compared like transitional ones', async () => {
  const strict = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  const bytes = await mutate(await source(), rewrite('word/document.xml', xml => xml.replace(`xmlns:w="${W}"`, `xmlns:w="${strict}"`)));
  expect((await compareDocxStructure(bytes, bytes, counts(), counts())).fields).toEqual({ source: 3, saved: 3 });
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>', '<w:r><w:t>1</w:t></w:r>')));
  expect((await compareDocxStructure(bytes, saved, counts(), counts())).lost).toContainEqual(expect.objectContaining({ kind: 'field', name: 'PAGE' }));
});

test('one removed drawing does not explain a second lost drawing', async () => {
  const pict = '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:imagedata r:id="rIdImage"/></v:shape></w:pict></w:r>';
  const bytes = await source([...BODY, `<w:p>${pict}</w:p>`].join(''));
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(drawing('rIdImage'), '').replace(pict, '')));
  const report = await compareDocxStructure(bytes, saved, counts(undefined, { paragraphs: 8, drawings: 2 }), counts(undefined, { paragraphs: 8, drawings: 1 }));
  expect(report.ok).toBe(false);
  expect(report.lost).toContainEqual(expect.objectContaining({ name: 'w:drawing|w:pict|w:object', source: 2, saved: 0, explained: 1 }));
});

test('fallback copies of drawings are counted once', async () => {
  const alternate = `<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps">${drawing('rIdImage').slice(5, -6)}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>`;
  const bytes = await source([...BODY, `<w:p>${alternate}</w:p>`].join(''));
  const report = await compareDocxStructure(bytes, bytes, counts(), counts());
  expect(report.drawings).toEqual({ source: 2, saved: 2 });
  const saved = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(`<w:p>${alternate}</w:p>`, '')));
  expect((await compareDocxStructure(bytes, saved, counts(undefined, { paragraphs: 8, drawings: 2 }), counts(undefined, { paragraphs: 7, drawings: 1 }))).lost).toEqual([]);
  const noFallback = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('<mc:Fallback><w:pict/></mc:Fallback>', '<mc:Fallback/>')));
  expect((await compareDocxStructure(bytes, noFallback, counts(undefined, { paragraphs: 8, drawings: 2 }), counts(undefined, { paragraphs: 8, drawings: 2 }))).lost)
    .toContainEqual(expect.objectContaining({ kind: 'element', name: 'mc:Fallback w:drawing|w:pict|w:object', source: 1, saved: 0 }));
});

test('equations, quoted field arguments, content types and relationship modes are checked', async () => {
  const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const math = `<w:p><m:oMath xmlns:m="${M}"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`;
  const quote = '<w:p><w:fldSimple w:instr=\' QUOTE "A  B" \'><w:r><w:t>A  B</w:t></w:r></w:fldSimple></w:p>';
  const bytes = await mutate(await source([...BODY, math, quote].join('')), async zip => {
    zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p/></w:hdr>`);
    await rewrite('[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'))(zip);
  });
  const fields = { PAGE: 1, 'REF _Ref1 \\h': 1, DATE: 1, 'QUOTE "A  B"': 1 };
  const model = counts(fields, { paragraphs: 9 });
  expect(normalizeInstruction(' QUOTE  "A  B"  \\* MERGEFORMAT ')).toBe('QUOTE "A  B" \\* MERGEFORMAT');
  expect((await compareDocxStructure(bytes, bytes, model, model)).ok).toBe(true);
  const flattened = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace(math, '<w:p><w:r><w:t>x</w:t></w:r></w:p>')));
  expect((await compareDocxStructure(bytes, flattened, model, model)).lost).toContainEqual(expect.objectContaining({ kind: 'element', name: 'm:oMath' }));
  const requoted = await mutate(bytes, rewrite('word/document.xml', xml => xml.replace('QUOTE "A  B"', 'QUOTE "A B"')));
  expect((await compareDocxStructure(bytes, requoted, model, model)).lost).toContainEqual(expect.objectContaining({ kind: 'field', name: 'QUOTE "A  B"' }));
  const untyped = await mutate(bytes, rewrite('[Content_Types].xml', xml => xml.replace(/<Override PartName="\/word\/header1.xml"[^>]*\/>/, '')));
  expect((await compareDocxStructure(bytes, untyped, model, model)).lost).toContainEqual(expect.objectContaining({ kind: 'contentType', part: 'word/header1.xml' }));
  const external = await mutate(bytes, rewrite('word/_rels/document.xml.rels', xml => xml.replace('Target="media/image1.png"/>', 'Target="media/image1.png" TargetMode="External"/>')));
  expect((await compareDocxStructure(bytes, external, model, model)).lost).toContainEqual(expect.objectContaining({ kind: 'relationship', name: expect.stringContaining('rIdImage') }));
});

test('model counts include fields nested in content controls', () => {
  const field = (instruction: string) => ({ kind: 'embed', embedKind: 'field', payload: { modelKind: 'field', instruction } });
  const session = {
    listComments: () => [],
    storyIds: () => ['body'],
    storySegments: () => [
      { kind: 'embed', embedKind: 'sdt', payload: { content: [field(' PAGE '), { kind: 'image', payload: { src: 'x' } }, { kind: 'sdt', payload: { content: [] } }] }, attributes: {} },
      { kind: 'embed', embedKind: 'field', payload: { modelKind: 'field', instruction: ' DATE ', resultProjection: { children: [{ items: [field(' DATE ')] }] } }, attributes: {} },
      { kind: 'pilcrow', paraId: '1', properties: {}, attributes: {} },
    ],
  } as unknown as YrsSession;
  expect(docxModelCounts(session)).toMatchObject({ comments: 0, parts: { 'word/document.xml': { paragraphs: 1, controls: 2, drawings: 1, fields: { PAGE: 1, DATE: 1 } } } });
});

test('many relationships compare in linear time', async () => {
  const links = Array.from({ length: 50_000 }, (_, index) => `<Relationship Id="rIdL${index}" Type="${R}/hyperlink" Target="https://example.com/${index}" TargetMode="External"/>`).join('');
  const body = Array.from({ length: 50_000 }, (_, index) => `<w:hyperlink r:id="rIdL${index}"><w:r><w:t>${index}</w:t></w:r></w:hyperlink>`).join('');
  const bytes = await mutate(await source(`<w:p>${body}</w:p>`), rewrite('word/_rels/document.xml.rels', xml => xml.replace('</Relationships>', `${links}</Relationships>`)));
  const report = await compareDocxStructure(bytes, bytes, counts({}, { paragraphs: 1, drawings: 0 }), counts({}, { paragraphs: 1, drawings: 0 }));
  expect(report).toMatchObject({ ok: true, relationships: { source: 50_003, saved: 50_003 } });
  expect(report.ms).toBeLessThan(10_000);
}, 60_000);

test('office_verify reports a clean structure after agent edits', async () => {
  const doc = await open(await source());
  const before = doc.list();
  const intro = before.items.find(paragraph => paragraph.text.startsWith('Introduction'))!;
  const spare = before.items.find(paragraph => paragraph.text.startsWith('Spare'))!;
  const match = doc.grep({ query: 'Introduction text that an agent translates.' }).matches[0];
  const proposal = doc.propose({ author: 'agent', edits: [{ match: match.match, newText: 'Einleitungstext, den ein Agent übersetzt.' }] });
  const pending = await doc.verify(proposal.id);
  expect(pending).toMatchObject({ reopened: true, checks: ['save', 'reopen', 'structure'], structure: { ok: true, lost: [], unattributed: [] } });
  await doc.accept(proposal.id);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_paragraph_style', paragraph: intro.ref, style: 'Quote' }, { op: 'delete_paragraph', paragraph: spare.ref }] });
  const verified = await doc.verify();
  expect(verified.structure).toMatchObject({ ok: true, lost: [], unattributed: [], fields: { source: 3, saved: 3 }, drawings: { source: 1, saved: 1 } });
  expect(verified.structure!.ms).toBeGreaterThanOrEqual(0);
  expect(await doc.verify(undefined, { structure: false })).toMatchObject({ checks: ['save', 'reopen'], unchecked: ['visual fidelity', 'semantic correctness', 'structure'] });
}, 30_000);

test('office_verify returns the structure report over MCP', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-verify-structure-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'docx-verify-structure', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  try {
    await writeFile(join(root, 'report.docx'), await source());
    const call = async (name: string, args: Record<string, unknown>) => {
      const output = await client.callTool({ name, arguments: args });
      expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
      return output.structuredContent as Record<string, any>;
    };
    const { document } = await call('office_open', { path: 'report.docx' });
    const verified = await call('office_verify', { document });
    expect(verified).toMatchObject({ checks: ['save', 'reopen', 'structure'], structure: { ok: true, lost: [], fields: { source: 3, saved: 3 }, relationships: { source: 3, saved: 3 } } });
    expect(await call('office_verify', { document, structure: false })).toMatchObject({ checks: ['save', 'reopen'] });
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
}, 30_000);

test('structure check cost on a large document', async () => {
  const paragraphs = Array.from({ length: 4000 }, (_, index) => index % 10 === 0
    ? `<w:p><w:r><w:t xml:space="preserve">Record ${index} page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>${index}</w:t></w:r></w:fldSimple></w:p>`
    : `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Record ${index}: </w:t></w:r><w:r><w:t>the assessment remains unchanged and the figures are carried forward.</w:t></w:r></w:p>`);
  const doc = await open(await source(paragraphs.join('')));
  const match = doc.grep({ query: 'Record 5: ' }).matches[0];
  await doc.accept(doc.propose({ author: 'agent', edits: [{ match: match.match, newText: 'Eintrag 5: ' }] }).id);
  await doc.verify(undefined, { structure: false });
  const started = performance.now();
  const plain = await doc.verify(undefined, { structure: false });
  const plainMs = performance.now() - started;
  const checkedStart = performance.now();
  const checked = await doc.verify();
  const checkedMs = performance.now() - checkedStart;
  expect(plain.checks).toEqual(['save', 'reopen']);
  expect(checked.structure).toMatchObject({ ok: true, lost: [], fields: { source: 400, saved: 400 } });
  const warmStart = performance.now();
  const warm = await doc.verify();
  const warmMs = performance.now() - warmStart;
  expect(warm.structure!.ok).toBe(true);
  console.log(`VERIFY-COST paragraphs=4000 plain=${plainMs.toFixed(0)}ms structure-first=${checkedMs.toFixed(0)}ms structure-warm=${warmMs.toFixed(0)}ms compare=${warm.structure!.ms}ms`);
}, 120_000);
