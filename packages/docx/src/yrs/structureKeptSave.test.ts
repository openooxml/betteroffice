import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const PAGE_FIELD = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
const SMART_TAG = '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Place</w:t></w:r></w:smartTag>';
const TARGET = `<w:p><w:r><w:rPr><w:snapToGrid w:val="0"/></w:rPr><w:t>Target</w:t></w:r>${SMART_TAG}</w:p>`;
const makeDocx = (body: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${SECT}</w:body></w:document>`)],
])));
const partOf = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });
async function cleared() {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(makeDocx(`<w:p><w:r><w:t>Before</w:t></w:r></w:p>${TARGET}<w:p><w:r><w:t>After</w:t></w:r></w:p>`));
  const target = (session.paragraphs('body') as Array<{ paraId: string; text: string }>).find(entry => entry.text.startsWith('Target'))!;
  session.clearFormatting({ story: 'body', start: { paraId: target.paraId, offset: 0 }, end: { paraId: target.paraId, offset: 6 } });
  return session;
}

test('saveYrsDocx keeps a paragraph whose rewrite would drop a smart tag, and reports it', async () => {
  const session = await cleared();
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  expect(partOf(saved.bytes)).toContain(TARGET);
  expect(saved.keptUnchanged?.structure).toBe(1);
  expect(saved.keptUnchanged?.paragraphs).toEqual([expect.objectContaining({ part: 'word/document.xml', ordinal: 1, text: 'TargetPlace', scope: 'paragraph', dropped: ['w:smartTag'] })]);
});

test('the editor save keeps a paragraph whose rewrite would drop a smart tag', async () => {
  const session = await cleared();
  const projected = yrsToDocument(session, session.materializeDocx()!);
  const bytes = new Uint8Array(await saveEditorDocument(session, projected, projected.package.document.comments ?? []));
  expect(partOf(bytes)).toContain(TARGET);
});

test('a formatting change the writer can patch is saved and reports nothing kept', async () => {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(makeDocx(`<w:p><w:r><w:t>Target</w:t></w:r>${SMART_TAG}</w:p>`));
  const target = (session.paragraphs('body') as Array<{ paraId: string; text: string }>)[0]!;
  session.formatRange({ story: 'body', start: { paraId: target.paraId, offset: 0 }, end: { paraId: target.paraId, offset: 6 } }, { bold: true });
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  const xml = partOf(saved.bytes);
  expect(xml).toContain(SMART_TAG);
  expect(xml).toContain('<w:b/>');
  expect(saved.keptUnchanged).toBeUndefined();
});

const seeded = async (paragraph: string) => {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(makeDocx(paragraph));
  const target = (session.paragraphs('body') as Array<{ paraId: string }>)[0]!;
  return { session, range: { story: 'body', start: { paraId: target.paraId, offset: 0 }, end: { paraId: target.paraId, offset: 4 } } };
};

test('rejecting a deletion is saved, not kept: restored text is not a formatting change', async () => {
  const { session, range } = await seeded('<w:p><w:del w:id="3" w:author="Ada"><w:r><w:delText>Link</w:delText></w:r></w:del></w:p>');
  session.rejectChange(range);
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  const xml = partOf(saved.bytes);
  expect(xml).not.toContain('<w:del ');
  expect(xml).toContain('>Link</w:t>');
  expect(saved.keptUnchanged).toBeUndefined();
});

test('removing a hyperlink is saved, not kept', async () => {
  const { session, range } = await seeded('<w:p><w:hyperlink w:anchor="target"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>');
  session.setHyperlink(range, null);
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  const xml = partOf(saved.bytes);
  expect(xml).not.toContain('<w:hyperlink');
  expect(xml).toContain('>Link</w:t>');
  expect(saved.keptUnchanged).toBeUndefined();
});

test('a new comment on a smart tag paragraph is saved with its anchors, not kept', async () => {
  const { session, range } = await seeded(`<w:p>${SMART_TAG.replace('Place', 'Link')}</w:p>`);
  session.addComment([range], 'Ada', '2026-09-22', []);
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  const xml = partOf(saved.bytes);
  expect(xml).toContain('<w:commentRangeStart');
  expect(xml).toContain('<w:commentRangeEnd');
  expect(saved.keptUnchanged).toBeUndefined();
});

test('deleting a hyperlink is saved, not kept: a text edit is not a formatting change', async () => {
  const session = await createYrsSession();
  open.push(session);
  const link = '<w:hyperlink w:anchor="target"><w:r><w:t>Link</w:t></w:r></w:hyperlink>';
  session.seedFromDocx(makeDocx(`<w:p><w:r><w:t>Target </w:t></w:r>${link}${PAGE_FIELD}</w:p>`));
  const target = (session.paragraphs('body') as Array<{ paraId: string; text: string }>)[0]!;
  session.replaceRange({ story: 'body', start: { paraId: target.paraId, offset: 7 }, end: { paraId: target.paraId, offset: 11 } }, '');
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  const xml = partOf(saved.bytes);
  expect(xml).not.toContain('<w:hyperlink');
  expect(xml).not.toContain('>Link<');
  expect(saved.keptUnchanged).toBeUndefined();
});
