import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { openDocx, type DocxAgentDocument } from '../src';
import { docxBatchSchema } from '../src/docx-schema';
import { createOfficeMcpServer } from '../src/mcp';
import { editFixture } from './fixture';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const LANG = '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>';
const PROPS = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val='24'/>${LANG}`;
const run = (text: string, properties = PROPS) => `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: number, text: string) => `<w:p w14:paraId="${id.toString(16).padStart(8, '0')}"><w:pPr><w:rPr>${PROPS}</w:rPr></w:pPr>${run(text)}</w:p>`;
const TAIL = paragraph(3, 'Untouched tail.');
const BODY = `<w:p w14:paraId="00000001"><w:pPr><w:rPr>${PROPS}</w:rPr></w:pPr>${run('Hello ')}${run('world.')}</w:p>${paragraph(2, 'Other text.')}${TAIL}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph(4, 'Cell text.')}</w:tc></w:tr></w:tbl>`;
const STYLES = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${PROPS}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>${PROPS}</w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:rPr>${LANG}</w:rPr></w:pPr><w:rPr>${PROPS}</w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style></w:styles>`;

async function fixture(body = BODY, styles = STYLES) {
  const zip = await JSZip.loadAsync(await editFixture());
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  zip.file('word/styles.xml', styles);
  return zip.generateAsync({ type: 'uint8array' });
}

const opened: DocxAgentDocument[] = [];
afterEach(() => opened.splice(0).forEach(doc => doc.close()));
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await fixture());
  opened.push(doc);
  return doc;
}
async function parts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  return { document: await zip.file('word/document.xml')!.async('string'), styles: await zip.file('word/styles.xml')!.async('string') };
}
const withoutLanguage = (xml: string) => xml.replace(/<w:lang\b[^>]*\/>/g, '');
const languages = (xml: string) => (xml.match(/<w:lang\b[^>]*\/>/g) ?? []).map(tag => Object.fromEntries([...tag.matchAll(/w:(val|eastAsia|bidi)="([^"]*)"/g)].map(([, key, value]) => [key, value])));
const first = (doc: DocxAgentDocument) => doc.list({ limit: 100 }).items[0];

test('paragraph sets every run and the paragraph mark, preserving other slots and XML', async () => {
  const bytes = await fixture();
  const before = await parts(bytes);
  const doc = await open(bytes);
  const version = doc.list().version;
  const result = await doc.edit({ version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] });
  expect(result).toEqual({ applied: true, version: doc.list().version, results: [{ index: 0, op: 'set_language', ref: first(doc).ref }] });
  expect(result.version).not.toBe(version);
  const after = await parts(await doc.export());
  expect(after.document).toBe(before.document.replace(before.document.match(/<w:p\b[\s\S]*?<\/w:p>/)![0], value => value.replaceAll('w:val="en-US"', 'w:val="de-DE"')));
  expect(after.styles).toBe(before.styles);
  expect(after.document).toContain(TAIL);
  expect(languages(after.document).slice(0, 3)).toEqual(Array(3).fill({ val: 'de-DE', eastAsia: 'ja-JP', bidi: 'ar-SA' }));
  expect(withoutLanguage(after.document)).toBe(withoutLanguage(before.document));
  expect(doc.session.undo()).toBe(true);
  expect((await parts(await doc.export())).document).toBe(before.document);
  expect(doc.session.redo()).toBe(true);
  expect((await parts(await doc.export())).document).toBe(after.document);
});

test('editing another paragraph preserves val-only run and mark language exactly', async () => {
  const language = '<w:lang w:val="fr-FR"/>';
  const untouched = `<w:p w14:paraId="00000002"><w:pPr><w:rPr>${language}</w:rPr></w:pPr>${run('Untouched French.', PROPS.replace(LANG, language))}</w:p>`;
  const target = paragraph(1, 'Hello world.');
  const bytes = await fixture(target + untouched);
  const before = await parts(bytes);
  const doc = await open(bytes);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] });
  const after = await parts(await doc.export());
  const saved = after.document.match(/<w:p\b[\s\S]*?<\/w:p>/g)![1]!;
  expect(saved).toBe(untouched);
  const untouchedRun = saved.match(/<w:r\b[\s\S]*?<\/w:r>/)![0];
  expect(untouchedRun.match(/<w:lang\b[^>]*\/>/)?.[0]).toBe('<w:lang w:val="fr-FR"/>');
  expect(saved.match(/<w:lang\b[^>]*\/>/g)).toEqual([language, language]);
  expect(after.document).toBe(before.document.replace(target, target.replaceAll('w:val="en-US"', 'w:val="de-DE"')));
  expect(after.styles).toBe(before.styles);
});

test('match splits runs at both boundaries and leaves the paragraph mark alone', async () => {
  const doc = await open();
  const found = doc.grep({ query: 'lo wor' });
  const result = await doc.edit({ version: found.version, edits: [{ op: 'set_language', match: found.matches[0].match, lang: 'de-DE' }] });
  expect(result.results[0]).toMatchObject({ ref: first(doc).ref, start: 3, end: 9 });
  const xml = (await parts(await doc.export())).document;
  const firstParagraph = xml.match(/<w:p\b[\s\S]*?<\/w:p>/)![0];
  const runs = firstParagraph.match(/<w:r\b[\s\S]*?<\/w:r>/g)!;
  expect(runs.map(value => value.match(/<w:t\b[^>]*>(.*?)<\/w:t>/)![1])).toEqual(['Hel', 'lo ', 'wor', 'ld.']);
  expect(runs.map(value => languages(value)[0])).toEqual(['en-US', 'de-DE', 'de-DE', 'en-US'].map(val => ({ val, eastAsia: 'ja-JP', bidi: 'ar-SA' })));
  expect(languages(firstParagraph)[0].val).toBe('en-US');
  expect(xml).toContain(TAIL);
});

test('each source run keeps its omitted language slots', async () => {
  const properties = PROPS.replace('ja-JP', 'zh-TW').replace('ar-SA', 'he-IL');
  const body = `<w:p w14:paraId="00000001"><w:pPr><w:rPr>${PROPS}</w:rPr></w:pPr>${run('First ')}${run('second.', properties)}</w:p>`;
  const bytes = await fixture(body + TAIL);
  const doc = await open(bytes);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] });
  const xml = (await parts(await doc.export())).document;
  expect(xml).toBe((await parts(bytes)).document.replace(body, body.replaceAll('w:val="en-US"', 'w:val="de-DE"')));
  expect(languages(xml).slice(0, 3)).toEqual([
    { val: 'de-DE', eastAsia: 'ja-JP', bidi: 'ar-SA' },
    { val: 'de-DE', eastAsia: 'ja-JP', bidi: 'ar-SA' },
    { val: 'de-DE', eastAsia: 'zh-TW', bidi: 'he-IL' },
  ]);
});

test('table cell paragraphs and empty paragraph marks accept language', async () => {
  const emptyParagraph = `<w:p w14:paraId="00000005"><w:pPr><w:rPr>${PROPS}</w:rPr></w:pPr></w:p>`;
  const doc = await open(await fixture(emptyParagraph + BODY));
  const items = doc.list({ limit: 100 }).items;
  const cell = items.find(item => item.text === 'Cell text.')!;
  const empty = items.find(item => item.text === '')!;
  expect(cell).toBeDefined();
  expect(empty).toBeDefined();
  await doc.edit({ version: doc.list().version, edits: [cell, empty].map(item => ({ op: 'set_language', paragraph: item.ref, lang: 'de-DE' })) });
  const xml = (await parts(await doc.export())).document;
  const changed = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g)!.filter(value => value.includes('Cell text.') || !value.includes('<w:t'));
  expect(changed).toHaveLength(2);
  changed.forEach(value => languages(value).forEach(language => expect(language.val).toBe('de-DE')));
  expect(xml).toContain(TAIL);
});

test('style name targets direct rPr and defaults change only their language elements', async () => {
  const bytes = await fixture();
  const before = await parts(bytes);
  const doc = await open(bytes);
  const result = await doc.edit({ version: doc.list().version, edits: [
    { op: 'set_language', style: 'heading 1', lang: 'de-DE' },
    { op: 'set_language', defaults: true, eastAsia: 'zh-Hant', bidi: null },
    { op: 'set_language', style: 'Strong', lang: 'fr-FR' },
  ] });
  expect(result.results[0].style).toBe('Heading1');
  expect(result.results[2].style).toBe('Strong');
  const after = await parts(await doc.export());
  expect(after.document).toBe(before.document);
  const heading = after.styles.match(/<w:style\b[^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/)![0];
  expect(languages(heading)).toEqual([{ val: 'en-US', eastAsia: 'ja-JP', bidi: 'ar-SA' }, { val: 'de-DE', eastAsia: 'ja-JP', bidi: 'ar-SA' }]);
  expect(languages(after.styles)[0]).toEqual({ val: 'en-US', eastAsia: 'zh-Hant' });
  expect(after.styles).toContain('<w:rPr><w:lang w:val="fr-FR"/></w:rPr>');
  expect(withoutLanguage(after.styles).replace('<w:rPr></w:rPr></w:style></w:styles>', '</w:style></w:styles>')).toBe(withoutLanguage(before.styles));
  const reopened = await open(await doc.export());
  expect((await parts(await reopened.export())).styles).toBe(after.styles);
});

test('null clears one slot and repeated edits retain the others', async () => {
  const doc = await open();
  const target = first(doc).ref;
  await doc.edit({ version: doc.list().version, edits: [
    { op: 'set_language', paragraph: target, lang: 'de-DE' },
    { op: 'set_language', paragraph: target, eastAsia: null },
    { op: 'set_language', style: 'Normal', lang: null },
  ] });
  let after = await parts(await doc.export());
  expect(languages(after.document).slice(0, 3)).toEqual(Array(3).fill({ val: 'de-DE', bidi: 'ar-SA' }));
  expect(after.styles).toContain('<w:lang w:eastAsia="ja-JP" w:bidi="ar-SA"/>');
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_language', paragraph: target, lang: null, bidi: null }] });
  after = await parts(await doc.export());
  expect(after.document.match(/<w:p\b[\s\S]*?<\/w:p>/)![0]).not.toContain('<w:lang');
});

test('language-only edits preserve rPrChange bytes on the minimal patch path', async () => {
  const revision = '<w:rPrChange w:author="Editor" w:id="7"><w:rPr><w:i/></w:rPr></w:rPrChange>';
  const body = paragraph(1, 'Revision').replaceAll(`${LANG}</w:rPr>`, `${LANG}${revision}</w:rPr>`);
  const bytes = await fixture(body + TAIL);
  const doc = await open(bytes);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] });
  expect((await parts(await doc.export())).document).toBe((await parts(bytes)).document.replace(body, body.replaceAll('w:val="en-US"', 'w:val="de-DE"')));
  const plan = doc.session.paragraphSavePlan();
  expect(plan.splicedParts?.flatMap(part => part.legacyParagraphs ?? []) ?? []).toEqual([]);
  expect(plan.splicedParts?.flatMap(part => part.paragraphPatches ?? []).length).toBeGreaterThan(0);
});

test('validation rejects missing or multiple targets, missing slots and malformed tags', () => {
  for (const edit of [
    { op: 'set_language', lang: 'de-DE' },
    { op: 'set_language', paragraph: 'p1', defaults: true, lang: 'de-DE' },
    { op: 'set_language', paragraph: 'p1' },
    ...['', 'de_DE', 'de--DE', 'x', '123', 'a'.repeat(36)].map(lang => ({ op: 'set_language', defaults: true, lang })),
    { op: 'set_language', defaults: false, lang: 'de-DE' },
  ]) expect(docxBatchSchema.safeParse({ version: 0, edits: [edit] }).success).toBe(false);
  for (const lang of ['de-DE', 'zh-Hant-TW', 'ar', null]) {
    expect(docxBatchSchema.safeParse({ version: 0, edits: [{ op: 'set_language', defaults: true, lang }] }).success).toBe(true);
  }
});

test('unknown styles and stale versions refuse the whole batch', async () => {
  const doc = await open();
  const before = await parts(await doc.export());
  const version = doc.list().version;
  await expect(doc.edit({ version, edits: [
    { op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' },
    { op: 'set_language', style: 'Missing', lang: 'de-DE' },
  ] })).rejects.toMatchObject({ code: 'UNKNOWN_STYLE' });
  expect(await parts(await doc.export())).toEqual(before);
  expect(doc.list().version).toBe(version);
  await doc.edit({ version, edits: [{ op: 'set_language', defaults: true, lang: 'de-DE' }] });
  await expect(doc.edit({ version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] })).rejects.toMatchObject({ code: 'STALE_VERSION' });
});

test('missing styles refuses defaults atomically and keeps export and paragraph edits working', async () => {
  const zip = await JSZip.loadAsync(await fixture());
  zip.remove('word/styles.xml');
  zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace(/<Relationship\b[^>]*Type="[^"]*\/styles"[^>]*\/>/, ''));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace(/<Override\b[^>]*PartName="\/word\/styles.xml"[^>]*\/>/, ''));
  const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
  const before = await JSZip.loadAsync(await doc.export());
  const original = await before.file('word/document.xml')!.async('string');
  const version = doc.list().version;
  await expect(doc.edit({ version, edits: [
    { op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' },
    { op: 'set_language', defaults: true, lang: 'de-DE' },
  ] })).rejects.toMatchObject({ code: 'UNSUPPORTED', message: expect.stringContaining('word/styles.xml') });
  expect(doc.list().version).toBe(version);
  expect(doc.session.styleLanguages()).toEqual({});
  await expect(doc.edit({ version, edits: [{ op: 'set_language', style: 'Normal', lang: 'de-DE' }] })).rejects.toMatchObject({ code: 'UNKNOWN_STYLE' });
  const refused = await JSZip.loadAsync(await doc.export());
  expect(refused.file('word/styles.xml')).toBeNull();
  expect(await refused.file('word/document.xml')!.async('string')).toBe(original);

  await doc.edit({ version, edits: [{ op: 'set_language', paragraph: first(doc).ref, lang: 'de-DE' }] });
  const after = await JSZip.loadAsync(await doc.export());
  expect(after.file('word/styles.xml')).toBeNull();
  expect(await after.file('word/document.xml')!.async('string')).toBe(original.replace(original.match(/<w:p\b[\s\S]*?<\/w:p>/)![0], value => value.replaceAll('w:val="en-US"', 'w:val="de-DE"')));
});

test('a later removed target rolls back text and style language edits', async () => {
  const doc = await open();
  const before = await parts(await doc.export());
  const version = doc.list().version;
  const target = first(doc).ref;
  await expect(doc.edit({ version, edits: [
    { op: 'set_language', defaults: true, lang: 'de-DE' },
    { op: 'delete_paragraph', paragraph: target },
    { op: 'set_language', paragraph: target, lang: 'de-DE' },
  ] })).rejects.toMatchObject({ code: 'UNKNOWN_REF' });
  expect(doc.session.styleLanguages()).toEqual({});
  expect(doc.list().version).toBe(version);
  expect(await parts(await doc.export())).toEqual(before);
});

test('MCP translation followed by set_language updates body, Normal and defaults', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-language-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'language-test', version: '1' });
  const [local, remote] = InMemoryTransport.createLinkedPair();
  await server.connect(remote);
  await client.connect(local);
  try {
    const source = await fixture(paragraph(1, 'Hello world.') + paragraph(2, 'Good morning.'));
    await writeFile(join(root, 'input.docx'), source);
    const call = async (name: string, args: Record<string, unknown>) => {
      const result = await client.callTool({ name, arguments: args });
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
      return result.structuredContent as Record<string, any>;
    };
    const { document } = await call('office_open', { path: 'input.docx' });
    for (const [query, newText] of [['Hello world.', 'Hallo Welt.'], ['Good morning.', 'Guten Morgen.']]) {
      const found = await call('office_grep', { document, query });
      const proposal = await call('office_propose', { document, author: 'Translator', edits: [{ match: found.matches[0].match, newText }] });
      await call('office_accept', { document, proposal: proposal.id });
    }
    const outline = await call('office_outline', { document, limit: 100 });
    const edited = await call('docx_edit', { document, version: outline.version, edits: [
      ...outline.items.map((item: any) => ({ op: 'set_language', paragraph: item.ref, lang: 'de-DE' })),
      { op: 'set_language', style: 'Normal', lang: 'de-DE' },
      { op: 'set_language', defaults: true, lang: 'de-DE' },
    ] });
    expect(edited).toMatchObject({ applied: true, version: expect.any(Number) });
    expect(edited.version).not.toBe(outline.version);
    expect(edited.results).toHaveLength(4);
    expect(edited.results.every((result: any) => result.op === 'set_language')).toBe(true);
    await call('office_export', { document, path: 'translated.docx' });
    const before = await parts(source);
    const after = await parts(await readFile(join(root, 'translated.docx')));
    expect(languages(after.document)).toHaveLength(4);
    languages(after.document).forEach(language => expect(language).toEqual({ val: 'de-DE', eastAsia: 'ja-JP', bidi: 'ar-SA' }));
    expect(after.document).toBe(before.document.replace('Hello world.', 'Hallo Welt.').replace('Good morning.', 'Guten Morgen.').replaceAll('w:val="en-US"', 'w:val="de-DE"'));
    const targeted = after.styles.slice(0, after.styles.indexOf('<w:style w:type="paragraph" w:styleId="Heading1"'));
    languages(targeted).forEach(language => expect(language.val).toBe('de-DE'));
    expect(withoutLanguage(after.styles)).toBe(withoutLanguage(before.styles));
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
});
