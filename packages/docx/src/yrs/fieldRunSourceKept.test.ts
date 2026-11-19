import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx, repackDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { ComplexField, Document, Paragraph, SimpleField } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, type YrsLoc, type YrsSession, type YrsStorySegment } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARK = '<w:lastRenderedPageBreak/>';
const field = (code: string) =>
  '<w:fldChar w:fldCharType="begin"/>' +
  `<w:instrText xml:space="preserve"> ${code} </w:instrText>` +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:fldChar w:fldCharType="end"/>';
const FIELDS = `${field('PAGE')}<w:t xml:space="preserve"> of </w:t>${field('NUMPAGES')}`;
const RUN = `<w:r>${FIELDS}</w:r>`;
const MARKED_RUN = `<w:r>${MARK}${FIELDS}</w:r>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const BODY =
  `<w:p>${text('Prev')}</w:p><w:p>${text('AB')}${RUN}${text('CD')}</w:p>` +
  `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${text('Two')}</w:p>`;
const sessions: YrsSession[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

function fixture(body = BODY): Uint8Array<ArrayBuffer> {
  const section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}${section}</w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function session(): Promise<YrsSession> {
  const created = await createYrsSession({ clientId: 97500 + sessions.length });
  sessions.push(created);
  return created;
}

function documentXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

interface Opened {
  session: YrsSession;
  save(): Promise<string>;
}

async function projectedSession(parsed: Document): Promise<Opened> {
  const opened = await session();
  documentToYrs(opened, parsed);
  return {
    session: opened,
    save: async () => documentXml(new Uint8Array(await repackDocx(yrsToDocument(opened, parsed)))),
  };
}

type Loader = (bytes: Uint8Array<ArrayBuffer>) => Promise<Opened>;

const nativeSeed: Loader = async (bytes) => {
  const opened = await session();
  opened.seedFromDocx(bytes);
  return {
    session: opened,
    save: async () => documentXml((await saveYrsDocx(opened, { updateModifiedDate: false })).bytes),
  };
};

const parsedProjection: Loader = async (bytes) =>
  projectedSession(await parseDocx(bytes.buffer, { preloadFonts: false }));

/** The parsed document without any kept run, as main parses it. */
const mainProjection: Loader = async (bytes) => {
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  withoutSource(parsed);
  return projectedSession(parsed);
};

const LOADERS: Array<[string, Loader]> = [
  ['native seed', nativeSeed],
  ['parseDocx and documentToYrs', parsedProjection],
];

function withoutSource(value: unknown): void {
  if (Array.isArray(value)) {
    for (const child of value) withoutSource(child);
  } else if (value !== null && typeof value === 'object') {
    const object = value as Record<string, unknown>;
    delete object.sourceRunXml;
    delete object.sourceRunDigest;
    for (const child of Object.values(object)) withoutSource(child);
  }
}

/** The `index`th paragraph written, without the ids each session mints. */
function paragraphXml(xml: string, index: number): string {
  const paragraph = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g)?.[index] ?? '';
  return paragraph.replace(/ (?:w:id|w14:paraId|w14:textId|w:rsid\w*)="[^"]*"/g, '');
}

async function fullRewrite(bytes: Uint8Array<ArrayBuffer>): Promise<string> {
  const opened = await session();
  opened.openDocx(bytes, true);
  const projected = yrsToDocument(opened, opened.materializeDocx()!);
  return documentXml(new Uint8Array(await saveEditorDocument(opened, projected, [], { full: true })));
}

const at = (session: YrsSession, paragraph: number, offset: number): YrsLoc => ({
  story: 'body',
  paraId: session.paragraphs('body')[paragraph]!.paraId,
  offset,
});

const remove = (session: YrsSession, paragraph: number, start: number, end: number) => {
  session.deleteRange({ story: 'body', start: at(session, paragraph, start), end: at(session, paragraph, end) });
};

const EDITS: Array<[string, (session: YrsSession) => void, [string, string]]> = [
  ['typing before it', (s) => { s.insertText(at(s, 1, 1), 'X'); }, ['AXB', 'CD']],
  ['typing after it', (s) => { s.insertText(at(s, 1, 4), 'Y'); }, ['AB', 'CYD']],
  ['deleting a character before it', (s) => remove(s, 1, 0, 1), ['B', 'CD']],
  ['deleting a character after it', (s) => remove(s, 1, 4, 5), ['AB', 'C']],
  ['a split before it', (s) => { s.splitParagraph(at(s, 1, 1)); }, ['B', 'CD']],
  ['a split after it', (s) => { s.splitParagraph(at(s, 1, 4)); }, ['AB', 'C']],
  ['a join with the previous paragraph', (s) => { s.mergeParagraphs('body', s.paragraphs('body')[0]!.paraId); }, ['PrevAB', 'CD']],
];

function visible(xml: string): string {
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');
}

/** The text written before and after the one copy of `run` in its paragraph. */
function around(xml: string, run = RUN): [string, string] | null {
  const at = xml.indexOf(run);
  if (at < 0 || xml.indexOf(run, at + 1) >= 0) return null;
  const start = Math.max(xml.lastIndexOf('<w:p>', at), xml.lastIndexOf('<w:p ', at));
  const end = xml.indexOf('</w:p>', at);
  return [visible(xml.slice(start, at)), visible(xml.slice(at + run.length, end))];
}

function fieldEmbeds(session: YrsSession): Array<Extract<YrsStorySegment, { kind: 'embed' }>> {
  return session.storySegments('body').filter(
    (segment): segment is Extract<YrsStorySegment, { kind: 'embed' }> =>
      segment.kind === 'embed' && segment.embedKind === 'field'
  );
}

function complexFieldOf(document: Document, paragraph: number): ComplexField {
  const content = (document.package.document.content[paragraph] as Paragraph).content;
  return content.find((child): child is ComplexField => child.type === 'complexField')!;
}

const sourceRun = (value: object) => (value as Record<string, unknown>).sourceRunXml;

for (const [loader, load] of LOADERS) {
  test(`${loader}: the session shows for the run what main shows`, async () => {
    const { session } = await load(fixture());
    expect(session.paragraphs('body').map(({ text }) => text)).toEqual(['Prev', 'ABCD', 'Two']);
    const fields = fieldEmbeds(session);
    expect(fields.map(({ payload }) => [payload.fieldType, payload.instruction])).toEqual([['PAGE', 'PAGE  NUMPAGES']]);
    expect(session.searchText('AB')).toHaveLength(1);
    expect(session.searchText('BC')).toHaveLength(0);
    expect(session.searchText(' of ')).toHaveLength(0);
  });

  test(`${loader}: a save without an edit keeps the run in place`, async () => {
    const opened = await load(fixture());
    expect(around(await opened.save())).toEqual(['AB', 'CD']);
  });

  for (const [name, edit, expected] of EDITS) {
    test(`${loader}: a save after ${name} keeps the run in place`, async () => {
      const opened = await load(fixture());
      edit(opened.session);
      expect(around(await opened.save())).toEqual(expected);
    });

    test(`${loader}: a save after undoing ${name} keeps the run in place`, async () => {
      const opened = await load(fixture());
      opened.session.beginUndoCapture();
      edit(opened.session);
      expect(opened.session.undo()).toBe(true);
      expect(around(await opened.save())).toEqual(['AB', 'CD']);
    });
  }

  test(`${loader}: deleting the field removes the run`, async () => {
    const opened = await load(fixture());
    remove(opened.session, 1, 2, 3);
    const xml = await opened.save();
    expect(xml).not.toContain('NUMPAGES');
    expect(xml).not.toContain(' of ');
    expect(visible(xml)).toBe('PrevABCDTwo');
  });
}

test('a full rewrite keeps the run in place', async () => {
  expect(around(await fullRewrite(fixture()))).toEqual(['AB', 'CD']);
});

test('a written result drops the source run and the field is written from the model', async () => {
  const parsed = await parseDocx(fixture().buffer, { preloadFonts: false });
  const original = complexFieldOf(parsed, 1);
  expect(sourceRun(original)).toBe(RUN);
  const nested: SimpleField = { type: 'simpleField', instruction: 'DATE', fieldType: 'DATE', content: [] };
  const paragraph = parsed.package.document.content[1] as Paragraph;
  paragraph.content = paragraph.content.map((child) =>
    child === original ? { ...original, structuredResult: { inline: [nested] } } : child);
  const opened = await projectedSession(parsed);
  expect(fieldEmbeds(opened.session)).toHaveLength(2);
  remove(opened.session, 1, 2, 3);
  expect(sourceRun(complexFieldOf(yrsToDocument(opened.session, parsed), 1))).toBeUndefined();
  const xml = await opened.save();
  expect(xml).not.toContain(RUN);
  expect(xml).toContain('PAGE  NUMPAGES');
});

const AUTHOR = { name: 'W', date: '2026-01-01T00:00:00Z' };

for (const [loader, load] of LOADERS) {
  test(`${loader}: a suggested deletion across the field is saved as main saves it`, async () => {
    const saves: string[] = [];
    for (const open of [load, mainProjection]) {
      const opened = await open(fixture());
      opened.session.deleteRange({ story: 'body', start: at(opened.session, 1, 1), end: at(opened.session, 1, 4) }, AUTHOR);
      saves.push(paragraphXml(await opened.save(), 1));
    }
    const [saved, main] = saves;
    expect(saved).toContain('w:delText');
    expect(saved).not.toContain(' of ');
    expect(saved).not.toContain(RUN);
    if (load === parsedProjection) expect(saved).toBe(main!);
  });
}

const FIELD_EDITS: Array<[string, (field: ComplexField) => void]> = [
  ['an edited instruction', (field) => { field.instruction = 'DATE'; }],
  ['a locked field', (field) => { field.fldLock = true; }],
  ['bold field properties', (field) => { field.formatting = { ...field.formatting, bold: true }; }],
];

const EDIT_ROUTES: Array<[string, (edit: (document: Document) => void) => Promise<string>]> = [
  ['parseDocx and documentToYrs', async (edit) => {
    const parsed = await parseDocx(fixture().buffer, { preloadFonts: false });
    const opened = await session();
    documentToYrs(opened, parsed);
    const projected = yrsToDocument(opened, parsed);
    edit(projected);
    return documentXml(new Uint8Array(await repackDocx(projected)));
  }],
  ['the native session, a full rewrite', async (edit) => {
    const opened = await session();
    opened.openDocx(fixture(), true);
    const projected = yrsToDocument(opened, opened.materializeDocx()!);
    edit(projected);
    return documentXml(new Uint8Array(await saveEditorDocument(opened, projected, [], { full: true })));
  }],
];

for (const [route, save] of EDIT_ROUTES) {
  for (const [name, edit] of FIELD_EDITS) {
    test(`${route}: ${name} is written as main writes it`, async () => {
      const saved = paragraphXml(await save((document) => edit(complexFieldOf(document, 1))), 1);
      const main = paragraphXml(await save((document) => {
        const field = complexFieldOf(document, 1);
        edit(field);
        withoutSource(field);
      }), 1);
      expect(saved).toContain('AB');
      expect(saved).not.toContain(RUN);
      expect(saved).toBe(main);
    });
  }
}

const BOLD = '<w:rPr><w:b/></w:rPr>';
const BOLD_RUN = `<w:r>${BOLD}${FIELDS}</w:r>`;
const BOLD_MARKED_RUN = `<w:r>${BOLD}${MARK}${FIELDS}</w:r>`;

/** Each body, its run as read, the run a written paragraph holds, and the text main writes the mark before. */
const MARKED: Array<[string, string, string, string, string]> = [
  ['inside the run', `<w:p>${MARKED_RUN}${text('CD')}</w:p>`, MARKED_RUN, RUN, 'CD'],
  ['inside the run after its properties', `<w:p>${BOLD_MARKED_RUN}${text('CD')}</w:p>`, BOLD_MARKED_RUN, BOLD_RUN, 'CD'],
  ['before the run', `<w:p><w:r>${MARK}<w:t>AB</w:t></w:r>${RUN}${text('CD')}</w:p>`, RUN, RUN, 'AB'],
];

const typingAfter = (load: Loader) => async (bytes: Uint8Array<ArrayBuffer>) => {
  const opened = await load(bytes);
  const paragraph = opened.session.paragraphs('body')[0]!;
  opened.session.insertText(at(opened.session, 0, paragraph.text.length + 1), 'Z');
  return opened.save();
};

/** Each route, and whether it writes the paragraph from the model rather than keeping its source. */
const MARK_ROUTES: Array<[string, boolean, (bytes: Uint8Array<ArrayBuffer>) => Promise<string>]> = [
  ['a full rewrite', true, fullRewrite],
  ['native seed, a save without an edit', false, async (bytes) => (await nativeSeed(bytes)).save()],
  ['native seed, typing after the run', true, typingAfter(nativeSeed)],
  ['parseDocx and documentToYrs, a save without an edit', true, async (bytes) => (await parsedProjection(bytes)).save()],
  ['parseDocx and documentToYrs, typing after the run', true, typingAfter(parsedProjection)],
];

for (const [where, body, read, kept, before] of MARKED) {
  for (const [route, written, save] of MARK_ROUTES) {
    test(`a rendered page break ${where} is written where main writes it: ${route}`, async () => {
      const xml = await save(fixture(body));
      expect(xml.split(MARK).length - 1).toBe(1);
      if (!written) {
        expect(xml).toContain(read);
        return;
      }
      expect(xml).toContain(kept);
      expect(xml).not.toContain(`<w:r>${MARK}<w:rPr>`);
      const mark = xml.indexOf(MARK);
      expect(visible(xml.slice(mark, xml.indexOf('</w:p>', mark))).startsWith(before)).toBe(true);
    });
  }
}
