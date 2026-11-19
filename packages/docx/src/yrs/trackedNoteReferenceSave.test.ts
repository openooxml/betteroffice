import { afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { documentToYrs } from './documentToYrs';
import type { Document } from '../types/document';
import {
  DirtyProjectionStories,
  editorSaveKeptUnchanged,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
} from './editorSave';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { setupStructuralPackageSaves, type SavePath } from './structuralSaveTestSupport';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const DATE = '2025-03-04T05:06:07Z';

type NoteKind = 'footnote' | 'endnote';
type Revision = 'deletion' | 'insertion';
type Apply = (session: YrsSession) => void;

interface Shape {
  name: string;
  note: NoteKind;
  revision: Revision;
  tag: 'del' | 'ins';
  id: number;
  author: string;
}

const SHAPES: Shape[] = (['footnote', 'endnote'] as const).flatMap((note) => [
  { name: `deleted ${note} reference`, note, revision: 'deletion' as const, tag: 'del' as const, id: 12, author: 'Bo' },
  { name: `inserted ${note} reference`, note, revision: 'insertion' as const, tag: 'ins' as const, id: 11, author: 'Ada' },
]);

const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const deletedText = (value: string) => `<w:r><w:delText xml:space="preserve">${value}</w:delText></w:r>`;
const style = (note: NoteKind) => (note === 'footnote' ? 'FootnoteReference' : 'EndnoteReference');
const referenceRun = (note: NoteKind, id = 1) =>
  `<w:r><w:rPr><w:rStyle w:val="${style(note)}"/></w:rPr><w:${note}Reference w:id="${id}"/></w:r>`;
const revisionSpan = (shape: Shape, content: string) =>
  `<w:${shape.tag} w:id="${shape.id}" w:author="${shape.author}" w:date="${DATE}">${content}</w:${shape.tag}>`;
const wrapped = (shape: Shape) => revisionSpan(shape, referenceRun(shape.note));
const paragraph = (id: string, content: string) => `<w:p w14:paraId="${id}" w14:textId="${id}">${content}</w:p>`;
const firstParagraph = (reference: string) => paragraph('0D000001', `${text('Head ')}${reference}${text(' tail')}`);
const SECOND = paragraph('0D000002', text('Other'));
const shapeOf = (note: NoteKind, revision: Revision) =>
  SHAPES.find((shape) => shape.note === note && shape.revision === revision)!;
const fieldParagraph = (note: NoteKind) => paragraph('0D000001',
  `${text('Head ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "target" </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:hyperlink w:anchor="target">${text('Link')}${referenceRun(note)}</w:hyperlink>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>${text(' tail')}`
);
const SUGGESTER = { name: 'B', date: '2025-03-05T00:00:00Z' };
const bareRun = (note: NoteKind, id: number) => `<w:r><w:${note}Reference w:id="${id}"/></w:r>`;
const languageRun = (note: NoteKind, id: number, properties = '') =>
  `<w:r><w:rPr>${properties}<w:lang w:val="fr-FR"/></w:rPr><w:${note}Reference w:id="${id}"/></w:r>`;
const pageBreakRun = (note: NoteKind, id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="${style(note)}"/></w:rPr><w:lastRenderedPageBreak/><w:${note}Reference w:id="${id}"/></w:r>`;

function notesXml(note: NoteKind, ids: number[]): string {
  return `<w:${note}s xmlns:w="${W}">` +
    `<w:${note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${note}>` +
    `<w:${note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${note}>` +
    ids.map((id) =>
      `<w:${note} w:id="${id}"><w:p><w:r><w:rPr><w:rStyle w:val="${style(note)}"/></w:rPr><w:${note}Ref/></w:r>${text(` ${note} body`)}</w:p></w:${note}>`
    ).join('') +
    `</w:${note}s>`;
}

function fixture(first: string, noteIds = [1]): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/>` +
      `<Override PartName="/word/endnotes.xml" ContentType="${OFFICE}.endnotes+xml"/></Types>`
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdFootnotes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdEndnotes" Type="${R}/endnotes" Target="endnotes.xml"/></Relationships>`
    )],
    ['word/styles.xml', toBytes(
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>'
    )],
    ['word/footnotes.xml', toBytes(notesXml('footnote', noteIds))],
    ['word/endnotes.xml', toBytes(notesXml('endnote', noteIds))],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${first}${SECOND}<w:sectPr/></w:body></w:document>`
    )],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function part(bytes: Uint8Array, name: string): string {
  const found = (unzipContainer(bytes) as Record<string, Uint8Array | undefined>)[name];
  return found ? new TextDecoder().decode(found) : '';
}

function bodyParagraphs(bytes: Uint8Array): string[] {
  return part(bytes, 'word/document.xml').match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
}

function savedNote(bytes: Uint8Array, note: NoteKind): string {
  return new RegExp(`<w:${note}\\b[^>]*\\bw:id="1"[^>]*>[\\s\\S]*?</w:${note}>`).exec(part(bytes, `word/${note}s.xml`))?.[0] ?? '';
}

const IN_RPR = '(?:(?!</w:rPr>)[\\s\\S])*';
const RUN_PROPERTIES = `(?:<w:rPr>${IN_RPR}</w:rPr>)?`;
const opening = (shape: Shape) => `<w:${shape.tag} w:id="\\d+" w:author="${shape.author}" w:date="${DATE}">`;
const bareSpan = (shape: Shape) => revisionSpan(shape, bareRun(shape.note, 1));
const BOLD_HEAD = new RegExp(`<w:r><w:rPr>${IN_RPR}<w:b/>${IN_RPR}</w:rPr><w:t(?:\\s[^>]*)?>Head</w:t></w:r>`);

function compound(shape: Shape): { around: string; alone: string; original: string } {
  const words = shape.tag === 'del' ? deletedText : text;
  const around = revisionSpan(shape, `${words('left')}${referenceRun(shape.note)}${words('right')}`);
  const alone = revisionSpan({ ...shape, id: shape.id + 10 }, referenceRun(shape.note, 2));
  return { around, alone, original: firstParagraph(`${around}${text(' and ')}${alone}`) };
}

function wrapperPattern(shape: Shape): RegExp {
  return new RegExp(`${opening(shape)}${bareRun(shape.note, 1)}</w:${shape.tag}>`);
}

function liveReferencePattern(note: NoteKind): RegExp {
  return new RegExp(`<w:r>${RUN_PROPERTIES}<w:${note}Reference w:id="1"/></w:r>`);
}

function wrapperSpan(shape: Shape, xml: string): string {
  return wrapperPattern(shape).exec(xml)?.[0] ?? '';
}

function count(xml: string, pattern: string): number {
  return xml.match(new RegExp(pattern, 'g'))?.length ?? 0;
}

const plainText = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');

function enclosingRun(xml: string, element: string): string {
  return new RegExp(`<w:r>(?:(?!<w:r>)[\\s\\S])*?${element}</w:r>`).exec(xml)?.[0] ?? '';
}

const savedReferenceRun = (xml: string, note: NoteKind, id = 1) => enclosingRun(xml, `<w:${note}Reference w:id="${id}"/>`);
const fieldCharacterRun = (xml: string, type: string) => enclosingRun(xml, `<w:fldChar w:fldCharType="${type}"/>`);

function fieldSpan(xml: string): string {
  const begin = fieldCharacterRun(xml, 'begin');
  const end = fieldCharacterRun(xml, 'end');
  const start = begin ? xml.indexOf(begin) : -1;
  const stop = end ? xml.indexOf(end, start) : -1;
  return start < 0 || stop < 0 ? '' : xml.slice(start, stop + end.length);
}

function noteReferences(session: YrsSession) {
  return session.storySegments('body').flatMap((segment) =>
    segment.kind === 'embed' && segment.embedKind === 'noteRef' ? [segment] : []
  );
}

function typeInto(session: YrsSession, index: number): void {
  const target = session.paragraphs('body')[index];
  if (!target) throw new Error(`no body paragraph ${index}`);
  session.insertText({ story: 'body', paraId: target.paraId, offset: 2 }, 'X');
}

function firstParagraphId(session: YrsSession): string {
  const target = session.paragraphs('body')[0];
  if (!target) throw new Error('no body paragraph');
  return target.paraId;
}

function referenceOffset(session: YrsSession): number {
  let offset = 0;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'pilcrow') break;
    if (segment.kind === 'embed' && segment.embedKind === 'noteRef') return offset;
    offset += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('no note reference in the first paragraph');
}

function suggestDeletingReference(session: YrsSession): void {
  const paraId = firstParagraphId(session);
  const offset = referenceOffset(session);
  session.deleteRange({ story: 'body', start: { paraId, offset }, end: { paraId, offset: offset + 1 } }, SUGGESTER);
  expect(noteReferences(session)[0]?.attributes.del).toMatchObject({ author: SUGGESTER.name });
}

function linkOffset(session: YrsSession): number {
  let offset = 0;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'pilcrow') break;
    if (segment.kind === 'text' && segment.text.includes('Link')) return offset + segment.text.indexOf('Link');
    offset += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('no Link text in the first paragraph');
}

function typeAt(session: YrsSession, offset: number): void {
  session.insertText({ story: 'body', paraId: firstParagraphId(session), offset }, 'X');
}

function boldHead(session: YrsSession): void {
  const paraId = firstParagraphId(session);
  session.toggleMark({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }, { type: 'bold' });
}

function boldWord(word: string): Apply {
  return (session) => {
    const paraId = firstParagraphId(session);
    let offset = 0;
    for (const segment of session.storySegments('body')) {
      if (segment.kind === 'pilcrow') break;
      if (segment.kind === 'text' && segment.text.includes(word)) {
        const start = offset + segment.text.indexOf(word);
        session.toggleMark({ story: 'body', start: { paraId, offset: start }, end: { paraId, offset: start + word.length } }, { type: 'bold' });
        return;
      }
      offset += segment.kind === 'text' ? segment.text.length : 1;
    }
    throw new Error(`no ${word} text in the first paragraph`);
  };
}

function firstParagraphLength(session: YrsSession): number {
  let length = 0;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'pilcrow') break;
    length += segment.kind === 'text' ? segment.text.length : 1;
  }
  return length;
}

function boldAll(session: YrsSession): void {
  const paraId = firstParagraphId(session);
  session.toggleMark({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: firstParagraphLength(session) } }, { type: 'bold' });
}

function resolveEach(session: YrsSession, decision: 'accept' | 'reject'): void {
  const ids = [...new Set(session.listRevisions().map(({ revisionId }) => revisionId))];
  expect(ids).toHaveLength(1);
  for (const revisionId of ids) {
    if (decision === 'accept') session.acceptChange({ revisionId });
    else session.rejectChange({ revisionId });
  }
  expect(session.listRevisions()).toEqual([]);
}

const save = setupStructuralPackageSaves();

async function saveFromParsedDocument(source: Uint8Array, apply: Apply): Promise<Uint8Array> {
  const parsed = await parseDocx(source.slice().buffer, { preloadFonts: false });
  const session = await createYrsSession();
  try {
    documentToYrs(session, parsed);
    apply(session);
    const dirty = new DirtyProjectionStories();
    dirty.add('body');
    const live = session.materializeDocx();
    const base = live ? mergeDocxHostMetadata(live, parsed) : parsed;
    const projected = yrsToDocument(session, base, dirty.projectionOptions());
    return new Uint8Array(await saveEditorDocument(session, projected, []));
  } finally {
    session.destroy();
  }
}

async function saveOnMainThreadWithReport(source: Uint8Array, apply: Apply) {
  const session = await createYrsSession();
  try {
    const host = session.openDocx(source.slice(), true).document;
    apply(session);
    const dirty = new DirtyProjectionStories();
    dirty.add('body');
    const base = session.materializeDocx() as Document;
    const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host), dirty.projectionOptions());
    const bytes = new Uint8Array(await saveEditorDocument(session, projected, []));
    return { bytes, kept: editorSaveKeptUnchanged(session) };
  } finally {
    session.destroy();
  }
}

const reporting: Array<{ destroy(): void }> = [];
let startReportingWorker: (clientId?: number) => InProcessResidentWorker;

beforeAll(async () => {
  startReportingWorker = await residentWorkerFactory();
});

afterEach(() => {
  for (const owned of reporting.splice(0)) owned.destroy();
});

async function saveOnWorkerWithReport(source: Uint8Array, apply: Apply) {
  const worker = startReportingWorker(412);
  const client = new ResidentEngineWorkerClient(worker);
  reporting.push(client);
  const { hostJson } = await client.open(source.slice());
  const host = decodeDocxHostJson(hostJson, source.slice()).document;
  const resident = worker.sessions[0]!;
  const peer = await createYrsSession();
  reporting.push(peer);
  peer.openDocx(source.slice(), false);
  peer.loadState(resident.encodeStateAsUpdate());
  apply(peer);
  resident.applyUpdate(peer.encodeStateAsUpdate());
  const save = spyOn(resident, 'save');
  const saveDocument = spyOn(resident, 'saveDocument');
  try {
    const saved = await client.save({ comments: [], host: hostSaveMetadata(host) });
    const record = save.mock.calls[0]?.[4] ?? saveDocument.mock.calls[0]?.[2];
    return { bytes: new Uint8Array(saved.bytes), kept: record?.keptUnchanged };
  } finally {
    save.mockRestore();
    saveDocument.mockRestore();
  }
}

interface EditorPath {
  name: string;
  keepsSourceBytes: boolean;
  open(source: Uint8Array, apply: Apply): Promise<Uint8Array>;
}

const sourcePath = (name: string, path: SavePath): EditorPath => ({
  name,
  keepsSourceBytes: true,
  open: (source, apply) => save(path, source, (engine) => apply(engine as unknown as YrsSession)),
});

const PATHS: EditorPath[] = [
  sourcePath('Rust loader session', 'main-thread (hybrid and in-thread)'),
  sourcePath('worker engine', 'worker-only'),
  { name: 'TypeScript loader', keepsSourceBytes: false, open: saveFromParsedDocument },
];

const noop: Apply = () => {};

describe.each(PATHS)('$name', (path) => {
  describe.each(SHAPES)('$name at paragraph level', (shape) => {
    const source = () => fixture(firstParagraph(wrapped(shape)));
    const exactSpan = bareSpan(shape);
    const sourceSpan = wrapped(shape);
    const revision = { kind: shape.revision, author: shape.author };
    const listed = (session: YrsSession) => session.listRevisions().map(({ kind, author }) => ({ kind, author }));

    it('holds the reference with its revision on load', async () => {
      let references: ReturnType<typeof noteReferences> = [];
      let revisions: Array<{ kind: string; author: string }> = [];
      await path.open(source(), (session) => {
        references = noteReferences(session);
        revisions = listed(session);
      });
      expect(references).toHaveLength(1);
      expect(Number(references[0]!.payload[`${shape.note}RefId`])).toBe(1);
      expect(references[0]!.attributes[shape.tag]).toMatchObject({ author: shape.author, date: DATE });
      expect(revisions).toContainEqual(revision);
    });

    it('keeps the span on a plain save', async () => {
      const original = source();
      const saved = await path.open(original, noop);
      if (path.keepsSourceBytes) {
        expect(part(saved, 'word/document.xml')).toContain(firstParagraph(wrapped(shape)));
        expect(part(saved, 'word/document.xml')).toContain(SECOND);
        expect(savedNote(saved, shape.note)).toBe(savedNote(original, shape.note));
      } else {
        expect(wrapperSpan(shape, bodyParagraphs(saved)[0] ?? '')).toBe(exactSpan);
      }
    });

    it('keeps the reference inside its wrapper when its paragraph is edited', async () => {
      const saved = await path.open(source(), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(wrapperSpan(shape, first)).toBe(exactSpan);
    });

    it('lists the revision after reopening a save of an edit in its paragraph', async () => {
      const saved = await path.open(source(), (session) => typeInto(session, 0));
      let revisions: Array<{ kind: string; author: string }> = [];
      await path.open(saved, (session) => {
        revisions = listed(session);
      });
      expect(revisions).toContainEqual(revision);
    });

    it('keeps the span when another paragraph is edited', async () => {
      const saved = await path.open(source(), (session) => typeInto(session, 1));
      const [first, second] = bodyParagraphs(saved);
      expect(plainText(second ?? '')).toBe('OtXher');
      if (path.keepsSourceBytes) expect(first).toBe(firstParagraph(wrapped(shape)));
      else expect(wrapperSpan(shape, first ?? '')).toBe(exactSpan);
    });

    for (const decision of ['accept', 'reject'] as const) {
      const removes = (decision === 'accept') === (shape.revision === 'deletion');
      it(`${decision}s the revision, then ${removes ? 'drops the reference and keeps the note body' : 'saves a live reference'}`, async () => {
        const saved = await path.open(source(), (session) => resolveEach(session, decision));
        const first = bodyParagraphs(saved)[0] ?? '';
        expect(plainText(first)).toBe('Head  tail');
        expect(first).not.toMatch(/<w:(?:del|ins)\b/);
        if (removes) expect(first).not.toContain(`<w:${shape.note}Reference`);
        else expect(first).toMatch(liveReferencePattern(shape.note));
        expect(plainText(savedNote(saved, shape.note))).toBe(` ${shape.note} body`);
      });

      it(`${decision}s the revision of a reopened edited save, then saves ${removes ? 'no reference and keeps the note body' : 'a bare live reference'}`, async () => {
        const edited = await path.open(source(), (session) => typeInto(session, 0));
        const saved = await path.open(edited, (session) => resolveEach(session, decision));
        const first = bodyParagraphs(saved)[0] ?? '';
        expect(plainText(first)).toBe('HeXad  tail');
        expect(first).not.toMatch(/<w:(?:del|ins)\b/);
        if (removes) expect(first).not.toContain(`<w:${shape.note}Reference`);
        else expect(savedReferenceRun(first, shape.note)).toBe(bareRun(shape.note, 1));
        expect(plainText(savedNote(saved, shape.note))).toBe(` ${shape.note} body`);
      });
    }

    it('reopens an edited save with its revision id, author and date, one reference and the note body', async () => {
      const saved = await path.open(source(), (session) => typeInto(session, 0));
      let references: ReturnType<typeof noteReferences> = [];
      await path.open(saved, (session) => {
        references = noteReferences(session);
      });
      expect(count(part(saved, 'word/document.xml'), `<w:${shape.note}Reference\\b`)).toBe(1);
      expect(references).toHaveLength(1);
      expect(Number(references[0]!.payload[`${shape.note}RefId`])).toBe(1);
      const mark = (references[0]!.attributes[shape.tag] ?? {}) as Record<string, unknown>;
      expect({ id: String(mark.id), author: mark.author, date: mark.date })
        .toEqual({ id: String(shape.id), author: shape.author, date: DATE });
      expect(plainText(savedNote(saved, shape.note))).toBe(` ${shape.note} body`);
    });

    it('keeps text and the reference of one revision in one wrapper, in order, when its paragraph is edited', async () => {
      const words = shape.tag === 'del' ? deletedText : text;
      const span = revisionSpan(shape, `${words('left')}${referenceRun(shape.note)}${words('right')}`);
      const saved = await path.open(fixture(firstParagraph(span)), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      const run = (value: string) => `<w:r>${RUN_PROPERTIES}<w:(?:t|delText)(?:\\s[^>]*)?>${value}</w:(?:t|delText)></w:r>`;
      expect(count(first, `<w:${shape.tag}\\b`)).toBe(1);
      expect(first).toMatch(new RegExp(
        `${opening(shape)}${run('left')}${bareRun(shape.note, 1)}${run('right')}</w:${shape.tag}>`
      ));
    });

    it('writes an untracked reference beside it bare when its paragraph is edited', async () => {
      const untracked = referenceRun(shape.note, 2);
      const saved = await path.open(fixture(firstParagraph(`${wrapped(shape)}${untracked}`), [1, 2]), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(savedReferenceRun(first, shape.note, 2)).toBe(bareRun(shape.note, 2));
      expect(wrapperSpan(shape, first)).toBe(exactSpan);
    });

    it('saves bold text beside a revision of a reference alone, with the reference inside its wrapper', async () => {
      const original = firstParagraph(wrapped(shape));
      const first = bodyParagraphs(await path.open(fixture(original), boldHead))[0] ?? '';
      if (path.keepsSourceBytes) {
        expect(first).toMatch(BOLD_HEAD);
        expect(count(first, `<w:${shape.tag}\\b`)).toBe(1);
        expect(first).toContain(sourceSpan);
      } else {
        expect(first).toMatch(BOLD_HEAD);
        expect(wrapperSpan(shape, first)).toBe(exactSpan);
      }
    });

    it('saves a bold edit beside one revision that holds text and the reference, with the reference inside its wrapper', async () => {
      const words = shape.tag === 'del' ? deletedText : text;
      const span = revisionSpan(shape, `${words('left')}${referenceRun(shape.note)}`);
      const first = bodyParagraphs(await path.open(fixture(firstParagraph(span)), boldHead))[0] ?? '';
      const left = `<w:r>${RUN_PROPERTIES}<w:(?:t|delText)(?:\\s[^>]*)?>left</w:(?:t|delText)></w:r>`;
      const reference = `<w:r>${RUN_PROPERTIES}<w:${shape.note}Reference w:id="1"/></w:r>`;
      expect(first).toMatch(BOLD_HEAD);
      expect(first).toMatch(new RegExp(`${opening(shape)}${left}${reference}</w:${shape.tag}>`));
    });

    it('saves bold text beside a revision of a reference alone and one with text around a reference, with both wrappers', async () => {
      const { alone, original } = compound(shape);
      const first = bodyParagraphs(await path.open(fixture(original, [1, 2]), boldHead))[0] ?? '';
      const run = (value: string) => `<w:r>${RUN_PROPERTIES}<w:(?:t|delText)(?:\\s[^>]*)?>${value}</w:(?:t|delText)></w:r>`;
      const reference = (id: number) => `<w:r>${RUN_PROPERTIES}<w:${shape.note}Reference w:id="${id}"/></w:r>`;
      if (path.keepsSourceBytes) {
        expect(first).toMatch(BOLD_HEAD);
        expect(count(first, `<w:${shape.tag}\\b`)).toBe(2);
        expect(first).toMatch(new RegExp(`${opening(shape)}${run('left')}${reference(1)}${run('right')}</w:${shape.tag}>`));
        expect(first).toContain(alone);
      } else {
        expect(first).toMatch(BOLD_HEAD);
        expect(first).toMatch(new RegExp(`${opening(shape)}${run('left')}${reference(1)}${run('right')}</w:${shape.tag}>`));
        expect(first).toMatch(new RegExp(`${opening(shape)}${reference(2)}</w:${shape.tag}>`));
      }
    });

    it('lists both revisions after reopening a save of bold text beside them', async () => {
      const saved = await path.open(fixture(compound(shape).original, [1, 2]), boldHead);
      let revisions: Array<{ id: string; kind: string; author: string }> = [];
      await path.open(saved, (session) => {
        revisions = session.listRevisions().map(({ revisionId, kind, author }) => ({ id: revisionId, kind, author }));
      });
      expect([...new Set(revisions.map(({ id }) => id))].sort()).toEqual([String(shape.id), String(shape.id + 10)].sort());
      expect([...new Set(revisions.map(({ kind, author }) => `${kind} ${author}`))]).toEqual([`${shape.revision} ${shape.author}`]);
    });
  });

  describe.each(['footnote', 'endnote'] as const)('untracked %s reference', (note) => {
    it('stays live and unwrapped when its paragraph is edited', async () => {
      const saved = await path.open(fixture(firstParagraph(referenceRun(note))), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(first).toMatch(liveReferencePattern(note));
      expect(first).not.toMatch(/<w:(?:del|ins)\b/);
    });
  });

  describe.each(['footnote', 'endnote'] as const)('%s reference in a field result', (note) => {
    it('stays once in the field result when suggested for deletion', async () => {
      const saved = await path.open(fixture(fieldParagraph(note)), suggestDeletingReference);
      const xml = part(saved, 'word/document.xml');
      const references = [...xml.matchAll(new RegExp(`<w:${note}Reference w:id="1"/>`, 'g'))];
      expect(references).toHaveLength(1);
      const at = references[0]!.index!;
      expect(xml.indexOf('w:fldCharType="separate"')).toBeGreaterThan(-1);
      expect(xml.indexOf('w:fldCharType="separate"')).toBeLessThan(at);
      expect(xml.indexOf('w:fldCharType="end"')).toBeGreaterThan(at);
    });

    it('is saved bare between bare field characters when the result text is edited', async () => {
      const saved = await path.open(fixture(fieldParagraph(note)), (session) => typeAt(session, linkOffset(session) + 2));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('Head LiXnk tail');
      expect({
        begin: fieldCharacterRun(first, 'begin'),
        separate: fieldCharacterRun(first, 'separate'),
        end: fieldCharacterRun(first, 'end'),
        reference: savedReferenceRun(first, note),
      }).toEqual({
        begin: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        separate: '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        end: '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        reference: bareRun(note, 1),
      });
    });

    it('keeps the field of its source paragraph byte for byte when only a deletion of the reference is suggested', async () => {
      const source = await path.open(fixture(fieldParagraph(note)), (session) => typeAt(session, 2));
      const sourceField = fieldSpan(bodyParagraphs(source)[0] ?? '');
      expect(plainText(bodyParagraphs(source)[0] ?? '')).toBe('HeXad Link tail');
      expect(count(sourceField, `<w:${note}Reference w:id="1"/>`)).toBe(1);
      const suggested = await path.open(source, suggestDeletingReference);
      expect(fieldSpan(bodyParagraphs(suggested)[0] ?? '')).toBe(sourceField);
    });
  });

  describe.each(['footnote', 'endnote'] as const)('%s reference formatting on a rewrite', (note) => {
    const deleted = shapeOf(note, 'deletion');
    const accept = (session: YrsSession) => resolveEach(session, 'accept');

    it('leaves the next reference bare after accepting the deletion of a reference with a language', async () => {
      const source = fixture(firstParagraph(`${revisionSpan(deleted, languageRun(note, 1))}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, accept))[0] ?? '';
      expect(count(first, `<w:${note}Reference w:id="1"/>`)).toBe(0);
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
    });

    it('leaves the next reference bare after accepting the deletion of a styled reference', async () => {
      const source = fixture(firstParagraph(`${wrapped(deleted)}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, accept))[0] ?? '';
      expect(count(first, `<w:${note}Reference w:id="1"/>`)).toBe(0);
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
    });

    it('leaves both references bare when other text is edited', async () => {
      const source = fixture(firstParagraph(`${wrapped(deleted)}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, (session) => typeInto(session, 0)))[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
      expect(wrapperSpan(deleted, first)).toBe(bareSpan(deleted));
    });
  });

  describe.each(['footnote', 'endnote'] as const)('%s reference with both marks', (note) => {
    it('is saved inside its insertion only, without the suggested deletion', async () => {
      const inserted = shapeOf(note, 'insertion');
      const saved = await path.open(fixture(firstParagraph(wrapped(inserted))), (session) => {
        suggestDeletingReference(session);
        expect(noteReferences(session)[0]?.attributes.ins).toMatchObject({ author: inserted.author });
      });
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(first).toMatch(new RegExp(
        `${opening(inserted)}<w:r>${RUN_PROPERTIES}<w:${note}Reference w:id="1"/></w:r></w:ins>`
      ));
      expect(first).not.toMatch(/<w:del\b/);
    });
  });
});

describe.each(PATHS.filter(({ keepsSourceBytes }) => keepsSourceBytes))('$name', (path) => {
  describe.each(SHAPES)('$name at paragraph level', (shape) => {
    it('writes the reference of a revision of a reference alone with its source bytes when other text in its paragraph is bolded', async () => {
      const plain = bodyParagraphs(await path.open(fixture(firstParagraph(wrapped(shape))), boldHead))[0] ?? '';
      expect(plain).toContain(wrapped(shape));
      const { alone, original } = compound(shape);
      const both = bodyParagraphs(await path.open(fixture(original, [1, 2]), boldHead))[0] ?? '';
      expect(both).toContain(alone);
    });

    it('keeps the source paragraph when text is bolded beside a revision of a reference alone and an untracked styled reference', async () => {
      const original = firstParagraph(`${wrapped(shape)}${text(' and ')}${referenceRun(shape.note, 2)}`);
      const first = bodyParagraphs(await path.open(fixture(original, [1, 2]), boldHead))[0] ?? '';
      expect(first).toContain(referenceRun(shape.note, 2));
      expect(first).toBe(original);
    });

    it('keeps the source paragraph when text is bolded beside a revision of a reference alone and one that ends with a styled reference', async () => {
      const words = shape.tag === 'del' ? deletedText : text;
      const ending = revisionSpan(shape, `${words('left')}${referenceRun(shape.note)}`);
      const alone = revisionSpan({ ...shape, id: shape.id + 10 }, referenceRun(shape.note, 2));
      const original = firstParagraph(`${ending}${text(' and ')}${alone}`);
      const first = bodyParagraphs(await path.open(fixture(original, [1, 2]), boldHead))[0] ?? '';
      expect(first).toBe(original);
    });
  });
});

const REPORTING = [
  { name: 'Rust loader session', save: saveOnMainThreadWithReport },
  { name: 'worker engine', save: saveOnWorkerWithReport },
];

describe.each(REPORTING)('$name', (route) => {
  describe.each([shapeOf('footnote', 'deletion')])('$name at paragraph level', (shape) => {
    const keptAndReported = async (original: string, noteIds: number[], apply: Apply = boldHead) => {
      const { bytes, kept } = await route.save(fixture(original, noteIds), apply);
      expect(bodyParagraphs(bytes)[0]).toBe(original);
      expect(kept?.structure ?? 0).toBeGreaterThanOrEqual(1);
      expect(kept?.paragraphs).toContainEqual(expect.objectContaining({
        part: 'word/document.xml',
        ordinal: 0,
        scope: 'paragraph',
        dropped: expect.arrayContaining([`w:${shape.tag}`]),
      }));
    };

    it('keeps the source paragraph and reports it when bold text sits beside a revision of a reference alone inside a hyperlink', async () => {
      await keptAndReported(firstParagraph(`<w:hyperlink w:anchor="a">${wrapped(shape)}</w:hyperlink>`), [1]);
    });

    it('keeps the source paragraph and reports it when bold text sits beside two wrappers of one revision that hold a reference alone', async () => {
      const twice = `${wrapped(shape)}${text(' and ')}${revisionSpan(shape, referenceRun(shape.note, 2))}`;
      await keptAndReported(firstParagraph(twice), [1, 2]);
    });

    it('keeps the source paragraph and reports it when bold text sits beside a revision of a reference alone and an untracked styled reference', async () => {
      await keptAndReported(firstParagraph(`${wrapped(shape)}${text(' and ')}${referenceRun(shape.note, 2)}`), [1, 2]);
    });

    it('keeps the source paragraph and reports it when bold text follows a deletion that opens with a reference after a rendered page break', async () => {
      const opening = revisionSpan(shape, `${pageBreakRun(shape.note, 1)}${referenceRun(shape.note, 2)}${deletedText('old')}`);
      const alone = revisionSpan({ ...shape, id: shape.id + 10 }, referenceRun(shape.note, 3));
      const original = paragraph('0D000001', `${opening}${alone}${text('Head ')}${referenceRun(shape.note, 4)}`);
      await keptAndReported(original, [1, 2, 3, 4], boldWord('Head'));
    });

    it('keeps the source paragraph and reports it when bold text sits between a leading styled reference after a rendered page break and a revision of a reference alone', async () => {
      const original = paragraph('0D000001', `${pageBreakRun(shape.note, 1)}${text('Head ')}${revisionSpan(shape, referenceRun(shape.note, 2))}`);
      await keptAndReported(original, [1, 2], boldWord('Head'));
    });

    it('keeps the source paragraph and reports it when bold text sits beside a deletion of a reference after a proofing mark and an untracked styled reference', async () => {
      const deletion = revisionSpan(shape, `<w:proofErr w:type="spellEnd"/>${referenceRun(shape.note, 1)}`);
      await keptAndReported(firstParagraph(`${deletion}${referenceRun(shape.note, 2)}`), [1, 2]);
    });

    it('keeps the source paragraph and reports it when bold text sits beside a deletion of a reference after empty deleted text and an untracked styled reference', async () => {
      const deletion = revisionSpan(shape, `<w:r><w:delText/></w:r>${referenceRun(shape.note, 1)}`);
      await keptAndReported(firstParagraph(`${deletion}${referenceRun(shape.note, 2)}`), [1, 2]);
    });

    it('keeps the source paragraph and reports it when bold text sits beside a revision of a reference alone and a bare reference inside a hyperlink', async () => {
      const linked = `<w:hyperlink w:anchor="a">${text('link')}${bareRun(shape.note, 1)}</w:hyperlink>`;
      await keptAndReported(firstParagraph(`${linked}${revisionSpan(shape, referenceRun(shape.note, 2))}`), [1, 2]);
    });

    it('saves bold over the whole paragraph and writes the reference of a revision of a reference alone as its source run', async () => {
      const { bytes, kept } = await route.save(fixture(firstParagraph(wrapped(shape)), [1]), boldAll);
      const first = bodyParagraphs(bytes)[0] ?? '';
      const bold = (value: string) => `<w:r><w:rPr>${IN_RPR}<w:b/>${IN_RPR}</w:rPr><w:t(?:\\s[^>]*)?>${value}</w:t></w:r>`;
      expect(first).toMatch(new RegExp(`${bold('Head ')}[\\s\\S]*${bold(' tail')}`));
      expect(first).toContain(wrapped(shape));
      expect(kept?.paragraphs ?? []).toEqual([]);
    });
  });
});
