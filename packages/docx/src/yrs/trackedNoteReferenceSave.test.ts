import { describe, expect, it } from 'bun:test';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { DirtyProjectionStories, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, type YrsSession } from './index';
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
const INSERTED_MARK = `<w:ins w:id="21" w:author="Ada" w:date="${DATE}">${text('!')}</w:ins>`;
const withInsertedMark = (content: string) => firstParagraph(`${content}${INSERTED_MARK}`);

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
const styled = (note: NoteKind) => `<w:rPr>${IN_RPR}<w:rStyle w:val="${style(note)}"/>${IN_RPR}</w:rPr>`;
const styledReference = (note: NoteKind) => `<w:r>${styled(note)}<w:${note}Reference w:id="1"/></w:r>`;
const opening = (shape: Shape) => `<w:${shape.tag} w:id="\\d+" w:author="${shape.author}" w:date="${DATE}">`;

function wrapperPattern(shape: Shape): RegExp {
  return new RegExp(`${opening(shape)}${styledReference(shape.note)}</w:${shape.tag}>`);
}

function liveReferencePattern(note: NoteKind): RegExp {
  return new RegExp(`<w:r>${RUN_PROPERTIES}<w:${note}Reference w:id="1"/></w:r>`);
}

function styledReferencePattern(note: NoteKind): RegExp {
  return new RegExp(styledReference(note));
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

function suggestTyping(session: YrsSession, offset: number): void {
  session.insertText({ story: 'body', paraId: firstParagraphId(session), offset }, 'X', SUGGESTER);
}

function boldHead(session: YrsSession): void {
  const paraId = firstParagraphId(session);
  session.toggleMark({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }, { type: 'bold' });
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
    const exactSpan = wrapped(shape);
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
      expect(first).toMatch(wrapperPattern(shape));
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

      it(`${decision}s the revision of a reopened edited save, then saves ${removes ? 'no reference and keeps the note body' : 'a live reference with its style'}`, async () => {
        const edited = await path.open(source(), (session) => typeInto(session, 0));
        const saved = await path.open(edited, (session) => resolveEach(session, decision));
        const first = bodyParagraphs(saved)[0] ?? '';
        expect(plainText(first)).toBe('HeXad  tail');
        expect(first).not.toMatch(/<w:(?:del|ins)\b/);
        if (removes) expect(first).not.toContain(`<w:${shape.note}Reference`);
        else expect(first).toMatch(styledReferencePattern(shape.note));
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
        `${opening(shape)}${run('left')}${styledReference(shape.note)}${run('right')}</w:${shape.tag}>`
      ));
    });

    it('keeps an untracked reference beside it byte for byte when its paragraph is edited', async () => {
      const untracked = referenceRun(shape.note, 2);
      const saved = await path.open(fixture(firstParagraph(`${wrapped(shape)}${untracked}`), [1, 2]), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(first).toContain(untracked);
      expect(first).toMatch(wrapperPattern(shape));
    });
  });

  describe.each(['footnote', 'endnote'] as const)('untracked %s reference', (note) => {
    it('stays live, unwrapped and styled when its paragraph is edited', async () => {
      const saved = await path.open(fixture(firstParagraph(referenceRun(note))), (session) => typeInto(session, 0));
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(first).toMatch(styledReferencePattern(note));
      expect(first).not.toMatch(/<w:(?:del|ins)\b/);
    });

    it('keeps its run byte for byte, and a deleted reference beside it, when text in its paragraph is bolded', async () => {
      const deleted = shapeOf(note, 'deletion');
      const untracked = referenceRun(note, 2);
      const saved = await path.open(fixture(firstParagraph(`${wrapped(deleted)}${untracked}`), [1, 2]), boldHead);
      const first = bodyParagraphs(saved)[0] ?? '';
      expect(first).toContain(untracked);
      expect(first).toMatch(wrapperPattern(deleted));
      expect(first).toMatch(new RegExp(`<w:r><w:rPr>${IN_RPR}<w:b/>${IN_RPR}</w:rPr><w:t(?:\\s[^>]*)?>Head</w:t></w:r>`));
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
      const saved = await path.open(fixture(fieldParagraph(note)), (session) => suggestTyping(session, linkOffset(session) + 2));
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

    it('keeps the field byte for byte when only a deletion of the reference is suggested', async () => {
      const typed = await path.open(fixture(fieldParagraph(note)), (session) => suggestTyping(session, 2));
      const suggested = await path.open(fixture(fieldParagraph(note)), suggestDeletingReference);
      const typedField = fieldSpan(bodyParagraphs(typed)[0] ?? '');
      expect(plainText(bodyParagraphs(typed)[0] ?? '')).toBe('HeXad Link tail');
      expect(typedField).not.toBe('');
      expect(fieldSpan(bodyParagraphs(suggested)[0] ?? '')).toBe(typedField);
    });
  });

  describe.each(['footnote', 'endnote'] as const)('%s reference formatting on a rewrite', (note) => {
    const deleted = shapeOf(note, 'deletion');
    const accept = (session: YrsSession) => resolveEach(session, 'accept');

    it('leaves the next reference bare after accepting the deletion of a reference with a language', async () => {
      const source = fixture(firstParagraph(`${revisionSpan(deleted, languageRun(note, 1))}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, accept))[0] ?? '';
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
    });

    it('leaves the next reference bare after accepting the deletion of a styled reference', async () => {
      const source = fixture(firstParagraph(`${wrapped(deleted)}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, accept))[0] ?? '';
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
    });

    it('writes a reference whose own run has only a language bare', async () => {
      const source = fixture(withInsertedMark(languageRun(note, 1)));
      const first = bodyParagraphs(await path.open(source, (session) => typeInto(session, 0)))[0] ?? '';
      expect(plainText(first)).toBe('HeXad ! tail');
      expect(savedReferenceRun(first, note)).toBe(bareRun(note, 1));
    });

    it('writes a reference whose own run has a style and a language with its style only', async () => {
      const own = languageRun(note, 1, `<w:rStyle w:val="${style(note)}"/>`);
      const source = fixture(withInsertedMark(own));
      const first = bodyParagraphs(await path.open(source, (session) => typeInto(session, 0)))[0] ?? '';
      expect(plainText(first)).toBe('HeXad ! tail');
      expect(savedReferenceRun(first, note)).toBe(referenceRun(note));
    });

    it('leaves a bare reference bare and a deleted styled reference styled when other text is edited', async () => {
      const source = fixture(firstParagraph(`${wrapped(deleted)}${bareRun(note, 2)}`), [1, 2]);
      const first = bodyParagraphs(await path.open(source, (session) => typeInto(session, 0)))[0] ?? '';
      expect(plainText(first)).toBe('HeXad  tail');
      expect(savedReferenceRun(first, note, 2)).toBe(bareRun(note, 2));
      expect(first).toMatch(wrapperPattern(deleted));
    });

    it('does not write a property of its own run that repeats its style', async () => {
      const own = `<w:r><w:rPr><w:rStyle w:val="${style(note)}"/><w:vertAlign w:val="superscript"/></w:rPr>` +
        `<w:${note}Reference w:id="1"/></w:r>`;
      const source = fixture(withInsertedMark(own));
      const first = bodyParagraphs(await path.open(source, (session) => typeInto(session, 0)))[0] ?? '';
      expect(plainText(first)).toBe('HeXad ! tail');
      expect(savedReferenceRun(first, note)).toBe(referenceRun(note));
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
