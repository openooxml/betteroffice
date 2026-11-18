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
const style = (note: NoteKind) => (note === 'footnote' ? 'FootnoteReference' : 'EndnoteReference');
const referenceRun = (note: NoteKind) =>
  `<w:r><w:rPr><w:rStyle w:val="${style(note)}"/></w:rPr><w:${note}Reference w:id="1"/></w:r>`;
const wrapped = (shape: Shape) =>
  `<w:${shape.tag} w:id="${shape.id}" w:author="${shape.author}" w:date="${DATE}">${referenceRun(shape.note)}</w:${shape.tag}>`;
const paragraph = (id: string, content: string) => `<w:p w14:paraId="${id}" w14:textId="${id}">${content}</w:p>`;
const firstParagraph = (reference: string) => paragraph('0D000001', `${text('Head ')}${reference}${text(' tail')}`);
const SECOND = paragraph('0D000002', text('Other'));

function notesXml(note: NoteKind): string {
  return `<w:${note}s xmlns:w="${W}">` +
    `<w:${note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${note}>` +
    `<w:${note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${note}>` +
    `<w:${note} w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="${style(note)}"/></w:rPr><w:${note}Ref/></w:r>${text(` ${note} body`)}</w:p></w:${note}>` +
    `</w:${note}s>`;
}

function fixture(first: string): Uint8Array {
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
    ['word/footnotes.xml', toBytes(notesXml('footnote'))],
    ['word/endnotes.xml', toBytes(notesXml('endnote'))],
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

const RPR = '<w:rPr>(?:(?!</w:rPr>)[\\s\\S])*</w:rPr>';
const RUN_PROPERTIES = `(?:${RPR})?`;

function wrapperPattern(shape: Shape): RegExp {
  return new RegExp(
    `<w:${shape.tag} w:id="\\d+" w:author="${shape.author}" w:date="${DATE}"><w:r>${RUN_PROPERTIES}<w:${shape.note}Reference w:id="1"/></w:r></w:${shape.tag}>`
  );
}

function liveReferencePattern(note: NoteKind): RegExp {
  return new RegExp(`<w:r>${RUN_PROPERTIES}<w:${note}Reference w:id="1"/></w:r>`);
}

function wrapperSpan(shape: Shape, xml: string): string {
  return (wrapperPattern(shape).exec(xml)?.[0] ?? '').replace(new RegExp(RPR), '');
}

const plainText = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');

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
    const exactSpan = `<w:${shape.tag} w:id="${shape.id}" w:author="${shape.author}" w:date="${DATE}"><w:r><w:${shape.note}Reference w:id="1"/></w:r></w:${shape.tag}>`;
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
    }
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
});
