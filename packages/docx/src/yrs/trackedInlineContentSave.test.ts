import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsAuthor, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = [
  `xmlns:w="${W}"`,
  `xmlns:r="${R}"`,
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
].join(' ');
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const INS = { tag: 'ins', kind: 'insertion', id: '41', author: 'Grace', date: '2025-03-04T05:06:07Z' } as const;
const DEL = { tag: 'del', kind: 'deletion', id: '42', author: 'Grace', date: '2025-03-04T05:06:07Z' } as const;
const MOVE_FROM = { tag: 'moveFrom', kind: 'deletion', id: '43', author: 'Grace', date: '2025-03-04T05:06:07Z' } as const;
const MOVE_TO = { tag: 'moveTo', kind: 'insertion', id: '44', author: 'Grace', date: '2025-03-04T05:06:07Z' } as const;
type Change = typeof INS | typeof DEL | typeof MOVE_FROM | typeof MOVE_TO;
const SUGGESTER: YrsAuthor = { name: 'Ada', date: '2026-01-01T00:00:00Z' };

const open = (change: Change) => `<w:${change.tag} w:id="${change.id}" w:author="${change.author}" w:date="${change.date}">`;
const tracked = (change: Change, inner: string) => `${open(change)}${inner}</w:${change.tag}>`;
const run = (inner: string) => `<w:r>${inner}</w:r>`;
const t = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
const dt = (value: string) => run(`<w:delText xml:space="preserve">${value}</w:delText>`);
const BREAK = run('<w:br/>');
const COMMENT_REFERENCE = run('<w:commentReference w:id="5"/>');
const SIMPLE_FIELD = `<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple>`;
const MATH = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>';
const COMMENTED = '<w:commentRangeStart w:id="5"/><w:r><w:t>noted</w:t></w:r><w:commentRangeEnd w:id="5"/>';

const textFor = (change: Change) => (change.kind === 'deletion' ? dt : t);
const paragraph = (inner: string) => `<w:p w14:paraId="1A2B3C01" w14:textId="1A2B3C01">${t('Head')}${inner}${t(' tail')}</w:p>`;
const documentXml = (body: string) => `<w:document ${NAMESPACES}><w:body>${body}${SECTION}</w:body></w:document>`;

function docx(body: string): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE_DOC}.wordprocessingml.comments+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/></Relationships>`)],
    ['word/comments.xml', toBytes(`<w:comments xmlns:w="${W}"><w:comment w:id="5" w:author="Lin" w:date="2025-03-04T05:06:07Z" w:initials="L"><w:p><w:r><w:t>note body</w:t></w:r></w:p></w:comment></w:comments>`)],
    ['word/document.xml', toBytes(documentXml(body))],
  ])));
}

const part = (bytes: Uint8Array, name: string) => new TextDecoder().decode(unzipContainer(bytes)[name]);
const count = (xml: string, pattern: RegExp) => xml.match(pattern)?.length ?? 0;
const EMPTY_RUN = /<w:r\b[^>]*\/>|<w:r\b[^>]*><\/w:r>/g;

const TOKEN = /<w:(t|delText)\b[^>]*>([^<]*)<\/w:\1>|<w:br\b[^>]*\/>|<w:commentReference\b[^>]*w:id="(\d+)"[^>]*\/>/g;

/** Text and deleted text verbatim; breaks and comment references as bracketed markers. */
function render(xml: string): string {
  let out = '';
  for (const [, , text, reference] of xml.matchAll(TOKEN)) {
    if (text !== undefined) out += text;
    else if (reference !== undefined) out += `[ref ${reference}]`;
    else out += '[br]';
  }
  return out;
}

const body = (xml: string) => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'));

/** Every saved wrapper that carries the change's id, author and date, with its rendered and raw content. */
function trackedContent(xml: string, change: Change): { wrappers: number; content: string; raw: string } {
  const pattern = new RegExp(`<w:${change.tag}\\b([^>]*[^/])>([\\s\\S]*?)</w:${change.tag}>`, 'g');
  const matching = [...xml.matchAll(pattern)].filter(([, attributes]) =>
    attributes!.includes(`w:id="${change.id}"`) &&
    attributes!.includes(`w:author="${change.author}"`) &&
    attributes!.includes(`w:date="${change.date}"`));
  const raw = matching.map(([, , inner]) => inner).join('');
  return { wrappers: matching.length, content: render(raw), raw };
}

function expectCommentResolves(saved: Uint8Array): void {
  const xml = part(saved, 'word/document.xml');
  const range = xml.match(/<w:commentRangeStart\b[^>]*w:id="5"[^>]*\/>([\s\S]*?)<w:commentRangeEnd\b[^>]*w:id="5"[^>]*\/>/);
  expect(range).not.toBeNull();
  expect(render(range![1]!)).toBe('noted');
  expect(count(xml, /<w:commentRangeStart\b/g)).toBe(1);
  expect(count(xml, /<w:commentRangeEnd\b/g)).toBe(1);
  expect(part(saved, 'word/comments.xml')).toMatch(/<w:comment\b[^>]*w:id="5"[^>]*>[\s\S]*note body[\s\S]*<\/w:comment>/);
}

const sessions: YrsSession[] = [];
let clientId = 77300;

async function opened(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

async function save(session: YrsSession): Promise<Uint8Array> {
  return (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
}

async function typedAndSaved(bytes: Uint8Array, typed: string, suggesting?: YrsAuthor): Promise<Uint8Array> {
  const session = await opened(bytes);
  const { paraId } = session.paragraphs('body')[0]!;
  session.insertText({ story: 'body', paraId, offset: 0 }, typed, suggesting);
  return save(session);
}

async function resolvedAndSaved(bytes: Uint8Array, change: Change, resolution: 'accept' | 'reject'): Promise<Uint8Array> {
  const session = await opened(bytes);
  const revision = session.listRevisions().find(entry => entry.kind === change.kind && entry.author === change.author);
  expect(revision).toBeDefined();
  if (resolution === 'accept') session.acceptChange({ revisionId: revision!.revisionId });
  else session.rejectChange({ revisionId: revision!.revisionId });
  return save(session);
}

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { sessions.splice(0).forEach(session => session.destroy()); });

const MODES: Array<[string, YrsAuthor | undefined]> = [
  ['after typing elsewhere in the paragraph', undefined],
  ['after a suggested insertion elsewhere in the paragraph', SUGGESTER],
];

const CASES: Array<{ name: string; change: Change; middle: string; content: string; pattern: RegExp; commented: boolean }> = [
  { name: 'a line break inside a tracked insertion stays inside it', change: INS, middle: BREAK, content: 'left[br]right', pattern: /<w:br\b/g, commented: false },
  { name: 'a line break inside a tracked deletion stays inside it', change: DEL, middle: BREAK, content: 'left[br]right', pattern: /<w:br\b/g, commented: false },
  { name: 'a comment reference inside a tracked insertion stays inside it', change: INS, middle: COMMENT_REFERENCE, content: 'left[ref 5]right', pattern: /<w:commentReference\b/g, commented: true },
  { name: 'a comment reference inside a tracked deletion stays inside it', change: DEL, middle: COMMENT_REFERENCE, content: 'left[ref 5]right', pattern: /<w:commentReference\b/g, commented: true },
];

describe('a line break and a comment reference inside a tracked change', () => {
  for (const { name, change, middle, content, pattern, commented } of CASES) {
    for (const [mode, suggesting] of MODES) {
      it(`${name} ${mode}`, async () => {
        const text = textFor(change);
        const source = paragraph(`${commented ? COMMENTED : ''}${tracked(change, `${text('left')}${middle}${text('right')}`)}`);
        const saved = await typedAndSaved(docx(source), 'X', suggesting);
        const xml = part(saved, 'word/document.xml');
        const wrapper = trackedContent(xml, change);
        expect(wrapper.wrappers).toBeGreaterThan(0);
        expect(wrapper.content).toBe(content);
        expect(render(body(xml))).toBe(`X${render(source)}`);
        expect(count(xml, pattern)).toBe(1);
        if (change === DEL) {
          expect(count(wrapper.raw, /<w:t\b/g)).toBe(0);
          expect(count(wrapper.raw, /<w:delText\b/g)).toBe(2);
        }
        if (commented) expectCommentResolves(saved);
      });
    }
  }

  const REFERENCE_BESIDE_RANGE = '<w:commentRangeEnd w:id="5"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="5"/></w:r>';

  function expectReferenceBesideRange(saved: Uint8Array): void {
    const xml = part(saved, 'word/document.xml');
    expect(count(xml, /<w:commentReference\b/g)).toBe(1);
    expect(xml.split(REFERENCE_BESIDE_RANGE)).toHaveLength(2);
    expect(count(xml, /<w:(ins|del)\b/g)).toBe(0);
    expect(render(body(xml))).toBe('Headnoted[ref 5] tail');
    expectCommentResolves(saved);
  }

  it('accepting a tracked deletion that holds a comment reference moves the reference beside the comment\'s range and keeps the comment', async () => {
    const source = paragraph(`${COMMENTED}${tracked(DEL, `${dt('left')}${COMMENT_REFERENCE}${dt('right')}`)}`);
    expectReferenceBesideRange(await resolvedAndSaved(docx(source), DEL, 'accept'));
  });

  it('rejecting a tracked insertion that holds a comment reference moves the reference beside the comment\'s range and keeps the comment', async () => {
    const source = paragraph(`${COMMENTED}${tracked(INS, `${t('left')}${COMMENT_REFERENCE}${t('right')}`)}`);
    expectReferenceBesideRange(await resolvedAndSaved(docx(source), INS, 'reject'));
  });

  it('saving without an edit writes the source document part unchanged', async () => {
    const bytes = docx(paragraph([
      COMMENTED,
      tracked(INS, `${t('left')}${BREAK}${t('right')}`),
      tracked(DEL, `${dt('left')}${COMMENT_REFERENCE}${dt('right')}`),
    ].join('')));
    const session = await opened(bytes);
    expect(part(await save(session), 'word/document.xml')).toBe(part(bytes, 'word/document.xml'));
  });

  it('an untracked line break next to a tracked insertion stays in place after typing elsewhere in the paragraph', async () => {
    const source = paragraph(`${BREAK}${t('mid')}${tracked(INS, t('added'))}`);
    const xml = part(await typedAndSaved(docx(source), 'X'), 'word/document.xml');
    expect(render(body(xml))).toBe('XHead[br]midadded tail');
    expect(count(xml, /<w:br\b/g)).toBe(1);
    expect(trackedContent(xml, INS).content).toBe('added');
  });

  it('a second edit after reopening a saved document keeps the tracked line break and comment reference', async () => {
    const source = paragraph([
      COMMENTED,
      tracked(INS, `${t('a')}${BREAK}${t('b')}`),
      t(' mid '),
      tracked(DEL, `${dt('c')}${COMMENT_REFERENCE}${dt('d')}`),
    ].join(''));
    const first = await typedAndSaved(docx(source), 'X');
    const firstXml = part(first, 'word/document.xml');
    expect(count(firstXml, EMPTY_RUN)).toBe(0);
    expect(trackedContent(firstXml, INS).content).toBe('a[br]b');
    expect(trackedContent(firstXml, DEL).content).toBe('c[ref 5]d');
    expectCommentResolves(first);

    const second = await typedAndSaved(first, 'Y');
    const secondXml = part(second, 'word/document.xml');
    expect(count(secondXml, EMPTY_RUN)).toBe(0);
    expect(trackedContent(secondXml, INS).content).toBe(trackedContent(firstXml, INS).content);
    expect(trackedContent(secondXml, DEL).content).toBe(trackedContent(firstXml, DEL).content);
    expect(render(body(secondXml))).toBe(`Y${render(body(firstXml))}`);
    const patterns = [/<w:br\b/g, /<w:commentReference\b/g, /<w:ins\b[^>]*[^/]>/g, /<w:del\b[^>]*[^/]>/g];
    const counts = (xml: string) => patterns.map(pattern => [String(pattern), count(xml, pattern)]);
    expect(counts(secondXml)).toEqual(counts(firstXml));
    expectCommentResolves(second);
  });

  for (const [change, tag, side] of [[MOVE_FROM, 'del', 'deletion'], [MOVE_TO, 'ins', 'insertion']] as const) {
    it(`a comment reference inside a tracked ${change.tag} is saved inside the same ${side} as its text`, async () => {
      const text = textFor(change);
      const source = paragraph(`${COMMENTED}${tracked(change, `${text('left')}${COMMENT_REFERENCE}${text('right')}`)}`);
      const saved = await typedAndSaved(docx(source), 'X');
      const xml = part(saved, 'word/document.xml');
      const wrappers = [...xml.matchAll(new RegExp(`<w:${tag}\\b([^>]*[^/])>([\\s\\S]*?)</w:${tag}>`, 'g'))]
        .map(([, attributes, inner]) => ({
          revision: ['id', 'author', 'date'].map(name => attributes!.match(new RegExp(`w:${name}="([^"]*)"`))?.[1]),
          inner: inner!,
        }));
      expect(wrappers.length).toBeGreaterThan(0);
      expect(wrappers[0]!.revision.every(value => value !== undefined)).toBe(true);
      for (const { revision } of wrappers) expect(revision).toEqual(wrappers[0]!.revision);
      const inner = wrappers.map(wrapper => wrapper.inner).join('');
      expect(render(inner)).toBe('left[ref 5]right');
      expect(count(inner, /<w:commentReference\b/g)).toBe(1);
      expect(count(xml, /<w:commentReference\b/g)).toBe(1);
      expect(count(xml, /<w:(moveFrom|moveTo)\b/g)).toBe(0);
      expect(render(body(xml))).toBe(`X${render(source)}`);
      expectCommentResolves(saved);
    });
  }

  it('both seeders give a comment reference inside a tracked change the same marks as the text around it', async () => {
    for (const change of [INS, DEL, MOVE_FROM, MOVE_TO]) {
      const text = textFor(change);
      const bytes = docx(paragraph(`${COMMENTED}${tracked(change, `${text('left')}${COMMENT_REFERENCE}${text('right')}`)}`));
      const engine = await createYrsSession({ clientId: (clientId += 1) });
      const projected = await createYrsSession({ clientId: (clientId += 1) });
      sessions.push(engine, projected);
      engine.seedFromDocx(bytes);
      documentToYrs(projected, await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false }));
      const key = change.kind === 'insertion' ? 'ins' : 'del';
      const reference = (session: YrsSession) => {
        const segments = session.storySegments('body');
        const found = segments.filter(segment =>
          segment.kind === 'embed' && segment.payload.modelKind === 'commentReference');
        expect(found).toHaveLength(1);
        const around = segments.filter(segment => segment.kind === 'text' && ['left', 'right'].includes(segment.text));
        expect(around).toHaveLength(2);
        for (const segment of around) expect(segment.attributes[key]).toEqual(found[0]!.attributes[key]);
        return found[0]!.attributes;
      };
      const attributes = reference(projected);
      expect(Object.keys(attributes)).toEqual([key]);
      const revision = attributes[key] as Record<string, unknown>;
      expect(String(revision.id)).toBe(change.id);
      expect(revision.author).toBe(change.author);
      expect(revision.date).toBe(change.date);
      expect('isMovePair' in revision).toBe(false);
      expect(reference(engine)).toEqual(attributes);
    }
  });

  it('a field and an equation inside a tracked insertion are not yet kept after typing elsewhere in the paragraph', async () => {
    const source = paragraph(tracked(INS, `${t('left')}${SIMPLE_FIELD}${MATH}${t('right')}`));
    const xml = part(await typedAndSaved(docx(source), 'X'), 'word/document.xml');
    expect(render(body(xml))).toBe('XHeadleftright tail');
    expect(trackedContent(xml, INS).content).toBe('leftright');
    expect(count(xml, /<w:fldSimple\b/g)).toBe(0);
    expect(count(xml, /<m:oMath\b/g)).toBe(0);
  });
});
