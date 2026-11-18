import { afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Comment, Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  DirtyProjectionStories,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
  type EditorSaveRecord,
} from './editorSave';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

const TIMEOUT = 60_000;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const COMMENT_PARTS = [
  'word/comments.xml',
  'word/commentsExtended.xml',
  'word/commentsIds.xml',
  'word/commentsExtensible.xml',
];
const CHANGEABLE = new Set(['word/document.xml', ...COMMENT_PARTS]);
const PARENTS: Record<number, number> = { 3: 2, 4: 2 };
const ALL = [1, 2, 3, 4, 5];

const RPR =
  '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:noProof/><w:webHidden/>' +
  '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/></w:rPr>';
const KEPT = ['<w:noProof/>', '<w:webHidden/>', '<w:szCs w:val="22"/>', '<w:rFonts w:ascii="Arial"'];
const run = (text: string) => `<w:r>${RPR}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const reference = (id: number) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;
const para = (n: number, inner: string) => `<w:p w14:paraId="0A00000${n}" w14:textId="0A00000${n}">${inner}</w:p>`;
const paraId = (id: number) => `0D00000${id}`;
const durableId = (id: number) => `1E00000${id}`;

const bookmark = (twice: boolean, name: string, inner: string) =>
  twice ? `<w:bookmarkStart w:id="7" w:name="${name}"/>${inner}<w:bookmarkEnd w:id="7"/>` : inner;

function body(ids: number[], twice = false): string {
  const has = (id: number) => ids.includes(id);
  const ranged = (id: number, inner: string) =>
    has(id) ? `<w:commentRangeStart w:id="${id}"/>${inner}<w:commentRangeEnd w:id="${id}"/>${reference(id)}` : inner;
  const thread = [2, 3, 4].filter(has);
  const field = `<w:fldSimple w:instr=" PAGE ">${run('1')}</w:fldSimple>`;
  const cell =
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(3, bookmark(twice, 'Cell', run('Cell paragraph')))}</w:tc></w:tr></w:tbl>`;
  return (
    para(1, `${run('Opening page ')}${field}${ranged(1, run('first note'))}${run(' after')}`) +
    para(2, bookmark(twice, 'Second', run('Second paragraph stays'))) +
    cell +
    para(
      4,
      run('Thread ') +
        thread.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join('') +
        run('discussed words') +
        thread.map((id) => `<w:commentRangeEnd w:id="${id}"/>${reference(id)}`).join('') +
        run(' tail')
    ) +
    para(5, `${run('Plain ')}${ranged(5, run('fifth note'))}${run(' end')}`)
  );
}

function built(ids: number[] = ALL, companions = true, twice = false): Uint8Array {
  const extra = companions
    ? [
        ['commentsExtended', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'],
        ['commentsIds', 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds'],
        ['commentsExtensible', 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible'],
      ]
    : [];
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` +
      extra.map(([name]) => `<Override PartName="/word/${name}.xml" ContentType="${OFFICE}.${name}+xml"/>`).join('') +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` +
      extra.map(([name, type], at) => `<Relationship Id="rIdC${at}" Type="${type}" Target="${name}.xml"/>`).join('') +
      '</Relationships>'
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${REL}" xmlns:w14="${W14}"><w:body>${body(ids, twice)}<w:sectPr/></w:body></w:document>`
    )],
    ['word/comments.xml', toBytes(
      `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
      ids.map((id) =>
        `<w:comment w:id="${id}" w:author="Grace" w:date="2025-03-04T05:06:07Z" w:initials="G">` +
        `<w:p w14:paraId="${paraId(id)}" w14:textId="77777777"><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`
      ).join('') +
      '</w:comments>'
    )],
  ]);
  if (companions) {
    parts.set('word/commentsExtended.xml', toBytes(
      `<w15:commentsEx xmlns:w15="${W15}">` +
      ids.map((id) =>
        `<w15:commentEx w15:paraId="${paraId(id)}"${PARENTS[id] ? ` w15:paraIdParent="${paraId(PARENTS[id]!)}"` : ''} w15:done="0"/>`
      ).join('') +
      '</w15:commentsEx>'
    ));
    parts.set('word/commentsIds.xml', toBytes(
      `<w16cid:commentsIds xmlns:w16cid="${W16CID}">` +
      ids.map((id) => `<w16cid:commentId w16cid:paraId="${paraId(id)}" w16cid:durableId="${durableId(id)}"/>`).join('') +
      '</w16cid:commentsIds>'
    ));
    parts.set('word/commentsExtensible.xml', toBytes(
      `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">` +
      ids.map((id) => `<w16cex:commentExtensible w16cex:durableId="${durableId(id)}" w16cex:dateUtc="2025-03-04T05:06:07Z"/>`).join('') +
      '</w16cex:commentsExtensible>'
    ));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const decoder = new TextDecoder();
const part = (bytes: Uint8Array, name: string) => {
  const data = unzipContainer(bytes)[name];
  return data ? decoder.decode(data) : '';
};
const paragraphs = (bytes: Uint8Array) =>
  part(bytes, 'word/document.xml').match(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g) ?? [];

/** Indices of the body paragraphs whose bytes differ from the source's. */
function differing(source: Uint8Array, saved: Uint8Array): number[] {
  const before = paragraphs(source);
  const after = paragraphs(saved);
  expect(after.length).toBe(before.length);
  return before.flatMap((xml, index) => (xml === after[index] ? [] : [index]));
}

/** Source parts outside `allowed` that the save changed, added or dropped. */
function changedParts(source: Uint8Array, saved: Uint8Array, allowed: ReadonlySet<string> = CHANGEABLE): string[] {
  const a = unzipContainer(source);
  const b = unzipContainer(saved);
  return [...new Set([...Object.keys(a), ...Object.keys(b)])]
    .filter((name) => !allowed.has(name))
    .filter((name) => {
      const x = a[name];
      const y = b[name];
      return !x || !y || x.length !== y.length || x.some((byte, at) => byte !== y[at]);
    })
    .sort();
}

/** The run formatting each paragraph holds, as counts of `KEPT`. */
const formatting = (bytes: Uint8Array) =>
  paragraphs(bytes).map((xml) => KEPT.map((needle) => xml.split(needle).length - 1).join('/'));

/** The run formatting of the paragraphs outside `touched`, keyed by their index. */
const untouchedFormatting = (bytes: Uint8Array, touched: number[]) =>
  Object.fromEntries(
    formatting(bytes).flatMap((counts, at) => (touched.includes(at) ? [] : [[`untouched paragraph ${at}`, counts]]))
  );

/** The text of body paragraph `at`. */
const textAt = (bytes: Uint8Array, at: number) => (paragraphs(bytes)[at] ?? '').replace(/<[^>]+>/g, '');

/** The comment's markers in the body, in document order. */
const marks = (bytes: Uint8Array, id: number) =>
  [...part(bytes, 'word/document.xml').matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"\/>/g)]
    .filter((match) => Number(match[2]) === id)
    .map((match) => match[1]!);

/** The markers of each comment in `ids`, keyed by id. */
const marksOf = (bytes: Uint8Array, ids: number[]) => Object.fromEntries(ids.map((id) => [id, marks(bytes, id)]));

/** The body's bookmark ranges with id 7, each from its start to its end, byte for byte. */
const bookmarkPairs = (bytes: Uint8Array) =>
  part(bytes, 'word/document.xml').match(/<w:bookmarkStart w:id="7"[\s\S]*?<w:bookmarkEnd w:id="7"\/>/g) ?? [];

const commentIds = (bytes: Uint8Array) =>
  [...part(bytes, 'word/comments.xml').matchAll(/<w:comment\b[^>]*?\bw:id="(\d+)"/g)].map((match) => Number(match[1]));

/** Comment ids whose entries the companion parts still hold. */
function companionIds(bytes: Uint8Array): Record<string, number[]> {
  const held = (name: string, pattern: RegExp) =>
    [...part(bytes, name).matchAll(pattern)].map((match) => Number(match[1]!.slice(-1))).sort();
  return {
    extended: held('word/commentsExtended.xml', /w15:paraId="(0D0000\d\d)"/g),
    ids: held('word/commentsIds.xml', /w16cid:paraId="(0D0000\d\d)"/g),
    extensible: held('word/commentsExtensible.xml', /w16cex:durableId="(1E0000\d\d)"/g),
  };
}

/** Reply parents commentsExtended names that it holds no entry for. */
function orphanReplies(bytes: Uint8Array): string[] {
  const xml = part(bytes, 'word/commentsExtended.xml');
  const present = new Set([...xml.matchAll(/w15:paraId="([0-9A-F]+)"/g)].map((match) => match[1]));
  return [...xml.matchAll(/w15:paraIdParent="([0-9A-F]+)"/g)].map((match) => match[1]!).filter((id) => !present.has(id));
}

function commentsPart(bytes: Uint8Array) {
  const parts = unzipContainer(bytes);
  const state = !parts['word/comments.xml'] ? 'removed' : commentIds(bytes).length ? 'present' : 'emptied';
  return {
    state,
    relationship: part(bytes, 'word/_rels/document.xml.rels').includes('Target="comments.xml"'),
    contentType: part(bytes, '[Content_Types].xml').includes('PartName="/word/comments.xml"'),
  };
}

const owned: Array<{ destroy(): void }> = [];
let startWorker: (clientId?: number) => InProcessResidentWorker;
let nextClient = 72_400;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});

afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

async function session(): Promise<YrsSession> {
  const created = await createYrsSession({ clientId: nextClient++ });
  owned.push(created);
  return created;
}

/** The text a comment's first anchor covers. */
function anchoredText(reader: YrsSession, id: string): string {
  const [anchor] = reader.resolveComment(id);
  if (!anchor) return '';
  let offset = 0;
  let text = '';
  for (const segment of reader.storySegments(anchor.story)) {
    if (segment.kind === 'text') {
      text += segment.text.slice(Math.max(0, anchor.start - offset), Math.max(0, anchor.end - offset));
      offset += segment.text.length;
    } else offset += 1;
  }
  return text;
}

async function reopened(bytes: Uint8Array) {
  const reader = await session();
  reader.openDocx(bytes.slice(), true);
  const comments = reader.listComments();
  return {
    ids: comments.map((comment) => Number(comment.id)).sort(),
    orphans: comments.filter((comment) => comment.parentId != null && !comments.some((other) => other.id === comment.parentId)).length,
    anchored: (id: number) => anchoredText(reader, String(id)),
  };
}

interface Edit {
  /** The comment the editor deletes: its replies leave the host's list, the session removes it. */
  remove?: number;
  /** Text typed at the end of the paragraph with this id. */
  type?: string;
  /** The host's comment list as the save receives it, from the list the host loaded. */
  comments?: (list: Comment[]) => Comment[];
  /** A change to the host's metadata before the save. */
  host?: (host: Document) => void;
}

interface Saved {
  bytes: Uint8Array;
  full: boolean | undefined;
}

const TYPED = ' typed';

function apply(engine: YrsSession, edit: Edit): boolean {
  if (edit.remove !== undefined) engine.applyRawOps('body', [{ op: 'removeComment', id: String(edit.remove) }]);
  if (edit.type !== undefined) {
    const result = engine.applyEdits({
      expectVersion: engine.version(),
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: edit.type }, at: 'end', text: TYPED }],
    });
    if (!result.ok) throw new Error(`insertText: ${JSON.stringify(result).slice(0, 200)}`);
  }
  return edit.remove !== undefined || edit.type !== undefined;
}

function hostList(host: Document, edit: Edit): Comment[] {
  const list = host.package.document.comments ?? [];
  if (edit.comments) return edit.comments(list);
  const id = edit.remove;
  return id === undefined ? list : list.filter((comment) => comment.id !== id && comment.parentId !== id);
}

async function mainThreadSave(source: Uint8Array, edit: Edit): Promise<Saved> {
  const editor = await session();
  const host = editor.openDocx(source.slice(), true).document;
  const dirty = new DirtyProjectionStories();
  if (apply(editor, edit)) dirty.add('body');
  edit.host?.(host);
  const projected = yrsToDocument(editor, mergeDocxHostMetadata(editor.materializeDocx() as Document, host), dirty.projectionOptions());
  const record: EditorSaveRecord = { full: false };
  const bytes = new Uint8Array(await saveEditorDocument(editor, projected, hostList(host, edit), record));
  return { bytes, full: record.full };
}

async function workerSave(source: Uint8Array, edit: Edit): Promise<Saved> {
  const worker = startWorker(nextClient++);
  const client = new ResidentEngineWorkerClient(worker);
  owned.push(client);
  const { hostJson } = await client.open(source.slice());
  const host = decodeDocxHostJson(hostJson, source.slice()).document;
  const resident = worker.sessions[0]!;
  const peer = await session();
  peer.openDocx(source.slice(), false);
  peer.loadState(resident.encodeStateAsUpdate());
  const edited = apply(peer, edit);
  resident.applyUpdate(peer.encodeStateAsUpdate());
  edit.host?.(host);
  const save = spyOn(resident, 'save');
  const saveDocument = spyOn(resident, 'saveDocument');
  try {
    const saved = await client.save({
      comments: hostList(host, edit),
      host: hostSaveMetadata(host),
      ...(edited ? { stories: ['body'] } : {}),
    });
    const record = save.mock.calls[0]?.[4] ?? saveDocument.mock.calls[0]?.[2];
    return { bytes: new Uint8Array(saved.bytes), full: record?.full };
  } finally {
    save.mockRestore();
    saveDocument.mockRestore();
  }
}

/** The agents engine's own delete of `id`, as its comment deletion saves it. */
async function agentsDelete(source: Uint8Array, id: number): Promise<Uint8Array> {
  const engine = await session();
  engine.seedFromDocx(source.slice());
  const previous = engine.listComments();
  engine.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
  const draft = await session();
  draft.openDocx(new Uint8Array(engine.materializeDocx()!.originalBuffer!), false);
  draft.loadState(engine.encodeState());
  const saved = await saveYrsDocx(draft, { commentBaseline: { bytes: source.slice(), comments: previous }, commentDeletion: true });
  return saved.bytes;
}

const ROUTES = { worker: workerSave, 'main-thread': mainThreadSave } as const;

for (const [route, save] of Object.entries(ROUTES)) {
  describe(`${route} editor save`, () => {
    test('T1 deleting one comment rewrites only the paragraph that held it', async () => {
      const source = built();
      const saved = await save(source, { remove: 1 });
      expect(differing(source, saved.bytes)).toEqual([0]);
      expect(untouchedFormatting(saved.bytes, [0])).toEqual(untouchedFormatting(source, [0]));
      expect(textAt(saved.bytes, 0)).toBe(textAt(source, 0));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 1)).toEqual([]);
      expect(commentIds(saved.bytes)).toEqual([2, 3, 4, 5]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [2, 3, 4, 5], ids: [2, 3, 4, 5], extensible: [2, 3, 4, 5] });
      expect(marksOf(saved.bytes, [2, 3, 4, 5])).toEqual(marksOf(source, [2, 3, 4, 5]));
      const reread = await reopened(saved.bytes);
      expect(reread.ids).toEqual([2, 3, 4, 5]);
      expect([reread.anchored(2), reread.anchored(5)]).toEqual(['discussed words', 'fifth note']);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment beside a bookmark id the source holds twice rewrites only its paragraph', async () => {
      const source = built(ALL, true, true);
      expect(bookmarkPairs(source)).toHaveLength(2);
      const saved = await save(source, { remove: 1 });
      expect(differing(source, saved.bytes)).toEqual([0]);
      expect(bookmarkPairs(saved.bytes)).toEqual(bookmarkPairs(source));
      expect(marks(saved.bytes, 1)).toEqual([]);
      expect(marksOf(saved.bytes, [2, 3, 4, 5])).toEqual(marksOf(source, [2, 3, 4, 5]));
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('T2 deleting the last comment keeps the comments part as the agents engine does', async () => {
      const source = built([5], false);
      const saved = await save(source, { remove: 5 });
      expect(differing(source, saved.bytes)).toEqual([4]);
      expect(untouchedFormatting(saved.bytes, [4])).toEqual(untouchedFormatting(source, [4]));
      expect(textAt(saved.bytes, 4)).toBe(textAt(source, 4));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 5)).toEqual([]);
      expect(commentsPart(saved.bytes)).toEqual(commentsPart(await agentsDelete(source, 5)));
      expect(commentsPart(saved.bytes)).toEqual({ state: 'emptied', relationship: true, contentType: true });
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('T3 deleting a comment and typing elsewhere rewrites only those two paragraphs', async () => {
      const source = built();
      const saved = await save(source, { remove: 1, type: '0A000002' });
      expect(differing(source, saved.bytes)).toEqual([0, 1]);
      expect(paragraphs(saved.bytes)[1]).toContain(TYPED);
      expect(untouchedFormatting(saved.bytes, [0, 1])).toEqual(untouchedFormatting(source, [0, 1]));
      expect(textAt(saved.bytes, 0)).toBe(textAt(source, 0));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 1)).toEqual([]);
      expect(commentIds(saved.bytes)).toEqual([2, 3, 4, 5]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [2, 3, 4, 5], ids: [2, 3, 4, 5], extensible: [2, 3, 4, 5] });
      expect(marksOf(saved.bytes, [2, 3, 4, 5])).toEqual(marksOf(source, [2, 3, 4, 5]));
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('T5 deleting a comment in a plain paragraph rewrites only that paragraph', async () => {
      const source = built();
      const saved = await save(source, { remove: 5 });
      expect(differing(source, saved.bytes)).toEqual([4]);
      expect(untouchedFormatting(saved.bytes, [4])).toEqual(untouchedFormatting(source, [4]));
      expect(textAt(saved.bytes, 4)).toBe(textAt(source, 4));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 5)).toEqual([]);
      expect(commentIds(saved.bytes)).toEqual([1, 2, 3, 4]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1, 2, 3, 4], ids: [1, 2, 3, 4], extensible: [1, 2, 3, 4] });
      expect(marksOf(saved.bytes, [1, 2, 3, 4])).toEqual(marksOf(source, [1, 2, 3, 4]));
      const reread = await reopened(saved.bytes);
      expect([reread.anchored(1), reread.anchored(2)]).toEqual(['first note', 'discussed words']);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('T6 deleting a comment with replies removes the replies with it', async () => {
      const source = built();
      const saved = await save(source, { remove: 2 });
      expect(differing(source, saved.bytes)).toEqual([3]);
      expect(untouchedFormatting(saved.bytes, [3])).toEqual(untouchedFormatting(source, [3]));
      expect(textAt(saved.bytes, 3)).toBe(textAt(source, 3));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marksOf(saved.bytes, [2, 3, 4])).toEqual({ 2: [], 3: [], 4: [] });
      expect(marksOf(saved.bytes, [1, 5])).toEqual(marksOf(source, [1, 5]));
      expect(commentIds(saved.bytes)).toEqual([1, 5]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1, 5], ids: [1, 5], extensible: [1, 5] });
      expect(orphanReplies(saved.bytes)).toEqual([]);
      const reread = await reopened(saved.bytes);
      expect([reread.ids, reread.orphans]).toEqual([[1, 5], 0]);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('today the paragraph that held a deleted comment is written again and keeps its text', async () => {
      const source = built();
      const saved = await save(source, { remove: 5 });
      const after = paragraphs(saved.bytes)[4] ?? '';
      expect(after === paragraphs(source)[4]).toBe(false);
      expect(textAt(saved.bytes, 4)).toBe('Plain fifth note end');
      expect(after.match(/<w:comment\w*\b[^>]*w:id="5"/g) ?? []).toEqual([]);
    }, TIMEOUT);

    test('T4 a missing comment with changed host metadata still writes every part', async () => {
      const source = built();
      const saved = await save(source, {
        remove: 1,
        host: (host) => {
          const document = host.package.document;
          document.finalSectionProperties = { ...document.finalSectionProperties, marginTop: 2000 };
        },
      });
      expect(saved.full).toBe(true);
      expect(marks(saved.bytes, 1)).toEqual([]);
    }, TIMEOUT);

    test('T4 a save without edits keeps every part', async () => {
      const source = built();
      const saved = await save(source, {});
      expect(saved.full).toBe(false);
      expect(changedParts(source, saved.bytes, new Set())).toEqual([]);
    }, TIMEOUT);

    test('an empty host comment list saves without the source comments', async () => {
      const source = built();
      const saved = await save(source, { comments: () => [] });
      expect(commentIds(saved.bytes)).toEqual([]);
      expect(ALL.flatMap((id) => marks(saved.bytes, id))).toEqual([]);
      expect(paragraphs(saved.bytes).map((xml) => xml.replace(/<[^>]+>/g, ''))).toEqual(
        paragraphs(source).map((xml) => xml.replace(/<[^>]+>/g, ''))
      );
    }, TIMEOUT);
  });
}
