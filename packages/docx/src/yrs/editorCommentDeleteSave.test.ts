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

/** A comment's range around `text`, then its reference run. */
const commented = (id: number, text: string) =>
  `<w:commentRangeStart w:id="${id}"/>${run(text)}<w:commentRangeEnd w:id="${id}"/>${reference(id)}`;

/** A table of one cell holding `paragraph`. */
const table = (paragraph: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl>`;

function body(ids: number[], twice = false): string {
  const has = (id: number) => ids.includes(id);
  const ranged = (id: number, inner: string) =>
    has(id) ? `<w:commentRangeStart w:id="${id}"/>${inner}<w:commentRangeEnd w:id="${id}"/>${reference(id)}` : inner;
  const thread = [2, 3, 4].filter(has);
  const field = `<w:fldSimple w:instr=" PAGE ">${run('1')}</w:fldSimple>`;
  const cell = table(para(3, bookmark(twice, 'Cell', run('Cell paragraph'))));
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

function built(ids: number[] = ALL, companions = true, twice = false, content = body(ids, twice)): Uint8Array {
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
      `<w:document xmlns:w="${W}" xmlns:r="${REL}" xmlns:w14="${W14}"><w:body>${content}<w:sectPr/></w:body></w:document>`
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

/** A part a fixture adds beside the document, related from it. */
interface Added {
  name: string;
  kind: string;
  type: string;
  xml: string;
}

/** `bytes` with each part in `added` related from the document as `rIdA<index>`, its section properties `section` when given. */
function withParts(bytes: Uint8Array, added: Added[], section?: string): Uint8Array {
  const parts: PartsMap = new Map(Object.entries(unzipContainer(bytes)));
  const change = (name: string, edit: (xml: string) => string) => parts.set(name, toBytes(edit(part(bytes, name))));
  const overrides = added.map(({ name, kind }) => `<Override PartName="/word/${name}.xml" ContentType="${OFFICE}.${kind}+xml"/>`);
  const relationships = added.map(({ name, type }, at) => `<Relationship Id="rIdA${at}" Type="${type}" Target="${name}.xml"/>`);
  change('[Content_Types].xml', (xml) => xml.replace('</Types>', `${overrides.join('')}</Types>`));
  change('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>', `${relationships.join('')}</Relationships>`));
  if (section) change('word/document.xml', (xml) => xml.replace('<w:sectPr/>', section));
  for (const { name, xml } of added) parts.set(`word/${name}.xml`, toBytes(xml));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** `bytes` with part `name` changed by `edit`. */
function withPart(bytes: Uint8Array, name: string, edit: (xml: string) => string): Uint8Array {
  const parts: PartsMap = new Map(Object.entries(unzipContainer(bytes)));
  parts.set(name, toBytes(edit(part(bytes, name))));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const header = (inner: string): Added => ({
  name: 'header1',
  kind: 'header',
  type: `${REL}/header`,
  xml: `<w:hdr xmlns:w="${W}" xmlns:w14="${W14}"><w:p w14:paraId="0C000001" w14:textId="0C000001">${inner}</w:p></w:hdr>`,
});
const HEADED = '<w:sectPr><w:headerReference w:type="default" r:id="rIdA0"/></w:sectPr>';

const footnotes = (inner: string): Added => ({
  name: 'footnotes',
  kind: 'footnotes',
  type: `${REL}/footnotes`,
  xml:
    `<w:footnotes xmlns:w="${W}" xmlns:w14="${W14}">` +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    `<w:footnote w:id="1"><w:p w14:paraId="0C000002" w14:textId="0C000002"><w:r><w:footnoteRef/></w:r>${inner}</w:p></w:footnote>` +
    '</w:footnotes>',
});

const PEOPLE: Added = {
  name: 'people',
  kind: 'people',
  type: 'http://schemas.microsoft.com/office/2011/relationships/people',
  xml:
    `<w15:people xmlns:w15="${W15}"><w15:person w15:author="Grace">` +
    '<w15:presenceInfo w15:providerId="None" w15:userId="Grace"/></w15:person></w15:people>',
};

/** The text of part `name`. */
const textOf = (bytes: Uint8Array, name: string) => part(bytes, name).replace(/<[^>]+>/g, '');
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

/** The comment's markers in part `name`, the body unless given, in document order. */
const marks = (bytes: Uint8Array, id: number, name = 'word/document.xml') =>
  [...part(bytes, name).matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id=["'](\d+)["']\s*\/>/g)]
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

/** Each companion part with whether it is present, related from the document and typed. */
const companionParts = (bytes: Uint8Array) =>
  COMMENT_PARTS.slice(1).map((name) => [
    name,
    name in unzipContainer(bytes),
    part(bytes, 'word/_rels/document.xml.rels').includes(`Target="${name.slice('word/'.length)}"`),
    part(bytes, '[Content_Types].xml').includes(`PartName="/${name}"`),
  ]);

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
  /** The text `type` types, `TYPED` unless given. */
  text?: string;
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
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: edit.type }, at: 'end', text: edit.text ?? TYPED }],
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
      expect(commentIds(saved.bytes)).toEqual([2, 3, 4, 5]);
      expect(marksOf(saved.bytes, [2, 3, 4, 5])).toEqual(marksOf(source, [2, 3, 4, 5]));
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment between two paragraphs that hold one bookmark id rewrites only its paragraph', async () => {
      const marked = (n: number, name: string) => para(n, bookmark(true, name, run(`${name} paragraph`)));
      const noted = `<w:commentRangeStart w:id="1"/>${run('noted')}<w:commentRangeEnd w:id="1"/>${reference(1)}`;
      const content = marked(1, 'First') + para(2, `${run('Lead ')}${noted}${run(' end')}`) + marked(3, 'Third');
      const source = built([1], false, false, content);
      expect(bookmarkPairs(source)).toHaveLength(2);
      const saved = await save(source, { remove: 1 });
      expect(differing(source, saved.bytes)).toEqual([1]);
      expect(bookmarkPairs(saved.bytes)).toEqual(bookmarkPairs(source));
      expect(marks(saved.bytes, 1)).toEqual([]);
      expect(commentIds(saved.bytes)).toEqual([]);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('T2 deleting the last comment keeps the comments part as the agents engine does', async () => {
      const source = withParts(built([5]), [PEOPLE]);
      const saved = await save(source, { remove: 5 });
      expect(differing(source, saved.bytes)).toEqual([4]);
      expect(untouchedFormatting(saved.bytes, [4])).toEqual(untouchedFormatting(source, [4]));
      expect(textAt(saved.bytes, 4)).toBe(textAt(source, 4));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(part(saved.bytes, 'word/people.xml')).toBe(PEOPLE.xml);
      expect(marks(saved.bytes, 5)).toEqual([]);
      expect(commentsPart(saved.bytes)).toEqual(commentsPart(await agentsDelete(source, 5)));
      expect(commentsPart(saved.bytes)).toEqual({ state: 'emptied', relationship: true, contentType: true });
      expect(companionIds(saved.bytes)).toEqual({ extended: [], ids: [], extensible: [] });
      expect(companionParts(saved.bytes)).toEqual(COMMENT_PARTS.slice(1).map((name) => [name, true, true, true]));
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

    test('a host-side deletion removes the markers of a comment whose ids are padded', async () => {
      const padded =
        `<w:commentRangeStart w:id="05"/>${run('fifth note')}<w:commentRangeEnd w:id="05"/>` +
        '<w:r><w:commentReference w:id="05"/></w:r>';
      const content = para(1, run('Lead paragraph')) + para(2, `${run('Plain ')}${padded}${run(' end')}`);
      const source = built([5], false, false, content);
      expect(marks(source, 5)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      const saved = await save(source, { comments: () => [] });
      expect(marks(saved.bytes, 5)).toEqual([]);
      expect(part(saved.bytes, 'word/document.xml')).not.toMatch(/<w:comment(RangeStart|RangeEnd|Reference)\b/);
      expect(commentIds(saved.bytes)).toEqual([]);
      expect(paragraphs(saved.bytes).map((xml) => xml.replace(/<[^>]+>/g, ''))).toEqual(
        paragraphs(source).map((xml) => xml.replace(/<[^>]+>/g, ''))
      );
    }, TIMEOUT);

    test('deleting a comment anchored in a header removes it from the header alone', async () => {
      const content = para(1, `${run('Body ')}${commented(1, 'body note')}`);
      const source = withParts(built([1, 6], true, false, content), [header(`${run('Header ')}${commented(6, 'header words')}`)], HEADED);
      const name = 'word/header1.xml';
      expect(marks(source, 6, name)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      const saved = await save(source, { remove: 6 });
      expect(marks(saved.bytes, 6, name)).toEqual([]);
      expect(textOf(saved.bytes, name)).toBe(textOf(source, name));
      expect(commentIds(saved.bytes)).toEqual([1]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1], ids: [1], extensible: [1] });
      expect(marks(saved.bytes, 1)).toEqual(marks(source, 1));
      expect(differing(source, saved.bytes)).toEqual([]);
      expect(changedParts(source, saved.bytes, new Set([...CHANGEABLE, name]))).toEqual([]);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment anchored in a footnote removes it from the footnotes alone', async () => {
      const content = para(1, `${run('Body ')}<w:r><w:footnoteReference w:id="1"/></w:r>${commented(1, 'body note')}`);
      const source = withParts(built([1, 6], true, false, content), [footnotes(`${run(' Note ')}${commented(6, 'footnote words')}`)]);
      const name = 'word/footnotes.xml';
      expect(marks(source, 6, name)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      const saved = await save(source, { remove: 6 });
      expect(marks(saved.bytes, 6, name)).toEqual([]);
      expect(textOf(saved.bytes, name)).toBe(textOf(source, name));
      expect(commentIds(saved.bytes)).toEqual([1]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1], ids: [1], extensible: [1] });
      expect(marks(saved.bytes, 1)).toEqual(marks(source, 1));
      expect(differing(source, saved.bytes)).toEqual([]);
      expect(changedParts(source, saved.bytes, new Set([...CHANGEABLE, name]))).toEqual([]);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment anchored in a table cell rewrites only the cell paragraph', async () => {
      const content =
        para(1, `${run('Body ')}${commented(1, 'body note')}`) +
        table(para(2, `${run('Cell ')}${commented(6, 'cell words')}`)) +
        para(3, run('After the table'));
      const source = built([1, 6], true, false, content);
      const saved = await save(source, { remove: 6 });
      expect(differing(source, saved.bytes)).toEqual([1]);
      expect(textAt(saved.bytes, 1)).toBe(textAt(source, 1));
      expect(untouchedFormatting(saved.bytes, [1])).toEqual(untouchedFormatting(source, [1]));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 6)).toEqual([]);
      expect(marks(saved.bytes, 1)).toEqual(marks(source, 1));
      expect(commentIds(saved.bytes)).toEqual([1]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1], ids: [1], extensible: [1] });
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment anchored in a text box removes it from the box', async () => {
      const box = (inner: string) =>
        '<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="Box" style="width:216pt;height:108pt">' +
        `<v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
      const content =
        para(1, `${run('Body ')}${commented(1, 'body note')}`) +
        para(2, `${run('Box ')}${box(para(3, `${run('In box ')}${commented(6, 'box words')}`))}`) +
        para(4, run('After the box'));
      const source = built([1, 6], true, false, content);
      expect(marks(source, 6)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      const saved = await save(source, { comments: (list) => list.filter((comment) => comment.id !== 6) });
      expect(marks(saved.bytes, 6)).toEqual([]);
      expect(textAt(saved.bytes, 1)).toBe(textAt(source, 1));
      expect(differing(source, saved.bytes)).toEqual([1]);
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 1)).toEqual(marks(source, 1));
      expect(commentIds(saved.bytes)).toEqual([1]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1], ids: [1], extensible: [1] });
    }, TIMEOUT);

    test('deleting a single reply keeps its root and the other reply', async () => {
      const source = built();
      const saved = await save(source, { remove: 3 });
      expect(differing(source, saved.bytes)).toEqual([3]);
      expect(untouchedFormatting(saved.bytes, [3])).toEqual(untouchedFormatting(source, [3]));
      expect(textAt(saved.bytes, 3)).toBe(textAt(source, 3));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 3)).toEqual([]);
      expect(marksOf(saved.bytes, [1, 2, 4, 5])).toEqual(marksOf(source, [1, 2, 4, 5]));
      expect(commentIds(saved.bytes)).toEqual([1, 2, 4, 5]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1, 2, 4, 5], ids: [1, 2, 4, 5], extensible: [1, 2, 4, 5] });
      expect(orphanReplies(saved.bytes)).toEqual([]);
      const reread = await reopened(saved.bytes);
      expect([reread.ids, reread.orphans]).toEqual([[1, 2, 4, 5], 0]);
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a resolved comment removes its done flag with it', async () => {
      const open = `w15:paraId="${paraId(5)}" w15:done="0"`;
      const source = withPart(built(),'word/commentsExtended.xml', (xml) => xml.replace(open, `w15:paraId="${paraId(5)}" w15:done="1"`));
      expect(part(source, 'word/commentsExtended.xml')).toContain('w15:done="1"');
      const saved = await save(source, { remove: 5 });
      expect(differing(source, saved.bytes)).toEqual([4]);
      expect(textAt(saved.bytes, 4)).toBe(textAt(source, 4));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 5)).toEqual([]);
      expect(marksOf(saved.bytes, [1, 2, 3, 4])).toEqual(marksOf(source, [1, 2, 3, 4]));
      expect(commentIds(saved.bytes)).toEqual([1, 2, 3, 4]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [1, 2, 3, 4], ids: [1, 2, 3, 4], extensible: [1, 2, 3, 4] });
      expect(part(saved.bytes, 'word/commentsExtended.xml')).not.toContain('w15:done="1"');
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('deleting a comment whose range spans two paragraphs writes every paragraph again, as before', async () => {
      const content =
        para(1, `${run('Lead ')}<w:commentRangeStart w:id="1"/>${run('first half')}`) +
        para(2, `${run('second half')}<w:commentRangeEnd w:id="1"/>${reference(1)}${run(' after')}`) +
        para(3, `${run('Plain ')}${commented(5, 'fifth note')}${run(' end')}`);
      const source = built([1, 5], true, false, content);
      const saved = await save(source, { remove: 1 });
      expect(differing(source, saved.bytes)).toEqual([0, 1, 2]);
      expect([0, 1, 2].map((at) => textAt(saved.bytes, at))).toEqual([0, 1, 2].map((at) => textAt(source, at)));
      expect(changedParts(source, saved.bytes)).toEqual([]);
      expect(marks(saved.bytes, 1)).toEqual([]);
      expect(marks(saved.bytes, 5)).toEqual(marks(source, 5));
      expect(commentIds(saved.bytes)).toEqual([5]);
      expect(companionIds(saved.bytes)).toEqual({ extended: [5], ids: [5], extensible: [5] });
      expect(saved.full).toBe(false);
    }, TIMEOUT);

    test('typing and deleting a comment save an untouched header alike', async () => {
      const content = para(1, `${run('Lead ')}${commented(1, 'noted')}${run(' end')}`) + para(2, run('Other paragraph'));
      const source = withParts(built([1], false, false, content), [header('<w:r><w:t>A & B</w:t></w:r>')], HEADED);
      const name = 'word/header1.xml';
      expect(part(source, name)).toContain('<w:t>A & B</w:t>');
      const typed = await save(source, { type: '0A000001', text: 'x' });
      const deleted = await save(source, { remove: 1 });
      expect(part(deleted.bytes, name)).toBe(part(typed.bytes, name));
    }, TIMEOUT);

    test('typing and deleting a comment treat a bookmark nested in an insertion alike', async () => {
      const inserted =
        '<w:ins w:id="20" w:author="Grace" w:date="2025-03-04T05:06:07Z"><w:bookmarkStart w:id="7" w:name="Inserted"/>' +
        `${run('inserted')}<w:bookmarkEnd w:id="7"/></w:ins>`;
      const content =
        para(1, bookmark(true, 'First', run('First paragraph'))) +
        para(2, `${run('Lead ')}${inserted}${commented(1, 'noted')}${run(' end')}`) +
        para(3, bookmark(true, 'Third', run('Third paragraph')));
      const source = built([1], false, false, content);
      const nested = (bytes: Uint8Array) => (paragraphs(bytes)[1] ?? '').match(/<w:bookmark(?:Start|End) w:id="7"/g)?.length ?? 0;
      expect(nested(source)).toBe(2);
      const typed = await save(source, { type: '0A000002', text: 'x' });
      const deleted = await save(source, { remove: 1 });
      expect(nested(deleted.bytes)).toBe(nested(typed.bytes));
    }, TIMEOUT);

    test('typing and deleting a comment write the paragraphs after the edit alike', async () => {
      const content =
        para(1, `${run('Lead ')}${commented(1, 'x note')}${run(' a')}`) +
        para(2, `${run('Y starts ')}<w:commentRangeStart w:id="2"/>${run('y words')}<w:commentRangeEnd w:id="2"/>${run(' more text')}`) +
        para(3, `${reference(2)}${run('C text')}`);
      const source = built([1, 2], false, false, content);
      const typed = await save(source, { type: '0A000001', text: 'x' });
      const deleted = await save(source, { remove: 1 });
      expect(paragraphs(deleted.bytes).slice(1)).toEqual(paragraphs(typed.bytes).slice(1));
    }, TIMEOUT);
  });
}
