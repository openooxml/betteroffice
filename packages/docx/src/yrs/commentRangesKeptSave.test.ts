import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Comment, Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { DirtyProjectionStories, hostSaveMetadata, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

const TIMEOUT = 60_000;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const DATE = '2025-03-04T05:06:07Z';
const OTHER_ID = '0D0000FF';

const RSID = 'w:rsidR="00B10001" w:rsidRPr="00B10001"';
const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:lang w:val="en-GB"/></w:rPr>';
const run = (value: string) => `<w:r ${RSID}>${RPR}<w:t xml:space="preserve">${value}</w:t></w:r>`;
const deletedRun = (value: string) => `<w:r ${RSID}>${RPR}<w:delText xml:space="preserve">${value}</w:delText></w:r>`;
const revision = (tag: 'ins' | 'del', id: number, inner: string) =>
  `<w:${tag} w:id="${id}" w:author="G" w:date="${DATE}">${inner}</w:${tag}>`;
const paragraph = (id: string, inner: string) =>
  `<w:p w14:paraId="${id}" w14:textId="${id}" w:rsidR="00B10001" w:rsidRDefault="00B10001" w:rsidP="00B10001">${inner}</w:p>`;
const shaped = (construct: string) => paragraph('0D000001', `${run('Head ')}${construct}${run(' tail')}`);
const OTHER = paragraph(OTHER_ID, run('Other paragraph'));
const rangeStart = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
const rangeEnd = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
const reference = (id: number) =>
  `<w:r ${RSID}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

const IN_INSERTION = revision('ins', 41, `${rangeStart(1)}${run('inserted')}${rangeEnd(1)}`);
const RANGE_IN_INSERTION = shaped(`${IN_INSERTION}${reference(1)}`);
const RANGE_IN_LINK = paragraph('0D000001',
  `${run('Head ')}<w:hyperlink r:id="rIdLink" w:history="1">${rangeStart(1)}${run('linked')}${rangeEnd(1)}</w:hyperlink>` +
  `${reference(1)}${run(' and ')}${rangeStart(2)}${run('noted')}${rangeEnd(2)}${reference(2)}${run(' tail')}`
);
const RANGE_IN_DELETION = shaped(
  `${revision('del', 42, `${rangeStart(1)}${deletedRun('removed')}${rangeEnd(1)}`)}${reference(1)}`
);
const RANGE_IN_CONTROL = shaped(
  `<w:sdt><w:sdtPr><w:id w:val="1234"/></w:sdtPr><w:sdtContent>${rangeStart(1)}${run('controlled')}${rangeEnd(1)}</w:sdtContent></w:sdt>` +
  reference(1)
);
const SECOND_COMMENT = `${rangeStart(2)}${run('noted')}${rangeEnd(2)}${reference(2)}`;
const BESIDE_SECOND = shaped(`${IN_INSERTION}${reference(1)}${run(' and ')}${SECOND_COMMENT}`);

function commentsXml(ids: number[]): string {
  return `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
    ids.map((id) =>
      `<w:comment w:id="${id}" w:author="Grace" w:date="${DATE}" w:initials="G">` +
      `<w:p w14:paraId="0E00000${id}" w14:textId="77777777"><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`
    ).join('') +
    '</w:comments>';
}

function fixture(first: string, comments: number[]): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdLink" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
      `<Relationship Id="rIdComments" Type="${REL}/comments" Target="comments.xml"/>` +
      '</Relationships>'
    )],
    ['word/styles.xml', toBytes(
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>'
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${REL}" xmlns:w14="${W14}"><w:body>${first}${OTHER}<w:sectPr/></w:body></w:document>`
    )],
    ['word/comments.xml', toBytes(commentsXml(comments))],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const decoder = new TextDecoder();
const part = (bytes: Uint8Array, name: string) => {
  const data = unzipContainer(bytes)[name];
  return data ? decoder.decode(data) : '';
};
const paragraphs = (bytes: Uint8Array) =>
  part(bytes, 'word/document.xml').match(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g) ?? [];
const textAt = (bytes: Uint8Array, at: number) => (paragraphs(bytes)[at] ?? '').replace(/<[^>]+>/g, '');
const marks = (bytes: Uint8Array, id: number) =>
  [...part(bytes, 'word/document.xml').matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id=["'](\d+)["']\s*\/>/g)]
    .filter((match) => Number(match[2]) === id)
    .map((match) => match[1]!);

const owned: Array<{ destroy(): void }> = [];
let startWorker: (clientId?: number) => InProcessResidentWorker;
let nextClient = 73_100;

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

interface Edit {
  /** The comment the editor deletes: the session removes it and the host's list leaves it out. */
  remove?: number;
  /** The comment only the host's list leaves out. */
  drop?: number;
  /** The Word paragraph id of the paragraph a letter X is typed at the end of. */
  type?: string;
}

function apply(engine: YrsSession, edit: Edit): boolean {
  if (edit.remove !== undefined) engine.applyRawOps('body', [{ op: 'removeComment', id: String(edit.remove) }]);
  if (edit.type !== undefined) {
    const result = engine.applyEdits({
      expectVersion: engine.version(),
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: edit.type }, at: 'end', text: 'X' }],
    });
    if (!result.ok) throw new Error(`insertText: ${JSON.stringify(result).slice(0, 200)}`);
  }
  return edit.remove !== undefined || edit.type !== undefined;
}

function hostList(host: Document, edit: Edit): Comment[] {
  const list = host.package.document.comments ?? [];
  const gone = [edit.remove, edit.drop].filter((id): id is number => id !== undefined);
  return list.filter((comment) => !gone.includes(comment.id) && (comment.parentId == null || !gone.includes(comment.parentId)));
}

async function mainThreadSave(source: Uint8Array, edit: Edit): Promise<Uint8Array> {
  const editor = await session();
  const host = editor.openDocx(source.slice(), true).document;
  const dirty = new DirtyProjectionStories();
  if (apply(editor, edit)) dirty.add('body');
  const projected = yrsToDocument(editor, mergeDocxHostMetadata(editor.materializeDocx() as Document, host), dirty.projectionOptions());
  return new Uint8Array(await saveEditorDocument(editor, projected, hostList(host, edit)));
}

async function workerSave(source: Uint8Array, edit: Edit): Promise<Uint8Array> {
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
  const saved = await client.save({
    comments: hostList(host, edit),
    host: hostSaveMetadata(host),
    ...(edited ? { stories: ['body'] } : {}),
  });
  return new Uint8Array(saved.bytes);
}

const ROUTES = { worker: workerSave, 'main-thread': mainThreadSave } as const;

const typedNext = (bytes: Uint8Array) => textAt(bytes, 1);

for (const [route, save] of Object.entries(ROUTES)) {
  describe(`${route} editor save`, () => {
    test('a letter typed in another paragraph leaves a comment range inside a tracked insertion as it was', async () => {
      const saved = await save(fixture(RANGE_IN_INSERTION, [1]), { type: OTHER_ID });
      expect(paragraphs(saved)[0]).toBe(RANGE_IN_INSERTION);
      expect(typedNext(saved)).toBe('Other paragraphX');
    }, TIMEOUT);

    test('a letter typed in another paragraph leaves a comment range inside a hyperlink as it was', async () => {
      const saved = await save(fixture(RANGE_IN_LINK, [1, 2]), { type: OTHER_ID });
      expect(paragraphs(saved)[0]).toBe(RANGE_IN_LINK);
      expect(typedNext(saved)).toBe('Other paragraphX');
    }, TIMEOUT);

    test('a letter typed in another paragraph leaves a comment range inside a tracked deletion as it was', async () => {
      const saved = await save(fixture(RANGE_IN_DELETION, [1]), { type: OTHER_ID });
      expect(paragraphs(saved)[0]).toBe(RANGE_IN_DELETION);
      expect(typedNext(saved)).toBe('Other paragraphX');
    }, TIMEOUT);

    test('a letter typed in another paragraph leaves a comment range inside a content control as it was', async () => {
      const saved = await save(fixture(RANGE_IN_CONTROL, [1]), { type: OTHER_ID });
      expect(paragraphs(saved)[0]).toBe(RANGE_IN_CONTROL);
      expect(typedNext(saved)).toBe('Other paragraphX');
    }, TIMEOUT);

    test('a comment the host list leaves out, saved beside a letter typed in another paragraph, takes only its markers out of a paragraph that also holds a comment range inside a tracked insertion', async () => {
      const saved = await save(fixture(BESIDE_SECOND, [1, 2]), { drop: 2, type: OTHER_ID });
      const expected = BESIDE_SECOND.replace(SECOND_COMMENT, run('noted'));
      expect(paragraphs(saved)[0]).toBe(expected);
      expect(marks(saved, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      expect(typedNext(saved)).toBe('Other paragraphX');
    }, TIMEOUT);

    test('deleting the comment whose range is inside a tracked insertion takes its markers out', async () => {
      const saved = await save(fixture(RANGE_IN_INSERTION, [1]), { remove: 1 });
      const first = paragraphs(saved)[0] ?? '';
      expect(marks(saved, 1)).toEqual([]);
      expect(first).toContain(revision('ins', 41, run('inserted')));
      expect(textAt(saved, 0)).toBe('Head inserted tail');
    }, TIMEOUT);

    for (const [name, first, comments] of [
      ['a comment range inside a tracked insertion', RANGE_IN_INSERTION, [1]],
      ['a comment range inside a hyperlink', RANGE_IN_LINK, [1, 2]],
    ] as const) {
      test(`a save with no edit leaves both paragraphs as they were, beside ${name}`, async () => {
        const saved = await save(fixture(first, [...comments]), {});
        expect(paragraphs(saved)).toEqual([first, OTHER]);
      }, TIMEOUT);
    }
  });
}
