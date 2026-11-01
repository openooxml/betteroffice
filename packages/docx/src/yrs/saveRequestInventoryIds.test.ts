import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import type { Comment } from '../types/document';
import { saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const BODY_IDS = ['10000001', '27BA826F', '6E9E7736', '10000004'];
const ROOT_PARA_ID = '7DDAEF1A';
const REPLY_PARA_ID = '187E0A8B';
const SOURCE_AUTHOR = 'w:author="Original" w:date="2025-01-01T00:00:00Z"';
const decoder = new TextDecoder();
const sessions: YrsSession[] = [];

function fixture(tracked = false): Uint8Array {
  const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const section = (kind: string) => `<w:sectPr><w:type w:val="${kind}"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  const paragraph = (id: string, content: string, kind?: string) =>
    `<w:p w14:paraId="${id}">${kind ? `<w:pPr>${section(kind)}</w:pPr>` : ''}${content}</w:p>`;
  const insertion = (id: number, text: string) => `<w:ins w:id="${id}" ${SOURCE_AUTHOR}>${run(text)}</w:ins>`;
  const deletion = (id: number, text: string) => `<w:del w:id="${id}" ${SOURCE_AUTHOR}><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
  const body = paragraph(BODY_IDS[0], (tracked ? insertion(7, 'Added ') + deletion(19, 'Removed ') : '') + run('First'), 'nextPage') +
    paragraph(BODY_IDS[1], run('Droppable')) +
    paragraph(BODY_IDS[2], (tracked ? insertion(2147483646, 'High ') + deletion(3, 'Gone ') : '') + run('Middle'), 'continuous') +
    paragraph(BODY_IDS[3], run('Tail')) + section('oddPage');
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`)],
  ])));
}

async function open(tracked = false): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 3 });
  sessions.push(session);
  session.openDocx(fixture(tracked), true);
  return session;
}

function expectPackageIds(
  bytes: Uint8Array,
  session: YrsSession,
  revisions: Array<[string, number]>,
  paragraphs = BODY_IDS,
  comments: Array<[number, string]> = []
): void {
  const parts = unzipContainer(bytes);
  const xml = decoder.decode(parts['word/document.xml']);
  expect([...xml.matchAll(/<w:(ins|del)\b[^>]*\bw:id="(\d+)"/g)]
    .map(([, kind, id]) => [kind, Number(id)])).toEqual(revisions);
  expect([...xml.matchAll(/<w:p\b[^>]*\bw14:paraId="([^"]+)"/g)]
    .map(([, id]) => id)).toEqual(paragraphs);
  const commentXml = decoder.decode(parts['word/comments.xml'] ?? new Uint8Array());
  expect([...commentXml.matchAll(/<w:comment\b[^>]*\bw:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/g)]
    .map(([, id, content]) => [Number(id), [...content.matchAll(/<w:p\b[^>]*\bw14:paraId="([^"]+)"/g)].map(([, paraId]) => paraId)]))
    .toEqual(comments.map(([id, paraId]) => [id, [paraId]]));
  expect(xml.match(/<w:sectPr\b/g)).toHaveLength(3);
  expect(session.listSections().map(({ authored }) => [authored.sectionStart, authored.pageWidth, authored.pageHeight]))
    .toEqual([['nextPage', 12240, 15840], ['continuous', 12240, 15840], ['oddPage', 12240, 15840]]);
  if (comments.length === 2) {
    const extended = decoder.decode(parts['word/commentsExtended.xml']);
    expect(extended).toMatch(new RegExp(`<w15:commentEx\\b[^>]*w15:paraId="${REPLY_PARA_ID}"[^>]*w15:paraIdParent="${ROOT_PARA_ID}"`));
  }
}

async function expectSavedIds(
  session: YrsSession,
  revisions: Array<[string, number]>,
  paragraphs = BODY_IDS,
  comments: Array<[number, string]> = []
): Promise<void> {
  expectPackageIds((await saveYrsDocx(session, { updateModifiedDate: false })).bytes, session, revisions, paragraphs, comments);
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))), 30_000);

afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

test('save request inventory retains resolved source revision IDs in unchanged sections', async () => {
  const session = await open(true);
  await expectSavedIds(session, [['ins', 7], ['del', 19], ['ins', 2147483646], ['del', 3]]);
  const resolveRevision = (text: string, kind: 'insertion' | 'deletion', accept: boolean) => {
    const revision = session.listRevisions().find(entry => entry.kind === kind && entry.preview.includes(text));
    if (!revision) throw new Error(`Missing ${kind}: ${text}`);
    const target = { revisionId: revision.revisionId };
    if (accept) session.acceptChange(target);
    else session.rejectChange(target);
  };
  resolveRevision('Added', 'insertion', true);
  await expectSavedIds(session, [['del', 19], ['ins', 2147483646], ['del', 3]]);
  resolveRevision('Removed', 'deletion', false);
  await expectSavedIds(session, [['ins', 2147483646], ['del', 3]]);
  resolveRevision('High', 'insertion', false);
  await expectSavedIds(session, [['del', 3]]);
  resolveRevision('Gone', 'deletion', true);
  await expectSavedIds(session, []);
  session.insertText({ story: 'body', paraId: '10000004', offset: 0 }, 'New ', {
    name: 'Second author', date: '2026-01-02T00:00:00Z',
  });
  await expectSavedIds(session, [['ins', 2147483647]]);
  session.replaceRange({
    story: 'body', start: { paraId: '10000004', offset: 4 }, end: { paraId: '10000004', offset: 8 },
  }, 'End', { name: 'Third author', date: '2026-01-03T00:00:00Z' });
  await expectSavedIds(session, [['ins', 2147483647], ['del', 1], ['ins', 1]]);
});

test('save request inventory preserves paragraph IDs through deletion and re-addition with unchanged sections', async () => {
  const session = await open();
  await expectSavedIds(session, []);
  const remove = (paraId: string) => {
    expect(session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: paraId, lastParaId: paraId }],
    }).ok).toBe(true);
  };
  remove('27BA826F');
  await expectSavedIds(session, [], ['10000001', '6E9E7736', '10000004']);
  const add = (paraId: string, ooxmlParaId: string) => {
    session.applyRawOps('body', [{ op: 'insertEmbed', index: 0, kind: 'pilcrow', payload: { paraId, ooxmlParaId } }]);
    session.insertText({ story: 'body', paraId, offset: 0 }, 'Readded');
  };
  add('readded-first', '20000001');
  await expectSavedIds(session, [], ['20000001', '10000001', '6E9E7736', '10000004']);
  remove('readded-first');
  await expectSavedIds(session, [], ['10000001', '6E9E7736', '10000004']);
  add('readded-again', '20000002');
  await expectSavedIds(session, [], ['20000002', '10000001', '6E9E7736', '10000004']);
});

test('save request inventory fixes comment and reply paragraph IDs through deletion and re-addition', async () => {
  const session = await open();
  await expectSavedIds(session, []);
  const at = { paraId: '10000004', offset: 0 };
  const ranges = [{ story: 'body', start: at, end: { ...at, offset: 4 } }];
  const add = (body: string, parentId?: string) => {
    const { commentId } = session.addComment(ranges, 'Comment author', '2026-01-04T00:00:00Z', body);
    if (parentId) session.updateComment(commentId, { parentId });
    return commentId;
  };
  const root = add('Root');
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID]]);
  const reply = add('Reply', root);
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID], [2, REPLY_PARA_ID]]);
  session.applyRawOps('body', [{ op: 'removeComment', id: reply }]);
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID]]);
  const anotherReply = add('Reply again', root);
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID], [2, REPLY_PARA_ID]]);
  session.applyRawOps('body', [{ op: 'removeComment', id: anotherReply }, { op: 'removeComment', id: root }]);
  await expectSavedIds(session, []);
  const anotherRoot = add('Root again');
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID]]);
  add('Final reply', anotherRoot);
  await expectSavedIds(session, [], BODY_IDS, [[1, ROOT_PARA_ID], [2, REPLY_PARA_ID]]);
});

test('full editor saves reserve comment paragraph IDs from unchanged sections after a source paragraph is deleted', async () => {
  const session = await open();
  const source = session.materializeDocx()!;
  const record: EditorSaveRecord = { full: true };
  expect(session.applyEdits({
    expectVersion: session.version(),
    steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: '27BA826F', lastParaId: '27BA826F' }],
  }).ok).toBe(true);
  let projected = yrsToDocument(session, source);
  const first = await saveEditorDocument(session, projected, [], record);
  const paragraphs = ['10000001', '6E9E7736', '10000004'];
  expectPackageIds(new Uint8Array(first), session, [], paragraphs);
  projected.originalBuffer = first;
  const at = { paraId: '10000004', offset: 0 };
  const { commentId } = session.addComment(
    [{ story: 'body', start: at, end: { ...at, offset: 4 } }],
    'Comment author', '2026-01-04T00:00:00Z', 'Root'
  );
  record.commentIds = new Map([[1, commentId]]);
  const comments: Comment[] = [{
    id: 1, author: 'Comment author', date: '2026-01-04T00:00:00Z',
    content: [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Root' }] }] }],
  }];
  projected = yrsToDocument(session, projected);
  const second = await saveEditorDocument(session, projected, comments, record);
  expectPackageIds(new Uint8Array(second), session, [], paragraphs, [[1, ROOT_PARA_ID]]);
});
