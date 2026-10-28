import { afterEach, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import * as commentOutput from '../src/comment-output';
import { commentFixture } from './fixture';

const opened: DocxAgentDocument[] = [];
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

async function part(bytes: Uint8Array, name: string) {
  return (await JSZip.loadAsync(bytes)).file(name)!.async('string');
}

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the batch to be refused.');
}

async function headerFixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await commentFixture());
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  const relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', relationships.replace('</Relationships>', '<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="header"/>'));
  zip.file('word/header1.xml', '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Review header draft.</w:t></w:r></w:p></w:hdr>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function openWithAcceptedChanges() {
  const doc = await open(await headerFixture());
  const header = doc.grep({ query: 'Review header draft.' }).matches[0];
  const body = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  expect(header.story).not.toBe('body');
  const proposal = doc.propose({ author: 'Editor', edits: [
    { match: header.match, newText: 'Review header approved.' },
    { match: body.match, newText: 'Revenue grew steadily' },
  ] });
  expect(await doc.accept(proposal.id)).toMatchObject({ status: 'accepted', changes: 2 });
  return doc;
}

for (const op of ['add_comment', 'reply_comment', 'resolve_comment'] as const) {
  test(`${op} after an accepted header and body change keeps both changes`, async () => {
    const doc = await openWithAcceptedChanges();
    const before = await doc.export();
    expect(await part(before, 'word/header1.xml')).toContain('Review header approved.');
    const edit: DocxAgentEdit = op === 'add_comment'
      ? { op, match: doc.grep({ query: 'Revenue grew steadily' }).matches[0].match, author: 'Editor', text: 'New note', date: '2026-10-01T00:00:00Z' }
      : op === 'reply_comment'
        ? { op, comment: '0', author: 'Editor', text: 'Follow up', date: '2026-10-01T00:00:00Z' }
        : { op, comment: '0' };
    expect((await doc.edit({ version: doc.list().version, edits: [edit] })).applied).toBe(true);
    const bytes = await doc.export();
    expect(await part(bytes, 'word/header1.xml')).toBe(await part(before, 'word/header1.xml'));
    const reopened = await open(bytes);
    expect(reopened.grep({ query: 'Revenue grew steadily' }).matches).toHaveLength(1);
    const comments = reopened.list({ story: 'comments', limit: 100 }).items;
    if (op === 'add_comment') expect(comments.find(comment => comment.text.startsWith('New note'))).toMatchObject({ parentId: null, author: 'Editor' });
    else if (op === 'reply_comment') expect(comments.find(comment => comment.text.startsWith('Follow up'))).toMatchObject({ parentId: '0', author: 'Editor' });
    else expect(comments.find(comment => comment.id === '0')!.done).toBe(true);
  });
}

test('comment verification compares against the package after each later accepted proposal', async () => {
  const doc = await openWithAcceptedChanges();
  await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'First', date: '2026-10-01T00:00:00Z' }] });
  const header = doc.grep({ query: 'Review header approved.' }).matches[0];
  await doc.accept(doc.propose({ author: 'Editor', edits: [{ match: header.match, newText: 'Review header published.' }] }).id);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'Second', date: '2026-10-02T00:00:00Z' }] });
  const bytes = await doc.export();
  expect(await part(bytes, 'word/header1.xml')).toContain('Review header published.');
  const texts = (await open(bytes)).list({ story: 'comments', limit: 100 }).items.map(comment => comment.text);
  expect(texts.some(text => text.startsWith('First'))).toBe(true);
  expect(texts.some(text => text.startsWith('Second'))).toBe(true);
});

test('comment verification still refuses a header change made by the current batch', async () => {
  const doc = await openWithAcceptedChanges();
  const before = await doc.export();
  const version = doc.list().version;
  const original = commentOutput.verifyCommentOutput;
  const verify = spyOn(commentOutput, 'verifyCommentOutput').mockImplementationOnce(async (source, bytes, added, expected, baseline) => {
    expect(baseline).toBeDefined();
    const zip = await JSZip.loadAsync(bytes);
    zip.file('word/header1.xml', (await zip.file('word/header1.xml')!.async('string')).replace('Review header approved.', 'Corrupted header.'));
    await original(source, await zip.generateAsync({ type: 'uint8array' }), added, expected, baseline);
  });
  try {
    const error = await refusal(doc.edit({ version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'Follow up', date: '2026-10-01T00:00:00Z' }] }));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain('word/header1.xml');
    expect(doc.list().version).toBe(version);
    expect(await part(await doc.export(), 'word/header1.xml')).toBe(await part(before, 'word/header1.xml'));
  } finally { verify.mockRestore(); }
});
