import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { editFixture } from './fixture';

const opened: DocxAgentDocument[] = [];
async function open(bytes?: Uint8Array) {
  const doc = await openDocx(bytes ?? await commentTextFixture());
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const part = async (bytes: Uint8Array, name: string) => (await JSZip.loadAsync(bytes)).file(name)?.async('string');

async function refusal(promise: Promise<unknown>): Promise<DocumentToolError> {
  try { await promise; } catch (error) { return error as DocumentToolError; }
  throw new Error('Expected the batch to be refused.');
}

const COMMENT_ZERO = '<w:comment w:id="0" w:author="First Reviewer" w:date="2026-01-05T09:00:00Z" w:initials="FR"><w:p w14:paraId="1A2B3C01" w14:textId="77777771" w:rsidR="00A1" w:rsidRDefault="00A2"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>First identity proofed here. </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>Second identity </w:t></w:r><w:r><w:rPr><w:color w:val="0066AA"/></w:rPr><w:t>proofed there.</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p><w:p w14:paraId="1A2B3C02" w14:textId="77777772" w:rsidR="00A3"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:t>Closing line stays.</w:t></w:r></w:p></w:comment>';
const COMMENT_ONE = '<w:comment w:id="1" w:author="Second Reviewer" w:date="2026-01-06T09:00:00Z" w:initials="SR"><w:p w14:paraId="1A2B3C11" w14:textId="77777773" w:rsidR="00B1"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Other page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple><w:r><w:t> and </w:t></w:r><w:hyperlink w:anchor="target"><w:r><w:t>identity proofed link</w:t></w:r></w:hyperlink><w:r><w:t> done.</w:t></w:r></w:p></w:comment>';
const COMMENT_TWO = '<w:comment w:id="2" w:author="First Reviewer" w:date="2026-01-07T09:00:00Z" w:initials="FR"><w:p w14:paraId="1A2B3C21" w14:textId="77777774" w:rsidR="00C1"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Reply identity proofed too.</w:t></w:r></w:p></w:comment>';

async function commentTextFixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await editFixture());
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  let types = await zip.file('[Content_Types].xml')!.async('string');
  let relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  types = types.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="application/vnd.ms-word.commentsIds+xml"/><Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.ms-word.commentsExtensible+xml"/></Types>');
  relationships = relationships.replace('</Relationships>', `<Relationship Id="comments" Type="${r}/comments" Target="comments.xml"/><Relationship Id="commentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="commentsIds" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/><Relationship Id="commentsExtensible" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/></Relationships>`);
  zip.file('[Content_Types].xml', types);
  zip.file('word/_rels/document.xml.rels', relationships);
  const anchor = (id: string, text: string) => `<w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r></w:p>`;
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml
    .replace('<w:p><w:r><w:t>Obsolete remark.</w:t></w:r></w:p>', anchor('0', 'Obsolete remark.'))
    .replace('<w:p><w:r><w:t>Closing note.</w:t></w:r></w:p>', anchor('1', 'Closing note.')));
  zip.file('word/comments.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${COMMENT_ZERO}${COMMENT_ONE}${COMMENT_TWO}</w:comments>`);
  zip.file('word/commentsExtended.xml', '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="1A2B3C02" w15:done="0"/><w15:commentEx w15:paraId="1A2B3C11" w15:done="1"/><w15:commentEx w15:paraId="1A2B3C21" w15:done="0" w15:paraIdParent="1A2B3C02"/></w15:commentsEx>');
  zip.file('word/commentsIds.xml', '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="1A2B3C02" w16cid:durableId="AA000001"/><w16cid:commentId w16cid:paraId="1A2B3C11" w16cid:durableId="AA000002"/><w16cid:commentId w16cid:paraId="1A2B3C21" w16cid:durableId="AA000003"/></w16cid:commentsIds>');
  zip.file('word/commentsExtensible.xml', '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="AA000001" w16cex:dateUtc="2026-01-05T09:00:00Z"/><w16cex:commentExtensible w16cex:durableId="AA000002" w16cex:dateUtc="2026-01-06T09:00:00Z"/><w16cex:commentExtensible w16cex:durableId="AA000003" w16cex:dateUtc="2026-01-07T09:00:00Z"/></w16cex:commentsExtensible>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

test('replace_comment_text rewrites every occurrence in one comment and saves only its w:t text', async () => {
  const source = await commentTextFixture();
  const doc = await open(source);
  const outline = doc.list({ story: 'comments', limit: 100 });
  expect(outline.items.map(item => item.id)).toEqual(['0', '1', '2']);
  const receipt = await doc.edit({ version: outline.version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'identity proofed', replace: 'identity-proofed' }] });
  expect(receipt.results).toEqual([{ index: 0, op: 'replace_comment_text', comment: '0', replacements: 2 }]);
  const edited = doc.list({ story: 'comments', limit: 100 }).items.find(item => item.id === '0')!;
  expect(edited.text).toBe('First identity-proofed here. Second identity-proofed there.\nClosing line stays.');
  expect(edited.author).toBe('First Reviewer');
  const bytes = await doc.export();
  const saved = (await part(bytes, 'word/comments.xml'))!;
  const expected = (await part(source, 'word/comments.xml'))!
    .replace('<w:t>First identity proofed here. </w:t>', '<w:t xml:space="preserve">First identity-proofed here. </w:t>')
    .replace('<w:t>Second identity </w:t>', '<w:t>Second identity-proofed</w:t>')
    .replace('<w:t>proofed there.</w:t>', '<w:t xml:space="preserve"> there.</w:t>');
  expect(saved).toBe(expected);
  expect(saved).toContain(COMMENT_ONE);
  expect(saved).toContain(COMMENT_TWO);
  for (const name of ['word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/document.xml']) {
    expect(await part(bytes, name)).toBe(await part(source, name));
  }
  const reopened = await open(bytes);
  const comments = reopened.list({ story: 'comments', limit: 100 }).items;
  expect(comments.find(item => item.id === '0')!.text).toContain('identity-proofed');
  expect(comments.find(item => item.id === '0')!.text).not.toContain('identity proofed');
  expect(comments.find(item => item.id === '2')!.parentId).toBe('0');
  const resaved = (await part(await reopened.export(), 'word/comments.xml'))!;
  expect(resaved).toContain('First identity-proofed here. ');
  expect(resaved).toContain(COMMENT_ONE);
});

test('replace_comment_text preserves edge whitespace of patched runs on save and reopen', async () => {
  const doc = await open();
  await doc.edit({ version: doc.list().version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'identity proofed', replace: 'identity-proofed' }] });
  const bytes = await doc.export();
  const saved = (await part(bytes, 'word/comments.xml'))!;
  expect(saved).toContain('<w:t xml:space="preserve">First identity-proofed here. </w:t>');
  expect(saved).toContain('<w:t xml:space="preserve"> there.</w:t>');
  expect(saved).toContain('<w:t>Second identity-proofed</w:t>');
  const reopened = await open(bytes);
  const edited = reopened.list({ story: 'comments', limit: 100 }).items.find(item => item.id === '0')!;
  expect(edited.text).toBe('First identity-proofed here. Second identity-proofed there.\nClosing line stays.');
});

test('replace_comment_text edits a reply without changing its thread or done state', async () => {
  const source = await commentTextFixture();
  const doc = await open(source);
  const receipt = await doc.edit({ version: doc.list().version, edits: [{ op: 'replace_comment_text', comment: '2', find: 'identity proofed', replace: 'identity-proofed' }] });
  expect(receipt.results).toEqual([{ index: 0, op: 'replace_comment_text', comment: '2', replacements: 1 }]);
  const bytes = await doc.export();
  const saved = (await part(bytes, 'word/comments.xml'))!;
  expect(saved).toContain(COMMENT_ZERO);
  expect(saved).toContain(COMMENT_ONE);
  expect(saved).toContain('Reply identity-proofed too.');
  expect(await part(bytes, 'word/commentsExtended.xml')).toBe(await part(source, 'word/commentsExtended.xml'));
  const reopened = await open(bytes);
  const comments = reopened.list({ story: 'comments', limit: 100 }).items;
  expect(comments.find(item => item.id === '2')).toMatchObject({ parentId: '0', done: false, text: 'Reply identity-proofed too.' });
  expect(comments.find(item => item.id === '0')!.done).toBe(false);
  expect(comments.find(item => item.id === '1')!.done).toBe(true);
});

test('replace_comment_text refuses missing text, unknown comments, line breaks and guarded content', async () => {
  const doc = await open();
  const edit = (edits: DocxAgentEdit[]) => doc.edit({ version: doc.list().version, edits });
  const missing = await refusal(edit([{ op: 'replace_comment_text', comment: '0', find: 'missing phrase', replace: 'x' }]));
  expect(missing.code).toBe('NO_MATCH');
  expect(missing.message).toContain('First identity proofed here.');
  const unknown = await refusal(edit([{ op: 'replace_comment_text', comment: '99', find: 'identity proofed', replace: 'x' }]));
  expect(unknown.code).toBe('UNKNOWN_REF');
  for (const edits of [
    [{ op: 'replace_comment_text', comment: '0', find: 'identity\nproofed', replace: 'x' }],
    [{ op: 'replace_comment_text', comment: '0', find: 'identity\tproofed', replace: 'x' }],
    [{ op: 'replace_comment_text', comment: '0', find: 'identity proofed', replace: 'identity\tproofed' }],
  ] as DocxAgentEdit[][]) {
    expect((await refusal(edit(edits))).code).toBe('INVALID_ARGUMENT');
  }
  const field = await refusal(edit([{ op: 'replace_comment_text', comment: '1', find: 'page 7', replace: 'page 8' }]));
  expect(field.code).toBe('UNSUPPORTED_TARGET');
  const link = await refusal(edit([{ op: 'replace_comment_text', comment: '1', find: 'identity proofed', replace: 'identity-proofed' }]));
  expect(link.code).toBe('UNSUPPORTED_TARGET');
  const comments = doc.list({ story: 'comments', limit: 100 }).items;
  expect(comments.find(item => item.id === '0')!.text).toContain('identity proofed');
  expect(comments.find(item => item.id === '1')!.text).toContain('identity proofed link');
});
