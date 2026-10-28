import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { readCommentPart, verifyCommentOutput } from '../src/comment-output';
import { commentFixture, editFixture, EXISTING_COMMENT } from './fixture';

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

const COMMENT_ZERO = '<w:comment w:id="0" w:author="First Reviewer" w:date="2026-01-05T09:00:00Z" w:initials="FR"><w:p w14:paraId="1A2B3C01" w14:textId="77777771" w:rsidR="00A1" w:rsidRDefault="00A2"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">First identity proofed here. </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Second identity </w:t></w:r><w:r><w:rPr><w:color w:val="0066AA"/></w:rPr><w:t>proofed there.</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p><w:p w14:paraId="1A2B3C02" w14:textId="77777772" w:rsidR="00A3"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:t>Closing line stays.</w:t></w:r></w:p></w:comment>';
const COMMENT_ONE = '<w:comment w:id="1" w:author="Second Reviewer" w:date="2026-01-06T09:00:00Z" w:initials="SR"><w:p w14:paraId="1A2B3C11" w14:textId="77777773" w:rsidR="00B1"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">Other page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink w:anchor="target"><w:r><w:t>identity proofed link</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> done.</w:t></w:r></w:p></w:comment>';
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
    .replace('<w:t xml:space="preserve">First identity proofed here. </w:t>', '<w:t xml:space="preserve">First identity-proofed here. </w:t>')
    .replace('<w:t xml:space="preserve">Second identity </w:t>', '<w:t xml:space="preserve">Second identity-proofed</w:t>')
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
  expect(saved).toContain('<w:t xml:space="preserve">Second identity-proofed</w:t>');
  const reopened = await open(bytes);
  const edited = reopened.list({ story: 'comments', limit: 100 }).items.find(item => item.id === '0')!;
  expect(edited.text).toBe('First identity-proofed here. Second identity-proofed there.\nClosing line stays.');
});

test('replace_comment_text replaces an existing default xml:space value when adding edge whitespace', async () => {
  const zip = await JSZip.loadAsync(await commentTextFixture());
  const xml = (await zip.file('word/comments.xml')!.async('string'))
    .replace('<w:t>Closing line stays.</w:t>', '<w:t xml:space="default">alpha</w:t>');
  zip.file('word/comments.xml', xml);
  const source = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const doc = await open(source);
  expect((await doc.edit({ version: doc.list().version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'alpha', replace: ' alpha ' }] })).applied).toBe(true);
  const bytes = await doc.export();
  expect(await part(bytes, 'word/comments.xml')).toBe(xml.replace('<w:t xml:space="default">alpha</w:t>', '<w:t xml:space="preserve"> alpha </w:t>'));
  for (const name of ['word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/document.xml']) {
    expect(await part(bytes, name)).toBe(await part(source, name));
  }
  const reopened = await open(bytes);
  expect(reopened.list({ story: 'comments', limit: 100 }).items.find(comment => comment.id === '0')!.text).toEndWith('\n alpha ');
});

for (const space of ['default', undefined]) {
  test(`replace_comment_text refuses to preserve untouched edge whitespace with ${space ?? 'missing'} xml:space atomically`, async () => {
    const zip = await JSZip.loadAsync(await commentFixture({ companions: [] }));
    const comment = `<w:comment w:id="0" w:author="Reviewer" w:initials="R"><w:p w14:paraId="1A2B3C4D"><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>alpha</w:t></w:r><w:r><w:t${space ? ` xml:space="${space}"` : ''}> beta </w:t></w:r></w:p></w:comment>`;
    const xml = (await zip.file('word/comments.xml')!.async('string')).replace(EXISTING_COMMENT, comment);
    zip.file('word/comments.xml', xml);
    const source = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    const doc = await open(source);
    const before = doc.list({ story: 'comments', limit: 100 });
    const state = doc.session.encodeState();
    const baseline = await doc.export();
    const error = await refusal(doc.edit({ version: before.version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'alpha', replace: 'omega' }] }));
    expect(error.code).toBe('UNSUPPORTED_TARGET');
    expect(error.message).toContain('Saving comments would rewrite an existing comment.');
    expect(error.message).toContain('No batch edits were applied.');
    expect(doc.list({ story: 'comments', limit: 100 })).toEqual(before);
    expect(doc.session.encodeState()).toEqual(state);
    expect(await doc.export()).toEqual(baseline);
  });
}

for (const op of ['reply_comment', 'resolve_comment'] as const) {
  for (const separate of [false, true]) {
    test(`replace_comment_text and ${op} add only the missing paragraph id in ${separate ? 'separate batches' : 'one batch'}`, async () => {
      const zip = await JSZip.loadAsync(await commentFixture({ companions: [] }));
      const original = COMMENT_ZERO.replace(/ w14:paraId="[^"]*"/g, '');
      const xml = (await zip.file('word/comments.xml')!.async('string')).replace(EXISTING_COMMENT, `${original}${COMMENT_ONE}`);
      zip.file('word/comments.xml', xml);
      const source = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
      const doc = await open(source);
      const edits: DocxAgentEdit[] = [
        { op: 'replace_comment_text', comment: '0', find: 'Closing line stays.', replace: 'Closing line updated.' },
        ...(op === 'reply_comment' ? [{ op, comment: '0', author: 'Editor', text: 'Follow up', date: '2026-10-01T00:00:00Z' } satisfies DocxAgentEdit] : []),
        { op: 'resolve_comment', comment: '0' },
      ];
      let paragraphId: string | undefined;
      let bytes = source;
      for (const batch of separate ? edits.map(edit => [edit]) : [edits]) {
        expect((await doc.edit({ version: doc.list().version, edits: batch })).applied).toBe(true);
        bytes = await doc.export();
        const saved = (await part(bytes, 'word/comments.xml'))!;
        const comments = saved.match(/<w:comment\b[\s\S]*?<\/w:comment>/g)!;
        const root = comments[0];
        const ids = [...root.matchAll(/ w14:paraId="([^"]*)"/g)];
        if (batch.every(edit => edit.op === 'replace_comment_text')) {
          expect(ids).toHaveLength(0);
        } else {
          expect(ids).toHaveLength(1);
          expect(root).toContain(`<w:p w14:textId="77777772" w:rsidR="00A3" w14:paraId="${ids[0][1]}">`);
          if (paragraphId) expect(ids[0][1]).toBe(paragraphId);
          paragraphId = ids[0][1];
        }
        expect(root.replace(/ w14:paraId="[^"]*"/g, '')).toBe(original.replace('Closing line stays.', 'Closing line updated.'));
        expect(comments[1]).toBe(COMMENT_ONE);
        const withoutReply = comments[2] ? saved.replace(comments[2], '') : saved;
        expect(withoutReply.replace(root, root.replace(/ w14:paraId="[^"]*"/g, ''))).toBe(xml.replace('Closing line stays.', 'Closing line updated.'));
      }
      expect(paragraphId).toBeDefined();
      const reopened = await open(bytes);
      const comments = reopened.list({ story: 'comments', limit: 100 }).items;
      expect(comments.find(comment => comment.id === '0')).toMatchObject({ done: true, text: 'First identity proofed here. Second identity proofed there.\nClosing line updated.' });
      expect(comments).toHaveLength(op === 'reply_comment' ? 3 : 2);
      if (op === 'reply_comment') expect(comments.find(comment => comment.text === 'Follow up')).toMatchObject({ parentId: '0', author: 'Editor' });
      const input = await readCommentPart(await zip.generateAsync({ type: 'arraybuffer' }));
      const expected = comments.map(comment => ({ id: comment.id, parentId: comment.parentId, ...(comment.id === '0' ? { done: true } : {}) }));
      const output = await JSZip.loadAsync(bytes);
      const saved = (await part(bytes, 'word/comments.xml'))!;
      for (const changed of [
        saved.replace('w:author="First Reviewer"', 'w:author="Changed Reviewer"'),
        saved.replace('w:rsidR="00A1"', 'w:rsidR="00FF"'),
        saved.replace('<w:b/>', '<w:i/>'),
        saved.replace('<w:p w14:textId="77777771"', '<w:p w14:paraId="12345678" w14:textId="77777771"'),
      ]) {
        output.file('word/comments.xml', changed);
        const error = await refusal(verifyCommentOutput(input, await output.generateAsync({ type: 'uint8array' }), op === 'reply_comment' ? 1 : 0, expected, new Set(['0'])));
        expect(error.code).toBe('UNSUPPORTED_TARGET');
        expect(error.message).toContain('rewrite an existing comment');
      }
    });
  }
}

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

for (const op of ['replace_comment_text', 'reply_comment', 'resolve_comment'] as const) {
  test(`a prior comment text replacement survives a later ${op} batch`, async () => {
    const doc = await open();
    const first = await doc.edit({ version: doc.list().version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'identity proofed', replace: 'identity-proofed' }] });
    const edit: DocxAgentEdit = op === 'replace_comment_text'
      ? { op, comment: '2', find: 'identity proofed', replace: 'identity-proofed' }
      : op === 'reply_comment'
        ? { op, comment: '0', author: 'Editor', text: 'Follow up', date: '2026-10-01T00:00:00Z' }
        : { op, comment: '0' };
    const second = await doc.edit({ version: first.version, edits: [edit] });
    expect(second.applied).toBe(true);
    const reopened = await open(await doc.export());
    const comments = reopened.list({ story: 'comments', limit: 100 }).items;
    expect(comments.find(comment => comment.id === '0')!.text).toBe('First identity-proofed here. Second identity-proofed there.\nClosing line stays.');
    if (op === 'replace_comment_text') {
      expect(comments.find(comment => comment.id === '2')).toMatchObject({ parentId: '0', text: 'Reply identity-proofed too.' });
    } else if (op === 'reply_comment') {
      expect(comments.find(comment => comment.text === 'Follow up')).toMatchObject({ parentId: '0', author: 'Editor' });
    } else {
      expect(comments.find(comment => comment.id === '0')!.done).toBe(true);
    }
  });
}

test('replace_comment_text refuses source comments with unpreserved edge whitespace', async () => {
  const zip = await JSZip.loadAsync(await commentTextFixture());
  const xml = await zip.file('word/comments.xml')!.async('string');
  zip.file('word/comments.xml', xml.replace('<w:t xml:space="preserve">Other page </w:t>', '<w:t>Other page </w:t>'));
  const doc = await open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const before = doc.list({ story: 'comments', limit: 100 });
  const error = await refusal(doc.edit({ version: before.version, edits: [{ op: 'replace_comment_text', comment: '0', find: 'identity proofed', replace: 'identity-proofed' }] }));
  expect(error.code).toBe('UNSUPPORTED_TARGET');
  expect(error.message).toContain('Saving comments would rewrite an existing comment.');
  expect(doc.list({ story: 'comments', limit: 100 }).items).toEqual(before.items);
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
