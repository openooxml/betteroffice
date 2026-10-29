import { afterEach, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument } from '../src';
import * as commentOutput from '../src/comment-output';
import { docxEditSchema } from '../src/docx-schema';
import { COMMENT_COMPANIONS, EXISTING_COMMENT, commentFixture } from './fixture';

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

const reference = (id: string) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const markers = (id: string, text: string) => `<w:commentRangeStart w:id="${id}"/>${text}<w:commentRangeEnd w:id="${id}"/>${reference(id)}`;

function withoutComments(xml: string, ids: readonly string[]) {
  for (const id of ids) {
    xml = xml.replace(`<w:commentRangeStart w:id="${id}"/>`, '').replace(`<w:commentRangeEnd w:id="${id}"/>`, '').replace(reference(id), '');
  }
  return xml;
}

async function fixture(story = 'body') {
  const zip = await JSZip.loadAsync(await commentFixture());
  const comments = await zip.file('word/comments.xml')!.async('string');
  zip.file('word/comments.xml', comments.replace('</w:comments>', [
    EXISTING_COMMENT.replace('w:id="0"', 'w:id="1"').replace('1A2B3C4D', '11223344').replace('Keep this note.', 'Reply note.'),
    EXISTING_COMMENT.replace('w:id="0"', 'w:id="2"').replace('1A2B3C4D', '55667788').replace('Keep this note.', 'Other note.'),
    '</w:comments>',
  ].join('')));
  for (const [name, prefix, , root, entry] of COMMENT_COMPANIONS) {
    const xml = await zip.file(`word/${name}.xml`)!.async('string');
    const reply = entry.replace('1A2B3C4D', '11223344').replace('2ECC71AA', 'AABBCCDD')
      .replace('w15:done="0"', 'w15:done="0" w15:paraIdParent="1A2B3C4D"');
    const other = entry.replace('1A2B3C4D', '55667788').replace('2ECC71AA', 'EEFF0011');
    zip.file(`word/${name}.xml`, xml.replace(`</${prefix}:${root}>`, `${reply}${other}</${prefix}:${root}>`));
  }
  let document = await zip.file('word/document.xml')!.async('string');
  document = document.replace('<w:commentRangeStart w:id="0"/>', '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>')
    .replace('<w:commentRangeEnd w:id="0"/>', '<w:commentRangeEnd w:id="1"/>' + reference('1') + '<w:commentRangeEnd w:id="0"/>')
    .replace('<w:r><w:t>Revenue grew strongly in the third quarter.</w:t></w:r>', markers('2', '<w:r><w:t>Revenue grew strongly in the third quarter.</w:t></w:r>'));
  if (story === 'table') {
    const anchor = markers('0', markers('1', '<w:r><w:t>Closing note.</w:t></w:r>'));
    document = withoutComments(document, ['0', '1']).replace('<w:r><w:rPr><w:b/></w:rPr><w:t>Quarter</w:t></w:r>', anchor.replace('Closing note.', 'Quarter'));
  } else if (story !== 'body') {
    const isNote = story === 'footnotes' || story === 'endnotes';
    const local = isNote ? story.slice(0, -1) : story;
    const path = isNote ? `word/${story}.xml` : `word/${story}1.xml`;
    const root = isNote ? story : story === 'header' ? 'hdr' : 'ftr';
    const content = `<w:p>${markers('0', markers('1', '<w:r><w:t>Story note.</w:t></w:r>'))}</w:p>`;
    zip.file(path, `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${isNote ? `<w:${local} w:id="1">${content}</w:${local}>` : content}</w:${root}>`);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}+xml"/></Types>`));
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', `<Relationship Id="anchor" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${story}" Target="${path.slice(5)}"/></Relationships>`));
    document = withoutComments(document, ['0', '1']);
    document = isNote ? document.replace('<w:body>', `<w:body><w:p><w:r><w:${local}Reference w:id="1"/></w:r></w:p>`)
      : document.replace('<w:sectPr>', `<w:sectPr><w:${story}Reference w:type="default" r:id="anchor"/>`);
  }
  zip.file('word/document.xml', document);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

test('delete_comment accepts only a comment ID', () => {
  expect(docxEditSchema.parse({ op: 'delete_comment', comment: '0' })).toEqual({ op: 'delete_comment', comment: '0' });
  expect(docxEditSchema.safeParse({ op: 'delete_comment', comment: '0', done: true }).success).toBe(false);
});

for (const story of ['body', 'table', 'header', 'footer', 'footnotes', 'endnotes']) {
  test(`delete_comment removes a root thread and only its XML in ${story}`, async () => {
    const source = await fixture(story);
    const doc = await open(source);
    const before = await doc.export();
    expect(await part(before, 'word/document.xml')).toBe(await part(source, 'word/document.xml'));
    expect(await part(before, 'word/comments.xml')).toBe(await part(source, 'word/comments.xml'));
    const receipt = await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
    expect(receipt.results[0]).toMatchObject({ comment: '0', deleted: ['0', '1'] });
    expect(doc.list({ story: 'comments' }).items.map(comment => comment.id)).toEqual(['2']);
    const bytes = await doc.export();
    const path = story === 'body' || story === 'table' ? 'word/document.xml'
      : story === 'header' || story === 'footer' ? `word/${story}1.xml` : `word/${story}.xml`;
    expect(await part(bytes, path)).toBe(withoutComments(await part(before, path), ['0', '1']));
    const comments = await part(before, 'word/comments.xml');
    expect(await part(bytes, 'word/comments.xml')).toBe(comments.replace(/<w:comment\b[^>]*w:id="[01]"[^>]*>[\s\S]*?<\/w:comment>/g, ''));
    for (const [name, prefix, , , entry] of COMMENT_COMPANIONS) {
      const original = await part(before, `word/${name}.xml`);
      const rows = original.match(new RegExp(`<${prefix}:[^ />]+\\b[^>]*\\/>`, 'g'))!;
      expect(await part(bytes, `word/${name}.xml`)).toBe(original.replace(rows[0], '').replace(rows[1], ''));
      expect(original).toContain(entry);
    }
    const reopened = await open(bytes);
    expect(reopened.list({ story: 'comments' }).items).toMatchObject([{ id: '2', text: 'Other note.' }]);
  });
}

test('delete_comment removes only a reply and preserves its root and other comments', async () => {
  const doc = await open(await fixture());
  const before = await doc.export();
  expect((await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '1' }] })).results[0])
    .toMatchObject({ comment: '1', deleted: ['1'] });
  const bytes = await doc.export();
  expect(await part(bytes, 'word/document.xml')).toBe(withoutComments(await part(before, 'word/document.xml'), ['1']));
  expect((await open(bytes)).list({ story: 'comments' }).items.map(comment => comment.id)).toEqual(['0', '2']);
});

test('delete_comment removes the last comment and keeps empty registered comment parts', async () => {
  const doc = await open(await commentFixture());
  const before = await doc.export();
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
  const bytes = await doc.export();
  expect(await part(bytes, 'word/document.xml')).toBe(withoutComments(await part(before, 'word/document.xml'), ['0']));
  expect(await part(bytes, 'word/comments.xml')).toBe((await part(before, 'word/comments.xml')).replace(EXISTING_COMMENT, ''));
  for (const [name, , , , entry] of COMMENT_COMPANIONS) {
    expect(await part(bytes, `word/${name}.xml`)).toBe((await part(before, `word/${name}.xml`)).replace(entry, ''));
  }
  expect((await open(bytes)).list({ story: 'comments' }).total).toBe(0);
});

test('delete_comment accepts a legacy comment without paragraph ids', async () => {
  const zip = await JSZip.loadAsync(await commentFixture({ companions: [] }));
  zip.file('word/comments.xml', (await zip.file('word/comments.xml')!.async('string'))
    .replace(' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', '')
    .replace(' w14:paraId="1A2B3C4D" w14:textId="77777777"', ''));
  const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
  expect((await open(await doc.export())).list({ story: 'comments' }).total).toBe(0);
});

test('delete_comment removes replies created earlier in the session', async () => {
  const doc = await open(await commentFixture());
  const reply = await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'Follow up' }] });
  expect((await doc.edit({ version: reply.version, edits: [{ op: 'delete_comment', comment: '0' }] })).results[0].deleted)
    .toEqual(['0', reply.results[0].comment!]);
  expect((await open(await doc.export())).list({ story: 'comments' }).total).toBe(0);
});

test('delete_comment followed by add_comment is atomic and survives reopening', async () => {
  const doc = await open(await fixture());
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  const receipt = await doc.edit({ version: doc.list().version, edits: [
    { op: 'delete_comment', comment: '0' },
    { op: 'add_comment', match: match.match, author: 'Editor', text: 'Correct anchor' },
  ] });
  expect(receipt.results[0]).toMatchObject({ deleted: ['0', '1'] });
  const comments = (await open(await doc.export())).list({ story: 'comments' }).items;
  expect(comments.map(comment => comment.text).sort()).toEqual(['Correct anchor', 'Other note.']);
});

test('delete_comment preserves other edits in the same batch', async () => {
  const doc = await open(await fixture());
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  await doc.edit({ version: doc.list().version, edits: [
    { op: 'format_text', match: match.match, italic: true },
    { op: 'delete_comment', comment: '0' },
  ] });
  const reopened = await open(await doc.export());
  const paragraph = reopened.list().items.find(paragraph => paragraph.text.startsWith('Revenue grew'))!;
  expect(reopened.read(paragraph.ref).runs.some(run => run.formatting.italic)).toBe(true);
  expect(reopened.list({ story: 'comments' }).items.map(comment => comment.id)).toEqual(['2']);
});

test('delete_comment after earlier accepted changes preserves those changes', async () => {
  const doc = await open(await fixture());
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  await doc.accept(doc.propose({ author: 'Editor', edits: [{ match: match.match, newText: 'Revenue grew steadily' }] }).id);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
  expect((await open(await doc.export())).grep({ query: 'Revenue grew steadily' }).matches).toHaveLength(1);
});

test('delete_comment and paragraph insertion share one atomic batch', async () => {
  const doc = await open(await fixture());
  const paragraph = doc.list().items[0];
  await doc.edit({ version: doc.list().version, edits: [
    { op: 'delete_comment', comment: '0' },
    { op: 'insert_paragraphs', after: paragraph.ref, paragraphs: [{ text: 'Added paragraph.' }] },
  ] });
  const reopened = await open(await doc.export());
  expect(reopened.grep({ query: 'Added paragraph.' }).matches).toHaveLength(1);
  expect(reopened.list({ story: 'comments' }).items.map(comment => comment.id)).toEqual(['2']);
});

test('unknown and already deleted comment IDs refuse the entire batch without mutation', async () => {
  const doc = await open(await fixture());
  for (const comment of ['missing', '0', '1']) {
    const version = doc.list().version;
    const state = doc.session.encodeState();
    const before = await doc.export();
    const error = await refusal(doc.edit({ version, edits: [
      { op: 'delete_comment', comment: '0' }, { op: 'delete_comment', comment },
    ] }));
    expect(error.code).toBe('UNKNOWN_REF');
    expect(error.details).toMatchObject({ failedIndex: 1 });
    expect(doc.session.encodeState()).toEqual(state);
    expect(doc.list().version).toBe(version);
    expect(await part(await doc.export(), 'word/comments.xml')).toBe(await part(before, 'word/comments.xml'));
  }
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
  expect((await refusal(doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] }))).code).toBe('UNKNOWN_REF');
});

for (const corruption of ['marker', 'reference', 'companion', 'text', 'comment', 'extra part']) {
  test(`delete_comment verification refuses an unexpected ${corruption}`, async () => {
    const doc = await open(await fixture());
    const version = doc.list().version;
    const state = doc.session.encodeState();
    const original = commentOutput.verifyCommentOutput;
    const verify = spyOn(commentOutput, 'verifyCommentOutput').mockImplementationOnce(async (...args) => {
      const zip = await JSZip.loadAsync(args[1]);
      if (corruption === 'companion') {
        const xml = await zip.file('word/commentsIds.xml')!.async('string');
        zip.file('word/commentsIds.xml', xml.replace('</w16cid:commentsIds>', COMMENT_COMPANIONS[1][4] + '</w16cid:commentsIds>'));
      } else if (corruption === 'comment') {
        const xml = await zip.file('word/comments.xml')!.async('string');
        zip.file('word/comments.xml', xml.replace('Other note.', 'Changed note.'));
      } else if (corruption === 'extra part') zip.file('word/extra.xml', '<extra/>');
      else {
        const xml = await zip.file('word/document.xml')!.async('string');
        zip.file('word/document.xml', corruption === 'text' ? xml.replace('Closing note.', 'Changed text.')
          : xml.replace('</w:p>', `${corruption === 'marker' ? '<w:commentRangeStart w:id="0"/>' : reference('0')}</w:p>`));
      }
      args[1] = await zip.generateAsync({ type: 'uint8array' });
      await original(...args);
    });
    try {
      expect((await refusal(doc.edit({ version, edits: [{ op: 'delete_comment', comment: '0' }] }))).code).toBe('UNSUPPORTED_TARGET');
      expect(doc.session.encodeState()).toEqual(state);
      expect(doc.list().version).toBe(version);
    } finally { verify.mockRestore(); }
  });
}
