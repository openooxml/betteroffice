import { afterEach, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocumentToolError, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import * as commentOutput from '../src/comment-output';
import { docxEditSchema } from '../src/docx-schema';
import { COMMENT_COMPANIONS, EXISTING_COMMENT, commentFixture, editFixture } from './fixture';

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

async function fixture(story = 'body', companions?: readonly string[]) {
  const zip = await JSZip.loadAsync(await commentFixture({ companions }));
  const comments = await zip.file('word/comments.xml')!.async('string');
  zip.file('word/comments.xml', comments.replace('</w:comments>', [
    EXISTING_COMMENT.replace('w:id="0"', 'w:id="1"').replace('1A2B3C4D', '11223344').replace('Keep this note.', 'Reply note.'),
    EXISTING_COMMENT.replace('w:id="0"', 'w:id="2"').replace('1A2B3C4D', '55667788').replace('Keep this note.', 'Other note.'),
    '</w:comments>',
  ].join('')));
  for (const [name, prefix, , root, entry] of COMMENT_COMPANIONS) {
    if (!zip.file(`word/${name}.xml`)) continue;
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

for (const [label, companions] of [
  ['no companion parts', []],
  ['only commentsExtended', ['commentsExtended']],
  ['all companion parts', COMMENT_COMPANIONS.map(([name]) => name)],
] as const) {
  for (const survivorRows of companions.length ? [false, true] : [false]) {
    test(`delete_comment preserves ${label} ${survivorRows ? 'with' : 'without'} survivor rows`, async () => {
      const zip = await JSZip.loadAsync(await commentFixture({ companions }));
      const comments = await zip.file('word/comments.xml')!.async('string');
      zip.file('word/comments.xml', comments.replace('</w:comments>',
        EXISTING_COMMENT.replace('w:id="0"', 'w:id="1"').replace('1A2B3C4D', '11223344').replace('Keep this note.', 'Surviving note.') + '</w:comments>'));
      const document = await zip.file('word/document.xml')!.async('string');
      const anchor = '<w:r><w:t>Revenue grew strongly in the third quarter.</w:t></w:r>';
      zip.file('word/document.xml', document.replace(anchor, markers('1', anchor)));
      for (const [name, prefix, , root, entry] of COMMENT_COMPANIONS) {
        if (!zip.file(`word/${name}.xml`) || !survivorRows) continue;
        const xml = await zip.file(`word/${name}.xml`)!.async('string');
        const survivor = entry.replace('1A2B3C4D', '11223344').replace('2ECC71AA', 'AABBCCDD');
        zip.file(`word/${name}.xml`, xml.replace(`</${prefix}:${root}>`, `${survivor}</${prefix}:${root}>`));
      }
      const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
      await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
      const bytes = await doc.export();
      const saved = await JSZip.loadAsync(bytes);
      zip.file('word/comments.xml', (await zip.file('word/comments.xml')!.async('string')).replace(EXISTING_COMMENT, ''));
      zip.file('word/document.xml', withoutComments(await zip.file('word/document.xml')!.async('string'), ['0']));
      for (const [name, , , , entry] of COMMENT_COMPANIONS) {
        const file = zip.file(`word/${name}.xml`);
        if (file) zip.file(file.name, (await file.async('string')).replace(entry, ''));
      }
      const paths = Object.keys(zip.files).filter(path => !zip.files[path].dir).sort();
      expect(Object.keys(saved.files).filter(path => !saved.files[path].dir).sort()).toEqual(paths);
      for (const path of paths) {
        expect(await saved.file(path)!.async('uint8array')).toEqual(await zip.file(path)!.async('uint8array'));
      }
      expect((await open(bytes)).list({ story: 'comments' }).items).toMatchObject([{ id: '1', text: 'Surviving note.' }]);
    });
  }
}

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

test('delete_comment verifies core properties without comparing different save timestamps', async () => {
  const zip = await JSZip.loadAsync(await commentFixture());
  const core = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Kept title</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>';
  zip.file('docProps/core.xml', core);
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
  zip.file('_rels/.rels', (await zip.file('_rels/.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="core" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
  const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
  const original = Date.prototype.toISOString;
  let now = Date.parse('2026-10-01T00:00:00Z');
  const clock = spyOn(Date.prototype, 'toISOString').mockImplementation(() => original.call(new Date(now += 1000)));
  try {
    expect((await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] })).applied).toBe(true);
  } finally { clock.mockRestore(); }
  expect(await part(await doc.export(), 'docProps/core.xml')).toContain('<dc:title>Kept title</dc:title>');
});

async function twoAddedComments() {
  const doc = await open(await editFixture());
  const first = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  const second = doc.grep({ query: 'Closing note.' }).matches[0];
  const result = await doc.edit({ version: doc.list().version, edits: [
    { op: 'add_comment', match: first.match, author: 'Editor', text: 'First note.' },
    { op: 'add_comment', match: second.match, author: 'Editor', text: 'Second note.' },
  ] });
  return { doc, first: result.results[0].comment!, second: result.results[1].comment! };
}

test('delete_comment preserves a session-created survivor and every companion identity', async () => {
  const { doc, first } = await twoAddedComments();
  const before = await doc.export();
  const comments = await part(before, 'word/comments.xml');
  const survivor = comments.match(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g)!.find(xml => xml.includes('Second note.'))!;
  const survivorId = survivor.match(/w:id="([^"]+)"/)![1];
  const anchors = (xml: string) => [...xml.matchAll(/<w:(commentRangeStart|commentRangeEnd|commentReference)\b[^>]*w:id="([^"]+)"[^>]*\/>/g)]
    .map(([xml, kind, id]) => ({ xml, kind, id }));
  const survivorAnchors = anchors(await part(before, 'word/document.xml')).filter(anchor => anchor.id === survivorId);
  expect(survivorAnchors.map(anchor => anchor.kind).sort()).toEqual(['commentRangeEnd', 'commentRangeStart', 'commentReference']);
  const deleted = comments.match(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g)!.find(xml => xml.includes('First note.'))!;
  const paraId = [...deleted.matchAll(/w14:paraId="([^"]+)"/g)].at(-1)![1];
  const ids = await part(before, 'word/commentsIds.xml');
  const deletedId = ids.match(/<w16cid:commentId\b[^>]*\/>/g)!.find(xml => xml.includes(paraId))!;
  const durableId = deletedId.match(/w16cid:durableId="([^"]+)"/)![1];
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: first }] });
  const bytes = await doc.export();
  expect(anchors(await part(bytes, 'word/document.xml'))).toEqual(survivorAnchors);
  expect(await part(bytes, 'word/comments.xml')).toBe(comments.replace(deleted, ''));
  for (const [name, prefix] of COMMENT_COMPANIONS) {
    const original = await part(before, `word/${name}.xml`);
    const removed = original.match(new RegExp(`<${prefix}:[^ />]+\\b[^>]*\\/>`, 'g'))!.find(row => row.includes(name === 'commentsExtensible' ? durableId : paraId))!;
    expect(await part(bytes, `word/${name}.xml`)).toBe(original.replace(removed, ''));
  }
  expect((await open(bytes)).list({ story: 'comments' }).items).toMatchObject([{ text: 'Second note.' }]);
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  await doc.edit({ version: doc.list().version, edits: [{ op: 'add_comment', match: match.match, author: 'Editor', text: 'Third note.' }] });
  const afterAdd = await part(await doc.export(), 'word/comments.xml');
  expect(afterAdd).toContain(survivor);
  expect(afterAdd).toContain('Third note.');
});

for (const [label, kinds, nativeId] of [
  ['start range id', ['commentRangeStart'], true],
  ['end range id', ['commentRangeEnd'], true],
  ['both range ids', ['commentRangeStart', 'commentRangeEnd'], true],
  ['unknown range ids', ['commentRangeStart', 'commentRangeEnd'], false],
] as const) {
  test(`delete_comment refuses a survivor's ${label} changed away from its saved id`, async () => {
    const { doc, first, second } = await twoAddedComments();
    const before = await doc.export();
    const comments = await part(before, 'word/comments.xml');
    const survivor = comments.match(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g)!.find(xml => xml.includes('Second note.'))!;
    const id = survivor.match(/w:id="([^"]+)"/)![1];
    const replacement = nativeId ? second : 'missing';
    expect(replacement).not.toBe(id);
    const state = doc.session.encodeState();
    const version = doc.list().version;
    const original = commentOutput.verifyCommentOutput;
    const verify = spyOn(commentOutput, 'verifyCommentOutput').mockImplementationOnce(async (...args) => {
      const zip = await JSZip.loadAsync(args[1]);
      const xml = await zip.file('word/document.xml')!.async('string');
      let tampered = xml;
      for (const kind of kinds) {
        tampered = tampered.replace(new RegExp(`(<w:${kind}\\b[^>]*w:id=")${id}(")`, 'g'),
          (_match, start, end) => `${start}${replacement}${end}`);
      }
      expect(tampered).not.toBe(xml);
      expect(tampered).toContain(`<w:commentReference w:id="${id}"`);
      expect(await zip.file('word/comments.xml')!.async('string')).toContain(survivor);
      zip.file('word/document.xml', tampered);
      args[1] = await zip.generateAsync({ type: 'uint8array' });
      await original(...args);
    });
    try {
      expect((await refusal(doc.edit({ version, edits: [{ op: 'delete_comment', comment: first }] }))).code).toBe('UNSUPPORTED_TARGET');
      expect(doc.session.encodeState()).toEqual(state);
      expect(doc.list().version).toBe(version);
      expect(await part(await doc.export(), 'word/document.xml')).toBe(await part(before, 'word/document.xml'));
    } finally { verify.mockRestore(); }
  });
}

test('delete_comment keeps the registered empty parts of a session-created last comment', async () => {
  const doc = await open(await editFixture());
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  const added = await doc.edit({ version: doc.list().version, edits: [{ op: 'add_comment', match: match.match, author: 'Editor', text: 'Temporary note.' }] });
  const before = await JSZip.loadAsync(await doc.export());
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: added.results[0].comment! }] });
  const bytes = await doc.export();
  expect(await part(bytes, 'word/comments.xml')).toBe((await before.file('word/comments.xml')!.async('string')).replace(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g, ''));
  for (const [name, prefix] of COMMENT_COMPANIONS) {
    expect(await part(bytes, `word/${name}.xml`)).toBe((await before.file(`word/${name}.xml`)!.async('string')).replace(new RegExp(`<${prefix}:[^ />]+\\b[^>]*\\/>`, 'g'), ''));
  }
  for (const path of ['word/_rels/document.xml.rels', '[Content_Types].xml']) {
    expect(await part(bytes, path)).toBe(await before.file(path)!.async('string'));
  }
  expect(await part(bytes, 'word/document.xml')).not.toContain('w:comment');
  expect((await open(bytes)).list({ story: 'comments' }).total).toBe(0);
});

for (const companions of [[], ['commentsExtended'], COMMENT_COMPANIONS.map(([name]) => name)]) {
  for (const target of ['added reply', 'added root', 'source reply']) {
    test(`delete_comment preserves the pre-batch package after deleting ${target} with ${companions.length} source companions`, async () => {
      const doc = await open(await fixture('body', companions));
      let comment = '1';
      if (target !== 'source reply') {
        const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
        const edit: DocxAgentEdit = target === 'added reply'
          ? { op: 'reply_comment', comment: '0', author: 'Editor', text: 'Temporary note.' }
          : { op: 'add_comment', match: match.match, author: 'Editor', text: 'Temporary note.' };
        comment = (await doc.edit({ version: doc.list().version, edits: [edit] })).results[0].comment!;
      }
      const before = await JSZip.loadAsync(await doc.export());
      const comments = await before.file('word/comments.xml')!.async('string');
      const removed = comments.match(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g)!
        .find(xml => target === 'source reply' ? xml.includes('w:id="1"') : xml.includes('Temporary note.'))!;
      const id = removed.match(/w:id="([^"]+)"/)![1];
      const paraId = [...removed.matchAll(/w14:paraId="([^"]+)"/g)].at(-1)![1];
      const ids = await before.file('word/commentsIds.xml')?.async('string');
      const durableId = ids?.match(/<w16cid:commentId\b[^>]*\/>/g)?.find(row => row.includes(paraId))?.match(/w16cid:durableId="([^"]+)"/)?.[1];
      const receipt = await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment }] });
      expect(receipt.results[0]).toMatchObject({ comment, deleted: [comment] });
      const bytes = await doc.export();
      const saved = await JSZip.loadAsync(bytes);
      before.file('word/comments.xml', comments.replace(removed, ''));
      const document = await before.file('word/document.xml')!.async('string');
      const removedReference = document.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)!
        .find(run => run.includes(`<w:commentReference w:id="${id}"`))!;
      expect(removedReference).not.toContain('<w:t');
      before.file('word/document.xml', document
        .replace(new RegExp(`<w:commentRange(?:Start|End)\\b[^>]*w:id="${id}"[^>]*/>`, 'g'), '')
        .replace(removedReference, ''));
      for (const [name, prefix] of COMMENT_COMPANIONS) {
        const file = before.file(`word/${name}.xml`);
        if (!file) continue;
        const xml = await file.async('string');
        const key = name === 'commentsExtensible' ? durableId : paraId;
        const row = key && xml.match(new RegExp(`<${prefix}:[^ />]+\\b[^>]*\\/>`, 'g'))?.find(row => row.includes(key));
        before.file(file.name, row ? xml.replace(row, '') : xml);
      }
      const paths = Object.keys(before.files).filter(path => !before.files[path].dir).sort();
      expect(Object.keys(saved.files).filter(path => !saved.files[path].dir).sort()).toEqual(paths);
      for (const path of paths) {
        expect(await saved.file(path)!.async('uint8array'), path).toEqual(await before.file(path)!.async('uint8array'));
      }
      expect((await open(bytes)).list({ story: 'comments' }).items.map(comment => comment.text).sort())
        .toEqual(doc.list({ story: 'comments' }).items.map(comment => comment.text).sort());
    });
  }
}

for (const corruption of ['text', 'numeric id', 'paragraph id', 'durable id', 'companion row', 'missing companion part']) {
  test(`delete_comment rejects a changed session-created survivor's ${corruption}`, async () => {
    const { doc, first } = await twoAddedComments();
    await doc.export();
    const state = doc.session.encodeState();
    const version = doc.list().version;
    const original = commentOutput.verifyCommentOutput;
    const verify = spyOn(commentOutput, 'verifyCommentOutput').mockImplementationOnce(async (...args) => {
      const zip = await JSZip.loadAsync(args[1]);
      const comments = await zip.file('word/comments.xml')!.async('string');
      const survivor = comments.match(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g)!.find(xml => xml.includes('Second note.'))!;
      if (corruption === 'text') zip.file('word/comments.xml', comments.replace('Second note.', 'Changed note.'));
      else if (corruption === 'numeric id') {
        const id = survivor.match(/w:id="([^"]+)"/)![1];
        zip.file('word/comments.xml', comments.replace(`w:id="${id}"`, 'w:id="99"'));
        zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string')).replaceAll(`w:id="${id}"`, 'w:id="99"'));
      } else if (corruption === 'paragraph id') {
        const id = [...survivor.matchAll(/w14:paraId="([^"]+)"/g)].at(-1)![1];
        for (const name of ['comments', 'commentsExtended', 'commentsIds']) zip.file(`word/${name}.xml`, (await zip.file(`word/${name}.xml`)!.async('string')).replaceAll(id, '01020304'));
      } else if (corruption === 'durable id') {
        const ids = await zip.file('word/commentsIds.xml')!.async('string');
        const id = ids.match(/w16cid:durableId="([^"]+)"/)![1];
        for (const name of ['commentsIds', 'commentsExtensible']) zip.file(`word/${name}.xml`, (await zip.file(`word/${name}.xml`)!.async('string')).replaceAll(id, '10203040'));
      } else if (corruption === 'missing companion part') zip.remove('word/commentsIds.xml');
      else zip.file('word/commentsExtended.xml', (await zip.file('word/commentsExtended.xml')!.async('string')).replace('<w15:commentEx ', '<w15:commentEx extra="changed" '));
      args[1] = await zip.generateAsync({ type: 'uint8array' });
      await original(...args);
    });
    try {
      expect((await refusal(doc.edit({ version, edits: [{ op: 'delete_comment', comment: first }] }))).code).toBe('UNSUPPORTED_TARGET');
      expect(doc.session.encodeState()).toEqual(state);
      expect(doc.list().version).toBe(version);
    } finally { verify.mockRestore(); }
  });
}

test('delete_comment accepts a legacy comment without paragraph ids', async () => {
  const zip = await JSZip.loadAsync(await commentFixture({ companions: [] }));
  zip.file('word/comments.xml', (await zip.file('word/comments.xml')!.async('string'))
    .replace(' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', '')
    .replace(' w14:paraId="1A2B3C4D" w14:textId="77777777"', ''));
  const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
  await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
  expect((await open(await doc.export())).list({ story: 'comments' }).total).toBe(0);
});

for (const companions of [[], COMMENT_COMPANIONS.map(([name]) => name)]) {
  test(`delete_comment removes replies created in an earlier batch with ${companions.length} source companions`, async () => {
    const doc = await open(await commentFixture({ companions }));
    const reply = await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'Follow up' }] });
    expect((await doc.edit({ version: reply.version, edits: [{ op: 'delete_comment', comment: '0' }] })).results[0].deleted)
      .toEqual(['0', reply.results[0].comment!]);
    expect((await open(await doc.export())).list({ story: 'comments' }).total).toBe(0);
  });

  for (const op of ['add_comment', 'reply_comment', 'resolve_comment'] as const) {
    for (const deleteFirst of [false, true]) {
      test(`delete_comment refuses ${op} ${deleteFirst ? 'after' : 'before'} deletion with ${companions.length} source companions`, async () => {
        const doc = await open(await fixture('body', companions));
        const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
        const other: DocxAgentEdit = op === 'add_comment'
          ? { op, match: match.match, author: 'Editor', text: 'Correct anchor' }
          : op === 'reply_comment' ? { op, comment: '0', author: 'Editor', text: 'Follow up' }
            : { op, comment: '0', done: true };
        const deletion: DocxAgentEdit = { op: 'delete_comment', comment: '0' };
        const version = doc.list().version;
        const state = doc.session.encodeState();
        const comments = doc.list({ story: 'comments' });
        const before = await JSZip.loadAsync(await doc.export());
        const error = await refusal(doc.edit({ version, edits: [
          { op: 'format_text', match: match.match, italic: true },
          ...(deleteFirst ? [deletion, other] : [other, deletion]),
        ] }));
        expect(error.code).toBe('INVALID_ARGUMENT');
        expect(error.message).toContain('delete_comment cannot be combined with other comment operations');
        expect(error.message).toContain('run it in its own docx_edit call');
        expect(doc.session.encodeState()).toEqual(state);
        expect(doc.list().version).toBe(version);
        expect(doc.list({ story: 'comments' })).toEqual(comments);
        const saved = await JSZip.loadAsync(await doc.export());
        const paths = Object.keys(before.files).filter(path => !before.files[path].dir).sort();
        expect(Object.keys(saved.files).filter(path => !saved.files[path].dir).sort()).toEqual(paths);
        for (const path of paths) {
          expect(await saved.file(path)!.async('uint8array')).toEqual(await before.file(path)!.async('uint8array'));
        }
      });
    }
  }

  test(`delete_comment then add_comment in consecutive batches survives reopening with ${companions.length} source companions`, async () => {
    const doc = await open(await fixture('body', companions));
    const deleted = await doc.edit({ version: doc.list().version, edits: [{ op: 'delete_comment', comment: '0' }] });
    expect(deleted.results[0]).toMatchObject({ deleted: companions.length ? ['0', '1'] : ['0'] });
    const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
    await doc.edit({ version: deleted.version, edits: [
      { op: 'add_comment', match: match.match, author: 'Editor', text: 'Correct anchor' },
    ] });
    const comments = (await open(await doc.export())).list({ story: 'comments' }).items;
    expect(comments.map(comment => comment.text).sort()).toEqual(companions.length
      ? ['Correct anchor', 'Other note.'] : ['Correct anchor', 'Other note.', 'Reply note.']);
  });
}

test('delete_comment refuses mixed comment batches before planning any edit', async () => {
  const doc = await open(await fixture());
  const version = doc.list().version;
  const state = doc.session.encodeState();
  const error = await refusal(doc.edit({ version, edits: [
    { op: 'format_text', match: 'missing', italic: true },
    { op: 'resolve_comment', comment: 'missing' },
    { op: 'delete_comment', comment: 'missing' },
  ] }));
  expect(error.code).toBe('INVALID_ARGUMENT');
  expect(error.message).toContain('run it in its own docx_edit call');
  expect(doc.session.encodeState()).toEqual(state);
  expect(doc.list().version).toBe(version);
});

for (const deleteFirst of [false, true]) {
  test(`delete_comment refuses raw replace_comment_text input ${deleteFirst ? 'after' : 'before'} deletion`, async () => {
    const doc = await open(await fixture());
    const version = doc.list().version;
    const state = doc.session.encodeState();
    const edits: DocxAgentEdit[] = [
      { op: 'delete_comment', comment: '0' },
      { op: 'replace_comment_text', comment: '0', text: 'Changed note.' } as unknown as DocxAgentEdit,
    ];
    const error = await refusal(doc.edit({ version, edits: deleteFirst ? edits : edits.reverse() }));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain('run it in its own docx_edit call');
    expect(doc.session.encodeState()).toEqual(state);
    expect(doc.list().version).toBe(version);
  });
}

test('several delete_comment operations can share a batch with a text format edit', async () => {
  const doc = await open(await fixture());
  const match = doc.grep({ query: 'Revenue grew strongly' }).matches[0];
  const receipt = await doc.edit({ version: doc.list().version, edits: [
    { op: 'delete_comment', comment: '1' },
    { op: 'format_text', match: match.match, italic: true },
    { op: 'delete_comment', comment: '0' },
  ] });
  expect(receipt.results[0]).toMatchObject({ comment: '1', deleted: ['1'] });
  expect(receipt.results[2]).toMatchObject({ comment: '0', deleted: ['0'] });
  const reopened = await open(await doc.export());
  const paragraph = reopened.list().items.find(paragraph => paragraph.text.startsWith('Revenue grew'))!;
  expect(reopened.read(paragraph.ref).runs.some(run => run.formatting.italic)).toBe(true);
  expect(reopened.list({ story: 'comments' }).items.map(comment => comment.id)).toEqual(['2']);
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
