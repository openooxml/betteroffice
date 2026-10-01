import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { attemptSelectiveSave, parseDocx, repackDocx } from '.';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { unzipContainer } from './wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsSession } from '../yrs';
import { getCommentText } from '../utils/comments';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"`;
const RICH = '<w:comment w:author="Rich reviewer" w:id="0" w:initials="R">\n' +
  '<w:p w14:paraId="10000001"><w:hyperlink r:id="link"><w:r><w:t>Linked</w:t></w:r></w:hyperlink>' +
  '<w:r><w:rPr><w:color w:val="CC0000"/><w:u w:val="single"/><w:sz w:val="28"/>' +
  '<w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve"> colorful</w:t></w:r></w:p>\n' +
  '<w:p w14:paraId="10000002"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>\n' +
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>' +
  '<w:p w14:paraId="10000003"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n' +
  '<w:p w14:paraId="10000005"><w:r><w:t>After</w:t></w:r></w:p>\n</w:comment>';
const COMMENT_RELS = `<Relationships xmlns="${RELS}">\n` +
  `  <Relationship TargetMode='External' Target='https://example.com/synthetic' Id='link' Type='${R}/hyperlink'/>\n` +
  '</Relationships>';
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

function fixture(): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  add('_rels/.rels', `<Relationships xmlns="${RELS}"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', `<Relationships xmlns="${RELS}"><Relationship Id="comments" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
  add('word/document.xml', `<w:document ${NS}><w:body><w:p w14:paraId="00000001">` +
    '<w:commentRangeStart w:id="0"/><w:r><w:t>Rich anchor</w:t></w:r><w:commentRangeEnd w:id="0"/>' +
    '<w:r><w:commentReference w:id="0"/></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>Plain anchor</w:t></w:r>' +
    '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>');
  add('word/comments.xml', `<w:comments ${NS}>\n${RICH}\n` +
    '<w:comment w:id="1" w:author="Plain reviewer"><w:p w14:paraId="10000004"><w:r><w:t>Plain</w:t></w:r></w:p></w:comment>\n</w:comments>');
  add('word/_rels/comments.xml.rels', COMMENT_RELS);
  return rezipPartsToArrayBuffer(parts);
}

function assertRich(bytes: ArrayBuffer | Uint8Array): void {
  const parts = unzipContainer(new Uint8Array(bytes));
  const xml = new TextDecoder().decode(parts['word/comments.xml']);
  const rich = xml.match(/<w:comment\b[^>]*w:id="0"[^>]*>[\s\S]*?<\/w:comment>/)?.[0];
  expect(toBytes(rich ?? '')).toEqual(toBytes(RICH));
  expect(parts['word/_rels/comments.xml.rels']).toEqual(toBytes(COMMENT_RELS));
}

async function open(): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 7 });
  sessions.push(session);
  session.openDocx(new Uint8Array(fixture()), true);
  return session;
}

test('parseDocx → repackDocx preserves rich comment XML and relationship bytes', async () => {
  const source = fixture();
  const document = await parseDocx(source, { preloadFonts: false });
  const saved = await repackDocx(document);
  assertRich(saved);
  expect(unzipContainer(new Uint8Array(saved))['word/comments.xml'])
    .toEqual(unzipContainer(new Uint8Array(source))['word/comments.xml']);
});

test('repackDocx saves an edited comment and the other comment text', async () => {
  const document = await parseDocx(fixture(), { preloadFonts: false });
  document.package.document.comments![1]!.content[0]!.content = [
    { type: 'run', content: [{ type: 'text', text: 'Edited plain comment' }] },
  ];
  const saved = await parseDocx(await repackDocx(document), { preloadFonts: false });
  const comments = saved.package.document.comments!;
  expect(getCommentText(comments.find(({ id }) => id === 1)!.content)).toBe('Edited plain comment');
  expect(getCommentText(comments.find(({ id }) => id === 0)!.content)).toContain('colorful');
});

test('repackDocx saves a comment deletion and the other comment text', async () => {
  const document = await parseDocx(fixture(), { preloadFonts: false });
  document.package.document.comments!.splice(1, 1);
  const saved = await parseDocx(await repackDocx(document), { preloadFonts: false });
  const comments = saved.package.document.comments!;
  expect(comments.map(({ id }) => id)).toEqual([0]);
  expect(getCommentText(comments[0]!.content)).toContain('colorful');
});

test('selective body save preserves rich comment XML and relationship bytes', async () => {
  const source = fixture();
  const document = await parseDocx(source, { preloadFonts: false });
  const paragraph = document.package.document.content[0]!;
  if (paragraph.type !== 'paragraph') throw new Error('Missing synthetic body paragraph');
  paragraph.content.unshift({ type: 'run', content: [{ type: 'text', text: 'Edited body ' }] });
  const saved = await attemptSelectiveSave(document, source, {
    changedParaIds: new Set(['00000001']), structuralChange: false, hasUntrackedChanges: false,
  });
  expect(saved).not.toBeNull();
  assertRich(saved!);
  expect(unzipContainer(new Uint8Array(saved!))['word/comments.xml'])
    .toEqual(unzipContainer(new Uint8Array(source))['word/comments.xml']);
  expect(new TextDecoder().decode(unzipContainer(new Uint8Array(saved!))['word/document.xml']))
    .toContain('Edited body ');
});

test('saveYrsDocx saves a session comment edit and the other comment text', async () => {
  const session = await open();
  const ranges = session.resolveComment('1').map(({ start, end }) => [start, end] as const);
  session.applyRawOps('body', [{ op: 'setComment', id: '1', ranges, body: 'Edited in session' }]);
  const saved = await parseDocx((await saveYrsDocx(session)).bytes, { preloadFonts: false });
  const comments = saved.package.document.comments!;
  expect(getCommentText(comments.find(({ id }) => id === 1)!.content)).toBe('Edited in session');
  expect(getCommentText(comments.find(({ id }) => id === 0)!.content)).toContain('colorful');
});

test('saveYrsDocx preserves comment bytes after only document body text changes', async () => {
  const session = await open();
  const [paragraph] = session.paragraphs('body');
  session.insertText({ story: 'body', paraId: paragraph!.paraId, offset: 0 }, 'Edited body ');
  const saved = (await saveYrsDocx(session)).bytes;
  assertRich(saved);
  const source = unzipContainer(new Uint8Array(fixture()));
  expect(unzipContainer(saved)['word/comments.xml']).toEqual(source['word/comments.xml']);
});
