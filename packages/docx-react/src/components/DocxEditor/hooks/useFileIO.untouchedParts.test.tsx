import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { posix, resolve } from 'node:path';
import { createRef } from 'react';
import { parseDocx, repackDocx } from '@betteroffice/docx/docx';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import type { Comment } from '@betteroffice/docx/types/content';
import { getCommentText } from '@betteroffice/docx/utils/comments';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { DocxEditor, type DocxEditorRef } from '../../../index';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');
const quiet = { error: console.error, warn: console.warn };
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`;
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
const PNG = new Uint8Array(Buffer.from(PNG_BASE64, 'base64'));
const SECTION =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const COMMENT_PARA_ID = '10000001';

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

type Fixture = { bytes: ArrayBuffer; parts: Map<string, Uint8Array> };
type FixtureOptions = {
  comments?: boolean;
  header?: boolean;
  footnote?: boolean;
  image?: boolean;
  chunk?: boolean;
};
type Paragraph = (content: string, properties?: string) => string;

function run(text: string): string {
  return `<w:r><w:t>${text}</w:t></w:r>`;
}

function paragraph(id: string, content: string, properties = ''): string {
  return `<w:p w14:paraId="${id}" w14:textId="77777777">${properties}${content}</w:p>`;
}

function relationship(id: string, type: string, target: string, external = false): string {
  return `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
}

function fixture(body: (p: Paragraph) => string, options: FixtureOptions = {}): Fixture {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(DECLARATION + xml));
  const overrides = [`<Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/>`];
  const relationships: string[] = [];
  const override = (name: string, kind: string) =>
    overrides.push(`<Override PartName="/word/${name}" ContentType="${WORD}.${kind}+xml"/>`);
  let references = '';
  if (options.comments) {
    override('comments.xml', 'comments');
    override('commentsExtended.xml', 'commentsExtended');
    relationships.push(relationship('comments', 'comments', 'comments.xml'));
    relationships.push(
      '<Relationship Id="commentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>'
    );
    add(
      'word/comments.xml',
      `<w:comments ${NS}><w:comment w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z" w:initials="A">${paragraph(COMMENT_PARA_ID, `<w:hyperlink w:anchor="target">${run('Important')}</w:hyperlink>`)}</w:comment></w:comments>`
    );
    add(
      'word/commentsExtended.xml',
      `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="${COMMENT_PARA_ID}" w15:done="0"/></w15:commentsEx>`
    );
  }
  if (options.header) {
    override('header1.xml', 'header');
    relationships.push(relationship('header', 'header', 'header1.xml'));
    references = '<w:headerReference w:type="default" r:id="header"/>';
    add(
      'word/header1.xml',
      `<w:hdr ${NS}>\n  <w:p w14:textId="77777777" w14:paraId="20000001"><w:permStart w:edGrp="everyone" w:id="9"/>${run('Header')}<w:permEnd w:id="9"/></w:p>\n</w:hdr>`
    );
  }
  if (options.footnote) {
    override('footnotes.xml', 'footnotes');
    relationships.push(relationship('notes', 'footnotes', 'footnotes.xml'));
    add(
      'word/footnotes.xml',
      `<w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1">${paragraph('30000001', '<w:r><w:separator/></w:r>')}</w:footnote><w:footnote w:type="continuationSeparator" w:id="0">${paragraph('30000002', '<w:r><w:continuationSeparator/></w:r>')}</w:footnote><w:footnote w:id="1">${paragraph('30000003', `<w:r><w:footnoteRef/></w:r>${run('Synthetic note')}`)}</w:footnote></w:footnotes>`
    );
  }
  if (options.image) {
    relationships.push(relationship('picture', 'image', 'media/pixel.png'));
    relationships.push(relationship('link', 'hyperlink', 'https://example.com/synthetic', true));
    parts.set('word/media/pixel.png', PNG);
  }
  if (options.chunk) {
    relationships.push(relationship('chunk', 'aFChunk', 'chunk.html'));
    parts.set('word/chunk.html', toBytes('<!doctype html><html><body><p>Synthetic chunk</p></body></html>'));
  }
  add(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${options.image ? '<Default Extension="png" ContentType="image/png"/>' : ''}${options.chunk ? '<Default Extension="html" ContentType="text/html"/>' : ''}${overrides.join('')}</Types>`
  );
  add('_rels/.rels', `<Relationships xmlns="${RELS}">${relationship('office', 'officeDocument', 'word/document.xml')}</Relationships>`);
  add('word/_rels/document.xml.rels', `<Relationships xmlns="${RELS}">${relationships.join('')}</Relationships>`);
  let ordinal = 0;
  const p: Paragraph = (content, properties) =>
    paragraph((++ordinal).toString(16).padStart(8, '0').toUpperCase(), content, properties);
  add('word/document.xml', `<w:document ${NS}><w:body>${body(p)}<w:sectPr>${references}${SECTION}</w:sectPr></w:body></w:document>`);
  return { bytes: rezipPartsToArrayBuffer(parts), parts };
}

function commentedFixture(options: FixtureOptions = {}): Fixture {
  return fixture(
    (p) => p(
      '<w:bookmarkStart w:id="7" w:name="target"/><w:commentRangeStart w:id="1"/>' +
      run('Body text') +
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:bookmarkEnd w:id="7"/>' +
      (options.footnote ? '<w:r><w:footnoteReference w:id="1"/></w:r>' : '')
    ),
    { ...options, comments: true }
  );
}

function xmlPart(parts: ReturnType<typeof unzipContainer>, name: string): string {
  const bytes = parts[name];
  if (!bytes) throw new Error(`Missing package part: ${name}`);
  return new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
}

function expectSameXml(actual: string, expected: string, name: string): void {
  if (actual === expected) return;
  let offset = 0;
  while (offset < Math.min(actual.length, expected.length) && actual[offset] === expected[offset]) {
    offset += 1;
  }
  const start = Math.max(0, offset - 35);
  throw new Error(
    `${name}: first difference at offset ${offset} (source ${expected.length}, saved ${actual.length} chars)\n` +
    `source: ${JSON.stringify(expected.slice(start, offset + 75))}\n` +
    `saved:  ${JSON.stringify(actual.slice(start, offset + 75))}`
  );
}

function expectUnchanged(source: Fixture, saved: ArrayBuffer, names: string[]): void {
  const parts = unzipContainer(new Uint8Array(saved));
  for (const name of names) {
    const expected = source.parts.get(name);
    if (!expected) throw new Error(`Missing source part: ${name}`);
    expectSameXml(xmlPart(parts, name), new TextDecoder().decode(expected), name);
  }
}

async function until(done: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

async function mount(bytes: ArrayBuffer) {
  const ref = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const onError = (error: Error) => errors.push(error);
  const view = render(
    <DocxEditor ref={ref} documentBuffer={bytes} onError={onError} downloadOnSave={false} />
  );
  await until(
    () => ref.current?.commands.getState('save').enabled === true &&
      !!ref.current.getEditorRef()?.getYrsSession()
  );
  if (unzipContainer(new Uint8Array(bytes))['word/comments.xml']) {
    await until(() => (ref.current?.getComments().length ?? 0) > 0);
  }
  expect(errors.map(({ message }) => message)).toEqual([]);
  return {
    ref,
    view,
    async save(): Promise<ArrayBuffer> {
      let saved: ArrayBuffer | null = null;
      await act(async () => {
        saved = await ref.current!.save();
      });
      expect(errors.map(({ message }) => message)).toEqual([]);
      expect(saved).not.toBeNull();
      return saved!;
    },
    setComments(comments: Comment[]) {
      view.rerender(
        <DocxEditor
          ref={ref}
          documentBuffer={bytes}
          comments={comments}
          onError={onError}
          downloadOnSave={false}
        />
      );
    },
  };
}

async function typeBody(editor: Awaited<ReturnType<typeof mount>>, text: string): Promise<void> {
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const [first] = session.paragraphs('body');
  await act(async () => {
    session.insertText({ story: 'body', paraId: first!.paraId, offset: 0 }, text);
    paged.syncYrsInputState(true, ['body']);
  });
}

async function reopened(buffer: ArrayBuffer) {
  return parseDocx(new Uint8Array(buffer), { preloadFonts: false });
}

function markers(xml: string, id: number): string[] {
  return [...xml.matchAll(new RegExp(`<w:comment(RangeStart|RangeEnd|Reference)\\b[^>]*\\bw:id="${id}"[^>]*/>`, 'g'))].map((match) => match[1]!);
}

function drawing(): string {
  return `<w:r><w:drawing><wp:inline xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="9525" cy="9525"/><wp:docPr id="42" name="Link"><a:hlinkClick xmlns:a="${A}" r:id="link"/></wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

test('no-edit React save preserves rich comments, body, header and footnotes byte-for-byte', async () => {
  const source = commentedFixture({ header: true, footnote: true });
  const editor = await mount(source.bytes);
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('Important');
  expectUnchanged(source, saved, ['word/comments.xml', 'word/document.xml', 'word/header1.xml', 'word/footnotes.xml']);
});

test('no-edit React save preserves a field nested inside a hyperlink byte-for-byte', async () => {
  const source = fixture((p) =>
    p('<w:hyperlink w:anchor="target"><w:fldSimple w:instr="DATE">' + run('October') + '</w:fldSimple></w:hyperlink>') +
    p('<w:bookmarkStart w:id="7" w:name="target"/>' + run('Target') + '<w:bookmarkEnd w:id="7"/>')
  );
  const editor = await mount(source.bytes);
  expectUnchanged(source, await editor.save(), ['word/document.xml']);
});

test('no-edit React save preserves altChunk and body AlternateContent byte-for-byte', async () => {
  const source = fixture((p) =>
    p(run('Before chunk')) + '<w:altChunk r:id="chunk"/>' +
    `<mc:AlternateContent><mc:Choice Requires="w14">${p(run('A'))}</mc:Choice><mc:Fallback>${p('')}</mc:Fallback></mc:AlternateContent>` +
    p(run('After chunk')),
    { chunk: true }
  );
  const editor = await mount(source.bytes);
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  const rels = xmlPart(parts, 'word/_rels/document.xml.rels');
  const relation = new DOMParser().parseFromString(rels, 'application/xml').getElementsByTagNameNS(RELS, 'Relationship');
  const chunk = Array.from(relation).find((entry) => entry.getAttribute('Id') === 'chunk');
  expect(chunk?.getAttribute('Type')).toBe(`${R}/aFChunk`);
  expect(chunk?.getAttribute('Target')).toBe('chunk.html');
  expectUnchanged(source, saved, ['word/document.xml', 'word/chunk.html']);
});

test('no-edit React save preserves text in a vertical-merge continuation cell byte-for-byte', async () => {
  const source = fixture((p) =>
    p(run('Table')) +
    '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>' + p(run('Top')) +
    '</w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + p(run('Right one')) + '</w:tc></w:tr>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr>' + p(run('LOST')) +
    '</w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + p(run('Right two')) + '</w:tc></w:tr></w:tbl>' + p(run('After table'))
  );
  const editor = await mount(source.bytes);
  expectUnchanged(source, await editor.save(), ['word/document.xml']);
});

test('no-edit React save preserves permission ranges byte-for-byte', async () => {
  const source = fixture((p) => p('<w:permStart w:id="3" w:edGrp="everyone"/>' + run('Editable') + '<w:permEnd w:id="3"/>'));
  const editor = await mount(source.bytes);
  expectUnchanged(source, await editor.save(), ['word/document.xml']);
});

test('no-edit React save preserves an image hyperlink and its relationship byte-for-byte', async () => {
  const source = fixture((p) => p(run('Linked image') + drawing()), { image: true });
  const editor = await mount(source.bytes);
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  const rels = new DOMParser().parseFromString(xmlPart(parts, 'word/_rels/document.xml.rels'), 'application/xml');
  const link = Array.from(rels.getElementsByTagNameNS(RELS, 'Relationship')).find((entry) => entry.getAttribute('Id') === 'link');
  expect(link?.getAttribute('Type')).toBe(`${R}/hyperlink`);
  expect(link?.getAttribute('Target')).toBe('https://example.com/synthetic');
  expect(link?.getAttribute('TargetMode')).toBe('External');
  expect(Array.from(parts['word/media/pixel.png'] ?? [])).toEqual(Array.from(PNG));
  expectUnchanged(source, saved, ['word/document.xml']);
});

test('no-edit React save preserves empty paragraph section properties byte-for-byte', async () => {
  const source = fixture((p) => p(run('First section'), '<w:pPr><w:sectPr/></w:pPr>') + p(run('Second section')));
  const editor = await mount(source.bytes);
  expectUnchanged(source, await editor.save(), ['word/document.xml']);
});

test('a React body edit saves its text and preserves untouched header and rich comments', async () => {
  const source = commentedFixture({ header: true });
  const editor = await mount(source.bytes);
  await typeBody(editor, 'Typed in body ');
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/document.xml') === new TextDecoder().decode(source.parts.get('word/document.xml'))).toBe(false);
  editor.view.unmount();
  const second = await mount(saved);
  const session = second.ref.current!.getEditorRef()!.getYrsSession()!;
  expect(session.paragraphs('body')[0]!.text).toContain('Typed in body Body text');
  expectUnchanged(source, saved, ['word/header1.xml', 'word/comments.xml']);
});

test('a comment text edited in React is saved and reopened', async () => {
  const source = commentedFixture();
  const editor = await mount(source.bytes);
  const comments = editor.ref.current!.getComments().map((comment) => ({
    ...comment,
    content: [{ type: 'paragraph' as const, paraId: COMMENT_PARA_ID, content: [{ type: 'run' as const, content: [{ type: 'text' as const, text: 'Edited by host' }] }] }],
    blockContent: undefined,
  }));
  await act(async () => editor.setComments(comments));
  const saved = await editor.save();
  const xml = xmlPart(unzipContainer(new Uint8Array(saved)), 'word/comments.xml');
  expect(xml).toContain('Edited by host');
  expect(xml).not.toContain('Important');
  const document = await reopened(saved);
  const [comment] = document.package.document.comments!;
  expect(comment?.id).toBe(1);
  expect(getCommentText(comment!.content)).toBe('Edited by host');
});

test('a comment resolved in React is saved in commentsExtended and reopened', async () => {
  const editor = await mount(commentedFixture().bytes);
  await act(async () => editor.ref.current!.resolveComment(1));
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('w:id="1"');
  const extended = new DOMParser().parseFromString(xmlPart(parts, 'word/commentsExtended.xml'), 'application/xml');
  const comment = Array.from(extended.getElementsByTagNameNS(W15, 'commentEx')).find((entry) => entry.getAttributeNS(W15, 'paraId') === COMMENT_PARA_ID);
  expect(comment?.getAttributeNS(W15, 'done')).toBe('1');
  expect((await reopened(saved)).package.document.comments?.find(({ id }) => id === 1)?.done).toBe(true);
});

test('a React reply saves its text, thread metadata and body range markers', async () => {
  const editor = await mount(commentedFixture().bytes);
  let replyId: number | null = null;
  await act(async () => {
    replyId = editor.ref.current!.replyToComment(1, 'Host reply', 'B');
  });
  expect(replyId).not.toBeNull();
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('Host reply');
  const documentXml = xmlPart(parts, 'word/document.xml');
  expect(markers(documentXml, replyId!)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
  expect(markers(documentXml, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
  const document = await reopened(saved);
  const reply = document.package.document.comments?.find(({ id }) => id === replyId);
  const parent = document.package.document.comments?.find(({ id }) => id === 1);
  expect(reply?.parentId).toBe(1);
  expect(reply?.author).toBe('B');
  expect(getCommentText(reply!.content)).toBe('Host reply');
  expect(reply?.paraId).toBeTruthy();
  expect(parent?.paraId).toBeTruthy();
  const extended = new DOMParser().parseFromString(xmlPart(parts, 'word/commentsExtended.xml'), 'application/xml');
  const replyEx = Array.from(extended.getElementsByTagNameNS(W15, 'commentEx')).find((entry) => entry.getAttributeNS(W15, 'paraId') === reply?.paraId);
  expect(replyEx?.getAttributeNS(W15, 'paraIdParent')).toBe(parent?.paraId);
});

test('a React comment added on selected body text saves its body and range', async () => {
  const editor = await mount(commentedFixture().bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const [first] = session.paragraphs('body');
  let id: number | null = null;
  await act(async () => {
    session.setSelection({ story: 'body', paraId: first!.paraId, offset: 0 }, { story: 'body', paraId: first!.paraId, offset: 4 });
    paged.syncYrsInputState(false);
    id = editor.ref.current!.addComment({ paraId: first!.paraId, search: 'Body', text: 'New host comment', author: 'C' });
  });
  expect(id).not.toBeNull();
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('New host comment');
  const xml = xmlPart(parts, 'word/document.xml');
  expect(markers(xml, id!)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
  const covered = xml.split(`<w:commentRangeStart w:id="${id}"`)[1]?.split('<w:commentRangeEnd')[0]?.replace(/^[^>]*>/, '').replace(/<[^>]+>/g, '');
  expect(covered).toBe('Body');
  const document = await reopened(saved);
  const added = document.package.document.comments?.find((comment) => comment.id === id);
  expect(added?.author).toBe('C');
  expect(getCommentText(added!.content)).toBe('New host comment');
});

test('two consecutive React saves retain a body edit in valid packages', async () => {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  await typeBody(editor, 'Saved twice ');
  const first = await editor.save();
  const second = await editor.save();
  for (const buffer of [first, second]) {
    const parts = unzipContainer(new Uint8Array(buffer));
    expect(xmlPart(parts, '[Content_Types].xml')).toContain('/word/document.xml');
    expect(xmlPart(parts, '_rels/.rels')).toContain('officeDocument');
    expect(xmlPart(parts, 'word/document.xml')).toContain('Saved twice ');
    const document = await reopened(buffer);
    expect(document.package.document.content.length).toBeGreaterThan(0);
  }
  editor.view.unmount();
  const secondEditor = await mount(second);
  expect(secondEditor.ref.current!.getEditorRef()!.getYrsSession()!.paragraphs('body')[0]!.text).toBe('Saved twice Body text');
});

test('two React saves retain an inserted image and a resolvable media relationship', async () => {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const [first] = session.paragraphs('body');
  await act(async () => {
    session.setSelection({ story: 'body', paraId: first!.paraId, offset: 0 });
    paged.syncYrsInputState(false);
    expect(paged.applyYrsCommand({ type: 'insertImage', image: { src: `data:image/png;base64,${PNG_BASE64}`, width: 1, height: 1, alt: 'Synthetic pixel' } })).toBe(true);
  });
  const firstSave = await editor.save();
  const secondSave = await editor.save();
  for (const buffer of [firstSave, secondSave]) {
    const parts = unzipContainer(new Uint8Array(buffer));
    const xml = new DOMParser().parseFromString(xmlPart(parts, 'word/document.xml'), 'application/xml');
    const embed = xml.getElementsByTagNameNS(A, 'blip')[0]?.getAttributeNS(R, 'embed');
    expect(embed).toBeTruthy();
    const rels = new DOMParser().parseFromString(xmlPart(parts, 'word/_rels/document.xml.rels'), 'application/xml');
    const relationship = Array.from(rels.getElementsByTagNameNS(RELS, 'Relationship')).find((entry) => entry.getAttribute('Id') === embed);
    expect(relationship?.getAttribute('Type')).toBe(`${R}/image`);
    const target = relationship?.getAttribute('Target');
    expect(target).toBeTruthy();
    const name = target!.startsWith('/') ? target!.slice(1) : posix.normalize(posix.join('word', target!));
    expect(name.startsWith('word/media/')).toBe(true);
    expect(Array.from(parts[name] ?? [])).toEqual(Array.from(PNG));
    expect(xmlPart(parts, '[Content_Types].xml')).toContain('image/png');
    await reopened(buffer);
  }
});

test('React save after undo restores source document bytes instead of the previous save', async () => {
  const synthetic = fixture((p) => p(run('Body text')));
  const source = await repackDocx(await reopened(synthetic.bytes));
  const sourceXml = xmlPart(unzipContainer(new Uint8Array(source)), 'word/document.xml');
  const editor = await mount(source);
  await typeBody(editor, 'Undo this ');
  const edited = await editor.save();
  expect(xmlPart(unzipContainer(new Uint8Array(edited)), 'word/document.xml')).toContain('Undo this ');
  await act(async () => {
    expect(editor.ref.current!.getEditorRef()!.undo()).toBe(true);
  });
  expect(editor.ref.current!.getEditorRef()!.getYrsSession()!.paragraphs('body')[0]!.text).toBe('Body text');
  const undone = await editor.save();
  expectSameXml(xmlPart(unzipContainer(new Uint8Array(undone)), 'word/document.xml'), sourceXml, 'word/document.xml after undo');
  await reopened(undone);
});

test('a host page-setup change is saved even when body text is untouched', async () => {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  await act(async () => {
    const outcome = await editor.ref.current!.commands.execute('pageSetup', null);
    expect(outcome.ok).toBe(true);
  });
  await until(() => !!editor.view.queryByRole('dialog'));
  const dialog = editor.view.getByRole('dialog');
  const [topMargin] = within(dialog).getAllByRole('spinbutton');
  await act(async () => fireEvent.change(topMargin!, { target: { value: '2' } }));
  await act(async () => fireEvent.click(within(dialog).getByRole('button', { name: 'Apply' })));
  await until(() => !editor.view.queryByRole('dialog'));
  const saved = await editor.save();
  const xml = new DOMParser().parseFromString(xmlPart(unzipContainer(new Uint8Array(saved)), 'word/document.xml'), 'application/xml');
  expect(xml.getElementsByTagNameNS(W, 'pgMar')[0]?.getAttributeNS(W, 'top')).toBe('2880');
  expect((await reopened(saved)).package.document.finalSectionProperties?.marginTop).toBe(2880);
});
