import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, useState } from 'react';
import { parseDocx, repackDocx } from '@betteroffice/docx/docx';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document as DocxDocument } from '@betteroffice/docx/types/document';
import { getCommentText } from '@betteroffice/docx/utils/comments';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { DocxEditor, type DocxEditorRef } from '../../../index';
import type { PagedEditorRef } from '../PagedEditor';
import type { PartEditTarget } from '../partEdit';
import { setupWorkerEngine } from '../__fixtures__/workerEngine';
import { useFileIO } from './useFileIO';
import { useHeaderFooterEditing } from './useHeaderFooterEditing';
import * as headerFooterEditing from './useHeaderFooterEditing';

const { act, cleanup, fireEvent, render, renderHook, within } = await import('@testing-library/react');
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

const RICH_COMMENT = '<w:comment w:author="Rich reviewer" w:id="0" w:initials="R">\n' +
  '<w:p w14:paraId="10000010"><w:hyperlink r:id="commentLink">' + run('Linked') + '</w:hyperlink>' +
  '<w:r><w:rPr><w:color w:val="CC0000"/><w:u w:val="single"/><w:sz w:val="28"/>' +
  '<w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve"> colorful</w:t></w:r></w:p>\n' +
  '<w:p w14:paraId="10000011"><w:pPr><w:jc w:val="center"/></w:pPr>' + run('Second') + '</w:p>\n' +
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>' +
  paragraph('10000012', run('Cell')) + '</w:tc></w:tr></w:tbl>\n' +
  paragraph('10000013', run('After')) + '\n</w:comment>';

function richCommentsFixture(): Fixture {
  const source = fixture((p) =>
    p('<w:commentRangeStart w:id="0"/>' + run('Rich anchor') + '<w:commentRangeEnd w:id="0"/>' +
      '<w:r><w:commentReference w:id="0"/></w:r>') +
    p('<w:commentRangeStart w:id="1"/>' + run('Plain anchor') + '<w:commentRangeEnd w:id="1"/>' +
      '<w:r><w:commentReference w:id="1"/></w:r>'),
    { comments: true }
  );
  source.parts.set('word/comments.xml', toBytes(DECLARATION + `<w:comments ${NS}>\n${RICH_COMMENT}\n` +
    `<w:comment w:id="1" w:author="Plain reviewer">${paragraph(COMMENT_PARA_ID, run('Plain'))}</w:comment>\n</w:comments>`));
  source.parts.set('word/commentsExtended.xml', toBytes(DECLARATION +
    `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="10000011" w15:done="0"/>` +
    `<w15:commentEx w15:paraId="${COMMENT_PARA_ID}" w15:done="0"/></w15:commentsEx>`));
  source.parts.set('word/_rels/comments.xml.rels', toBytes(DECLARATION +
    `<Relationships xmlns="${RELS}">\n  ${relationship('commentLink', 'hyperlink', 'https://example.com/synthetic', true)}\n</Relationships>`));
  source.bytes = rezipPartsToArrayBuffer(source.parts);
  return source;
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

function xmlElements(root: Document | Element, namespace: string, localName: string): Element[] {
  return Array.from(root.getElementsByTagName('*')).filter(
    (element) => element.namespaceURI === namespace && element.localName === localName
  );
}

function savedCommentParaId(parts: ReturnType<typeof unzipContainer>, id: number): string {
  const xml = new DOMParser().parseFromString(xmlPart(parts, 'word/comments.xml'), 'application/xml');
  const comment = xmlElements(xml, W, 'comment').find((entry) => entry.getAttribute('w:id') === String(id));
  expect(comment).toBeDefined();
  const paraId = xmlElements(comment!, W, 'p').at(-1)?.getAttribute('w14:paraId');
  expect(paraId).toBeTruthy();
  return paraId!;
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

test('React saves a plain comment edit and the other comment text', async () => {
  const source = richCommentsFixture();
  const editor = await mount(source.bytes);
  const comments = editor.ref.current!.getComments().map((comment) => comment.id === 1 ? {
    ...comment,
    content: [{ type: 'paragraph' as const, paraId: COMMENT_PARA_ID, content: [
      { type: 'run' as const, content: [{ type: 'text' as const, text: 'Edited plain comment' }] },
    ] }],
  } : comment);
  await act(async () => editor.setComments(comments));
  const saved = await editor.save();
  const reopenedComments = (await reopened(saved)).package.document.comments!;
  expect(getCommentText(reopenedComments.find(({ id }) => id === 1)!.content)).toBe('Edited plain comment');
  expect(getCommentText(reopenedComments.find(({ id }) => id === 0)!.content)).toContain('colorful');
});

test('React saves a plain comment deletion and the other comment text', async () => {
  const source = richCommentsFixture();
  const editor = await mount(source.bytes);
  await act(async () => editor.setComments(editor.ref.current!.getComments().filter(({ id }) => id !== 1)));
  await until(() => editor.ref.current!.getComments().length === 1);
  const saved = await editor.save();
  const comments = (await reopened(saved)).package.document.comments!;
  expect(comments.map(({ id }) => id)).toEqual([0]);
  expect(getCommentText(comments[0]!.content)).toContain('colorful');
});

test('React saves an added comment and the existing comment text', async () => {
  const source = richCommentsFixture();
  const editor = await mount(source.bytes);
  const added: Comment = {
    id: 2,
    author: 'New reviewer',
    content: [{ type: 'paragraph', content: [
      { type: 'run', content: [{ type: 'text', text: 'Added comment' }] },
    ] }],
  };
  await act(async () => editor.setComments([...editor.ref.current!.getComments(), added]));
  const saved = await editor.save();
  const comments = (await reopened(saved)).package.document.comments!;
  expect(getCommentText(comments.find(({ id }) => id === 2)!.content)).toBe('Added comment');
  expect(getCommentText(comments.find(({ id }) => id === 0)!.content)).toContain('colorful');
  expect(getCommentText(comments.find(({ id }) => id === 1)!.content)).toBe('Plain');
});

test('React saves a resolved comment and the other comment text', async () => {
  const source = richCommentsFixture();
  const editor = await mount(source.bytes);
  await act(async () => editor.ref.current!.resolveComment(0));
  const saved = await editor.save();
  const comments = (await reopened(saved)).package.document.comments!;
  expect(comments.find(({ id }) => id === 0)?.done).toBe(true);
  expect(getCommentText(comments.find(({ id }) => id === 0)!.content)).toContain('colorful');
  expect(getCommentText(comments.find(({ id }) => id === 1)!.content)).toBe('Plain');
});

test('React save after only body text changes preserves both comment elements and rels', async () => {
  const source = richCommentsFixture();
  const editor = await mount(source.bytes);
  await typeBody(editor, 'Edited body ');
  expectUnchanged(source, await editor.save(), ['word/comments.xml', 'word/_rels/comments.xml.rels']);
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

test('a package part the host replaced in originalBuffer is saved', async () => {
  const source = fixture((p) => p(run('Linked image') + drawing()), { image: true });
  const editor = await mount(source.bytes);
  const replaced = new Uint8Array([...PNG, 0]);
  const document = editor.ref.current!.getDocument()!;
  const parts = unzipContainer(new Uint8Array(document.originalBuffer!));
  parts['word/media/pixel.png'] = replaced;
  document.originalBuffer = rezipPartsToArrayBuffer(new Map(Object.entries(parts)));
  const saved = unzipContainer(new Uint8Array(await editor.save()));
  expect(Array.from(saved['word/media/pixel.png'] ?? [])).toEqual(Array.from(replaced));
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

test('an in-place host comment edit keeps its text and resolved state on save', async () => {
  const editor = await mount(commentedFixture().bytes);
  const [comment] = editor.ref.current!.getComments();
  comment!.content = [{ type: 'paragraph', paraId: COMMENT_PARA_ID, content: [{ type: 'run', content: [{ type: 'text', text: 'Edited in place by host' }] }] }];
  comment!.blockContent = undefined;
  comment!.done = true;
  await act(async () => editor.setComments([comment!]));
  expect(editor.ref.current!.getComments()[0]).toBe(comment);
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('Edited in place by host');
  expect(xmlPart(parts, 'word/comments.xml')).not.toContain('Important');
  const paraId = savedCommentParaId(parts, 1);
  const extended = new DOMParser().parseFromString(xmlPart(parts, 'word/commentsExtended.xml'), 'application/xml');
  const entry = xmlElements(extended, W15, 'commentEx').find((element) => element.getAttribute('w15:paraId') === paraId);
  expect(entry?.getAttribute('w15:done')).toBe('1');
  const [reopenedComment] = (await reopened(saved)).package.document.comments!;
  expect(getCommentText(reopenedComment!.content)).toBe('Edited in place by host');
  expect(reopenedComment?.done).toBe(true);
});

function twoCommentFixture(): ArrayBuffer {
  const source = commentedFixture();
  const secondParaId = '10000002';
  const extend = (name: string, closing: string, xml: string) => {
    const original = new TextDecoder().decode(source.parts.get(name)!);
    source.parts.set(name, toBytes(original.replace(closing, xml + closing)));
  };
  extend('word/comments.xml', '</w:comments>', `<w:comment w:id="2" w:author="B" w:date="2024-01-01T00:00:00Z" w:initials="B">${paragraph(secondParaId, run('Second comment'))}</w:comment>`);
  extend('word/commentsExtended.xml', '</w15:commentsEx>', `<w15:commentEx w15:paraId="${secondParaId}" w15:done="0"/>`);
  extend('word/document.xml', '<w:sectPr>', paragraph('00000002', '<w:commentRangeStart w:id="2"/>' + run('Second body text') + '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>'));
  return rezipPartsToArrayBuffer(source.parts);
}

async function clickCommentAfterHeaderReturn(
  mountEditor: (bytes: ArrayBuffer) => Promise<Pick<Awaited<ReturnType<typeof mount>>, 'ref' | 'view'>>,
  moveBodyCaret = true
) {
  const useEditing = useHeaderFooterEditing;
  let editing!: ReturnType<typeof useHeaderFooterEditing>;
  const captureEditing = spyOn(headerFooterEditing, 'useHeaderFooterEditing').mockImplementation((options) => {
    editing = useEditing(options);
    return editing;
  });
  try {
    const source = fixture((p) =>
      p('<w:commentRangeStart w:id="1"/>' + run('Body text') + '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>') +
      p(run('Second paragraph')),
      { comments: true, header: true }
    );
    const editor = await mountEditor(source.bytes);
    await editor.ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
    await act(async () => { await editor.ref.current!.flushPendingInput(); });
    const paged = editor.ref.current!.getEditorRef()!;
    const session = paged.getYrsSession()!;
    const [first, second] = session.paragraphs('body');
    if (!editor.ref.current!.commands.getState('commentsSidebar').active) {
      await act(async () => {
        expect((await editor.ref.current!.commands.execute('commentsSidebar', null)).ok).toBe(true);
      });
    }
    await until(() => !!editor.view.container.querySelector('.docx-unified-sidebar .docx-comment-card[data-comment-id="1"]'));
    const card = editor.view.container.querySelector<HTMLElement>('.docx-unified-sidebar .docx-comment-card[data-comment-id="1"]')!;
    expect(within(card).queryByTitle('More options')).toBeNull();
    if (moveBodyCaret) {
      await act(async () => {
        expect(editor.ref.current!.scrollToParaId(second!.paraId)).toBe(true);
      });
      expect(session.selection()?.head).toMatchObject({ story: 'body', paraId: second!.paraId, offset: 0 });
    } else {
      expect(session.selection()?.head).toMatchObject({ story: 'body', paraId: first!.paraId, offset: 0 });
    }
    await act(async () => editing.handleHeaderFooterDoubleClick('header', 1));
    await until(() => session.selection()?.head.story === 'hf:header');
    await act(async () => {
      paged.insertText('Edited ');
      await paged.flushPendingInput();
    });
    expect(session.paragraphs('hf:header')[0]!.text).toBe('Edited Header');
    await act(async () => editing.handleBodyClick());
    await until(() => session.selection()?.head.story === 'body');
    expect(session.selection()?.head).toMatchObject({ story: 'body', paraId: first!.paraId, offset: 0 });
    await editor.ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
    const page = paged.getLayout()!.pages[0]!;
    const fragment = page.fragments.find((entry) => entry.kind === 'paragraph')!;
    const canvas = editor.view.container.querySelector<HTMLCanvasElement>('canvas[data-page-index="0"]')!;
    canvas.getBoundingClientRect = () => new DOMRect(0, 0, page.size.w, page.size.h);
    const point = { clientX: fragment.x, clientY: fragment.y + fragment.height / 2, button: 0 };
    await until(() => paged.getPositionAtPoint(point.clientX, point.clientY)?.target.start.offset === 0);
    expect(paged.getPositionAtPoint(point.clientX, point.clientY)?.target).toMatchObject({
      story: 'body', start: { paraId: first!.paraId, offset: 0 }, end: { paraId: first!.paraId, offset: 0 },
    });
    await act(async () => {
      fireEvent.mouseDown(canvas, point);
      fireEvent.mouseUp(canvas, point);
      fireEvent.click(canvas, point);
      await paged.flushPendingInput();
    });
    expect(session.selection()?.head).toMatchObject({ story: 'body', paraId: first!.paraId, offset: 0 });
    expect(within(card).queryByTitle('More options')).not.toBeNull();
  } finally {
    captureEditing.mockRestore();
  }
}

test('clicking the first body caret after header editing expands its covering comment', async () => {
  await clickCommentAfterHeaderReturn(mount);
}, 30_000);

test('clicking the startup body caret after header editing expands its covering comment without prior navigation', async () => {
  await clickCommentAfterHeaderReturn(mount, false);
}, 30_000);

async function deleteTwoCommentsAcrossSaves() {
  const editor = await mount(twoCommentFixture());
  await until(() => editor.ref.current!.getComments().length === 2);
  if (!editor.ref.current!.commands.getState('commentsSidebar').active) {
    await act(async () => {
      expect((await editor.ref.current!.commands.execute('commentsSidebar', null)).ok).toBe(true);
    });
  }
  await until(() => editor.view.container.querySelectorAll('.docx-unified-sidebar .docx-comment-card').length === 2);
  expect(editor.view.container.querySelectorAll('.docx-unified-sidebar .docx-comment-card [title="More options"]')).toHaveLength(0);
  const deleteComment = async (id: number) => {
    await until(() => !!editor.view.container.querySelector(`.docx-unified-sidebar .docx-comment-card[data-comment-id="${id}"]`));
    const card = editor.view.container.querySelector<HTMLElement>(`.docx-unified-sidebar .docx-comment-card[data-comment-id="${id}"]`)!;
    await act(async () => fireEvent.click(card));
    await act(async () => fireEvent.click(within(card).getByTitle('More options')));
    await act(async () => fireEvent.click(within(card).getByRole('menuitem', { name: 'Delete', hidden: true })));
    await until(() => !editor.ref.current!.getComments().some((comment) => comment.id === id));
  };
  await deleteComment(1);
  const first = unzipContainer(new Uint8Array(await editor.save()));
  await deleteComment(2);
  return { first, saved: await editor.save() };
}

async function removeCommentThroughProp(): Promise<ArrayBuffer> {
  const editor = await mount(twoCommentFixture());
  await until(() => editor.ref.current!.getComments().length === 2);
  await act(async () => editor.setComments(editor.ref.current!.getComments().filter(({ id }) => id !== 1)));
  await until(() => editor.ref.current!.getComments().length === 1);
  return editor.save();
}

test('a comment the host removes through the comments prop leaves its range and definition out', async () => {
  const saved = await removeCommentThroughProp();
  const parts = unzipContainer(new Uint8Array(saved));
  const comments = new DOMParser().parseFromString(xmlPart(parts, 'word/comments.xml'), 'application/xml');
  expect(xmlElements(comments, W, 'comment').map((entry) => entry.getAttribute('w:id'))).toEqual(['2']);
  expect(markers(xmlPart(parts, 'word/document.xml'), 1).filter((marker) => marker !== 'Reference')).toEqual([]);
  expect((await reopened(saved)).package.document.comments?.map(({ id }) => id)).toEqual([2]);
});

test.todo('a removed comment leaves no commentReference in the body (deferred-after-0.4.1: deleted comment keeps its commentReference)', async () => {
  const saved = await removeCommentThroughProp();
  expect(markers(xmlPart(unzipContainer(new Uint8Array(saved)), 'word/document.xml'), 1)).toEqual([]);
});

test('deleting two comments across saves does not resurrect the first comment', async () => {
  const { first, saved } = await deleteTwoCommentsAcrossSaves();
  const firstComments = new DOMParser().parseFromString(xmlPart(first, 'word/comments.xml'), 'application/xml');
  expect(xmlElements(firstComments, W, 'comment').map((entry) => entry.getAttribute('w:id'))).toEqual(['2']);
  expect(markers(xmlPart(first, 'word/document.xml'), 1).filter((marker) => marker !== 'Reference')).toEqual([]);
  const last = unzipContainer(new Uint8Array(saved));
  if (last['word/comments.xml']) {
    const lastComments = new DOMParser().parseFromString(xmlPart(last, 'word/comments.xml'), 'application/xml');
    expect(xmlElements(lastComments, W, 'comment').some((entry) => entry.getAttribute('w:id') === '1')).toBe(false);
  }
  expect((await reopened(saved)).package.document.comments?.some((comment) => comment.id === 1) ?? false).toBe(false);
});

test.todo('deleting the last comment removes every deleted comment\'s body range markers (deferred-after-0.4.1: last-comment delete leaves markers)', async () => {
  const { first, saved } = await deleteTwoCommentsAcrossSaves();
  expect(markers(xmlPart(first, 'word/document.xml'), 1)).toEqual([]);
  const xml = xmlPart(unzipContainer(new Uint8Array(saved)), 'word/document.xml');
  expect(markers(xml, 1)).toEqual([]);
  expect(markers(xml, 2)).toEqual([]);
});

test('a comment resolved in React is saved in commentsExtended and reopened', async () => {
  const editor = await mount(commentedFixture().bytes);
  await act(async () => editor.ref.current!.resolveComment(1));
  const saved = await editor.save();
  const parts = unzipContainer(new Uint8Array(saved));
  expect(xmlPart(parts, 'word/comments.xml')).toContain('w:id="1"');
  const paraId = savedCommentParaId(parts, 1);
  const extended = new DOMParser().parseFromString(xmlPart(parts, 'word/commentsExtended.xml'), 'application/xml');
  const comment = xmlElements(extended, W15, 'commentEx').find((entry) => entry.getAttribute('w15:paraId') === paraId);
  expect(comment?.getAttribute('w15:done')).toBe('1');
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
  const replyParaId = savedCommentParaId(parts, replyId!);
  const parentParaId = savedCommentParaId(parts, 1);
  expect(reply?.paraId).toBe(replyParaId);
  expect(parent?.paraId).toBe(parentParaId);
  const extended = new DOMParser().parseFromString(xmlPart(parts, 'word/commentsExtended.xml'), 'application/xml');
  const replyEx = xmlElements(extended, W15, 'commentEx').find((entry) => entry.getAttribute('w15:paraId') === replyParaId);
  expect(replyEx?.getAttribute('w15:paraIdParent')).toBe(parentParaId);
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

test('two React saves retain a drawing inserted through the public image picker', async () => {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const [first] = session.paragraphs('body');
  await act(async () => {
    session.setSelection({ story: 'body', paraId: first!.paraId, offset: 0 });
    paged.syncYrsInputState(false);
    const outcome = await editor.ref.current!.commands.execute('insertImage', null);
    expect(outcome.ok).toBe(true);
    if (outcome.ok) expect(outcome.status).toBe('opened');
  });
  const originalImage = globalThis.Image;
  class LoadedImage {
    naturalWidth = 1;
    naturalHeight = 1;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  const imageInserted = () => session.storySegments('body').some(
    (segment) => segment.kind === 'embed' && segment.embedKind === 'image'
  );
  try {
    globalThis.Image = LoadedImage as never;
    const input = editor.view.container.querySelector<HTMLInputElement>('input[type="file"][accept="image/*"]');
    expect(input).not.toBeNull();
    await act(async () => {
      fireEvent.change(input!, { target: { files: [new File([PNG], 'Synthetic pixel.png', { type: 'image/png' })] } });
    });
    await until(imageInserted);
    await act(async () => editor.ref.current!.flushPendingInput());
  } finally {
    globalThis.Image = originalImage;
  }
  const firstSave = await editor.save();
  const secondSave = await editor.save();
  for (const buffer of [firstSave, secondSave]) {
    const parts = unzipContainer(new Uint8Array(buffer));
    const xml = new DOMParser().parseFromString(xmlPart(parts, 'word/document.xml'), 'application/xml');
    const embed = xmlElements(xml, A, 'blip')[0]?.getAttribute('r:embed');
    expect(embed).toBeTruthy();
    expect(xmlElements(xml, WP, 'docPr')[0]?.getAttribute('descr')).toBe('Synthetic pixel.png');
    const document = await reopened(buffer);
    const images = document.package.document.content.flatMap((block) => block.type === 'paragraph'
      ? block.content.flatMap((content) => content.type === 'run'
        ? content.content.flatMap((item) => item.type === 'drawing' ? [item.image] : [])
        : [])
      : []);
    expect(images).toHaveLength(1);
    expect(images[0]?.rId).toBe(embed!);
    expect(images[0]?.alt).toBe('Synthetic pixel.png');
    expect(images[0]?.size).toEqual({ width: 9525, height: 9525 });
  }
});

test('two React saves each register a data-URL image the session inserted', async () => {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const [first] = session.paragraphs('body');
  await act(async () => {
    session.setSelection({ story: 'body', paraId: first!.paraId, offset: 0 });
    paged.syncYrsInputState(false);
    expect(paged.applyYrsCommand({ type: 'insertImage', image: { src: `data:image/png;base64,${PNG_BASE64}`, width: 1, height: 1, alt: 'Synthetic pixel' } })).toBe(true);
  });
  for (const buffer of [await editor.save(), await editor.save()]) {
    const parts = unzipContainer(new Uint8Array(buffer));
    const xml = new DOMParser().parseFromString(xmlPart(parts, 'word/document.xml'), 'application/xml');
    const embed = xmlElements(xml, A, 'blip')[0]?.getAttribute('r:embed');
    expect(embed).toBeTruthy();
    const rels = new DOMParser().parseFromString(xmlPart(parts, 'word/_rels/document.xml.rels'), 'application/xml');
    const relationship = Array.from(rels.getElementsByTagNameNS(RELS, 'Relationship')).find((entry) => entry.getAttribute('Id') === embed);
    expect(relationship?.getAttribute('Type')).toBe(`${R}/image`);
    const target = relationship!.getAttribute('Target')!;
    const name = target.startsWith('/') ? target.slice(1) : `word/${target}`;
    expect(name.startsWith('word/media/')).toBe(true);
    expect(Array.from(parts[name] ?? [])).toEqual(Array.from(PNG));
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

test('React save after undo restores raw source document bytes with permission ranges', async () => {
  const source = fixture((p) =>
    p('<w:permStart w:id="3" w:edGrp="everyone"/>' + run('Editable') + '<w:permEnd w:id="3"/>') +
    p(run('Body text'))
  );
  const editor = await mount(source.bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  const second = session.paragraphs('body')[1]!;
  await act(async () => {
    session.insertText({ story: 'body', paraId: second.paraId, offset: 0 }, 'Undo this ');
    paged.syncYrsInputState(true, ['body']);
  });
  const edited = await editor.save();
  expect(xmlPart(unzipContainer(new Uint8Array(edited)), 'word/document.xml')).toContain('Undo this ');
  await act(async () => {
    expect(paged.undo()).toBe(true);
  });
  expect(session.paragraphs('body')[0]!.text).toBe('Editable');
  expect(session.paragraphs('body')[1]!.text).toBe('Body text');
  const undone = await editor.save();
  expectUnchanged(source, undone, ['word/document.xml']);
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
  expect((topMargin as HTMLInputElement).value).toBe('2');
  await act(async () => fireEvent.click(within(dialog).getByRole('button', { name: 'Apply' })));
  await until(() => !editor.view.queryByRole('dialog'));
  const saved = await editor.save();
  const xml = new DOMParser().parseFromString(xmlPart(unzipContainer(new Uint8Array(saved)), 'word/document.xml'), 'application/xml');
  expect(xmlElements(xml, W, 'pgMar')[0]?.getAttribute('w:top')).toBe('2880');
  expect((await reopened(saved)).package.document.finalSectionProperties?.marginTop).toBe(2880);
});

async function mountHeaderFooterEditing() {
  const editor = await mount(fixture((p) => p(run('Body text'))).bytes);
  const paged = editor.ref.current!.getEditorRef()!;
  const session = paged.getYrsSession()!;
  let host = paged.getDocument()!;
  const errors: Error[] = [];
  const pagedEditorRef = {
    current: {
      getYrsSession: () => session,
      getDocument: () => host,
      flushPendingInput: () => paged.flushPendingInput(),
    } as PagedEditorRef,
  };
  const hook = renderHook(() => {
    const [document, pushDocument] = useState<DocxDocument>(host);
    const [partEditTarget, setPartEditTarget] = useState<PartEditTarget | null>(null);
    host = document;
    const editing = useHeaderFooterEditing({ document, pushDocument, partEditTarget, setPartEditTarget });
    const io = useFileIO({
      pagedEditorRef,
      resolveImage: () => null,
      comments: [],
      documentName: undefined,
      onSave: undefined,
      downloadOnSave: false,
      onOpen: undefined,
      onError: (error) => errors.push(error),
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    });
    return { editing, save: io.handleSave };
  });
  const save = async (): Promise<ArrayBuffer> => {
    let saved: ArrayBuffer | null = null;
    await act(async () => {
      saved = await hook.result.current.save();
    });
    expect(errors.map(({ message }) => message)).toEqual([]);
    expect(saved).not.toBeNull();
    return saved!;
  };
  return { hook, save, document: () => host };
}

function expectPackageTargetsExist(saved: ArrayBuffer): void {
  const parts = unzipContainer(new Uint8Array(saved));
  const rels = new DOMParser().parseFromString(xmlPart(parts, 'word/_rels/document.xml.rels'), 'application/xml');
  for (const entry of xmlElements(rels, RELS, 'Relationship')) {
    if (entry.getAttribute('TargetMode') === 'External') continue;
    const target = entry.getAttribute('Target')!;
    const name = new URL(target, 'https://package.test/word/document.xml').pathname.slice(1);
    expect(parts[name]).toBeDefined();
  }
  const types = new DOMParser().parseFromString(xmlPart(parts, '[Content_Types].xml'), 'application/xml');
  for (const entry of xmlElements(types, 'http://schemas.openxmlformats.org/package/2006/content-types', 'Override')) {
    expect(parts[entry.getAttribute('PartName')!.slice(1)]).toBeDefined();
  }
}

test('adding and removing a header across saves keeps relationship and content-type targets', async () => {
  const { hook, save, document } = await mountHeaderFooterEditing();
  await act(async () => hook.result.current.editing.handleHeaderFooterDoubleClick('header', 1));
  expect(document().package.headers?.size).toBe(1);
  const firstSave = await save();
  expect(unzipContainer(new Uint8Array(firstSave))['word/header1.xml']).toBeDefined();
  await act(async () => hook.result.current.editing.handleRemoveHeaderFooter());
  expect(document().package.headers?.size).toBe(0);
  const lastSave = await save();
  for (const saved of [firstSave, lastSave]) expectPackageTargetsExist(saved);
  expect((await reopened(lastSave)).package.document.finalSectionProperties?.headerReferences ?? []).toEqual([]);
});

test.each(['header', 'footer'] as const)('adding and removing a %s before the first save keeps all package targets present', async (kind) => {
  const { hook, save, document } = await mountHeaderFooterEditing();
  const mapKey = kind === 'header' ? 'headers' : 'footers';
  const refKey = kind === 'header' ? 'headerReferences' : 'footerReferences';
  await act(async () => hook.result.current.editing.handleHeaderFooterDoubleClick(kind, 1));
  expect(document().package[mapKey]?.size).toBe(1);
  await act(async () => hook.result.current.editing.handleRemoveHeaderFooter());
  expect(document().package[mapKey]?.size).toBe(0);
  const saved = await save();
  expectPackageTargetsExist(saved);
  expect(unzipContainer(new Uint8Array(saved))[`word/${kind}1.xml`]).toBeUndefined();
  const opened = await reopened(saved);
  expect(opened.package.document.finalSectionProperties?.[refKey] ?? []).toEqual([]);
  expect(opened.package[mapKey]?.size ?? 0).toBe(0);
});

test('an edit in one paragraph keeps the rest of the body byte-for-byte', async () => {
  const source = fixture((p) =>
    p(run('Edited here')) +
    p('<w:hyperlink w:anchor="target"><w:fldSimple w:instr="DATE">' + run('October') + '</w:fldSimple></w:hyperlink>') +
    p(run('Kept'), '<w:pPr><w:pageBreakBefore w:val="0"/><w:spacing w:after="0"/></w:pPr>') +
    '<w:altChunk r:id="chunk"/>' +
    p('<w:permStart w:edGrp="everyone" w:id="9"/>' + run('Permission') + '<w:permEnd w:id="9"/>') +
    p('<w:bookmarkStart w:id="7" w:name="target"/>' + run('Target') + '<w:bookmarkEnd w:id="7"/>'),
    { chunk: true }
  );
  const editor = await mount(source.bytes);
  await typeBody(editor, 'Typed ');
  const saved = xmlPart(unzipContainer(new Uint8Array(await editor.save())), 'word/document.xml');
  const original = new TextDecoder().decode(source.parts.get('word/document.xml'));
  const start = original.indexOf('<w:p ');
  const end = original.indexOf('</w:p>', start) + '</w:p>'.length;
  expectSameXml(saved.slice(0, start), original.slice(0, start), 'before the edited paragraph');
  expectSameXml(saved.slice(saved.length - (original.length - end)), original.slice(end), 'after the edited paragraph');
  expect(saved.slice(start, saved.length - (original.length - end)).replace(/<[^>]+>/g, '')).toBe(
    'Typed Edited here'
  );
});

describe('DocxEditor saves (worker engine)', () => {
  const workers = setupWorkerEngine();

  async function mountWorker(bytes: ArrayBuffer) {
    const opens = workers.flatMap((worker) => worker.requests).filter((type) => type === 'open').length;
    const ref = createRef<DocxEditorRef>();
    const errors: Error[] = [];
    const onError = (error: Error) => errors.push(error);
    const view = render(
      <DocxEditor ref={ref} experimentalWorkerOpen documentBuffer={bytes} onError={onError} downloadOnSave={false} />
    );
    await until(() => ref.current?.commands.getState('save').enabled === true);
    if (unzipContainer(new Uint8Array(bytes))['word/comments.xml']) {
      await until(() => (ref.current?.getComments().length ?? 0) > 0);
    }
    expect(errors.map(({ message }) => message)).toEqual([]);
    await act(async () => { await ref.current!.flushPendingInput(); });
    expect(workers.length).toBeGreaterThan(0);
    expect(workers.flatMap((worker) => worker.requests).filter((type) => type === 'open').length).toBeGreaterThan(opens);
    expect(workers.some((worker) => worker.sessions.length > 0)).toBe(true);
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
    };
  }

  test('clicking the first body caret after header editing expands its covering comment on the worker engine', async () => {
    await clickCommentAfterHeaderReturn(mountWorker);
  }, 30_000);

  test('clicking the startup body caret after header editing expands its covering comment without prior navigation on the worker engine', async () => {
    await clickCommentAfterHeaderReturn(mountWorker, false);
  }, 30_000);

  async function deleteTwoCommentsAcrossSaves() {
    const editor = await mountWorker(twoCommentFixture());
    await until(() => editor.ref.current!.getComments().length === 2);
    if (!editor.ref.current!.commands.getState('commentsSidebar').active) {
      await act(async () => {
        expect((await editor.ref.current!.commands.execute('commentsSidebar', null)).ok).toBe(true);
      });
    }
    await until(() => editor.view.container.querySelectorAll('.docx-unified-sidebar .docx-comment-card').length === 2);
    expect(editor.view.container.querySelectorAll('.docx-unified-sidebar .docx-comment-card [title="More options"]')).toHaveLength(0);
    const deleteComment = async (id: number) => {
      await until(() => !!editor.view.container.querySelector(`.docx-unified-sidebar .docx-comment-card[data-comment-id="${id}"]`));
      const card = editor.view.container.querySelector<HTMLElement>(`.docx-unified-sidebar .docx-comment-card[data-comment-id="${id}"]`)!;
      await act(async () => fireEvent.click(card));
      await act(async () => fireEvent.click(within(card).getByTitle('More options')));
      await act(async () => fireEvent.click(within(card).getByRole('menuitem', { name: 'Delete', hidden: true })));
      await until(() => !editor.ref.current!.getComments().some((comment) => comment.id === id));
    };
    await deleteComment(1);
    const first = unzipContainer(new Uint8Array(await editor.save()));
    await deleteComment(2);
    return { first, saved: await editor.save() };
  }

  test('deleting two comments across saves does not resurrect the first comment on the worker engine', async () => {
    const { first, saved } = await deleteTwoCommentsAcrossSaves();
    const firstComments = new DOMParser().parseFromString(xmlPart(first, 'word/comments.xml'), 'application/xml');
    expect(xmlElements(firstComments, W, 'comment').map((entry) => entry.getAttribute('w:id'))).toEqual(['2']);
    expect(markers(xmlPart(first, 'word/document.xml'), 1).filter((marker) => marker !== 'Reference')).toEqual([]);
    const last = unzipContainer(new Uint8Array(saved));
    if (last['word/comments.xml']) {
      const lastComments = new DOMParser().parseFromString(xmlPart(last, 'word/comments.xml'), 'application/xml');
      expect(xmlElements(lastComments, W, 'comment').some((entry) => entry.getAttribute('w:id') === '1')).toBe(false);
    }
    expect((await reopened(saved)).package.document.comments?.some((comment) => comment.id === 1) ?? false).toBe(false);
  });

  test('adding and removing a header across saves keeps relationship and content-type targets on the worker engine', async () => {
    const useEditing = useHeaderFooterEditing;
    let mountedHost!: Parameters<typeof useHeaderFooterEditing>[0];
    const captureHost = spyOn(headerFooterEditing, 'useHeaderFooterEditing').mockImplementation((options) => {
      mountedHost = options;
      return useEditing(options);
    });
    try {
      const editor = await mountWorker(fixture((p) => p(run('Body text'))).bytes);
      const paged = editor.ref.current!.getEditorRef()!;
      const session = paged.getYrsSession()!;
      expect(session).not.toBeNull();
      expect(mountedHost).toBeDefined();
      let host = paged.getDocument()!;
      const errors: Error[] = [];
      const pagedEditorRef = {
        current: {
          getYrsSession: () => session,
          getDocument: () => host,
          flushPendingInput: () => paged.flushPendingInput(),
        } as PagedEditorRef,
      };
      const hook = renderHook(() => {
        const [document, setDocument] = useState<DocxDocument>(host);
        const [partEditTarget, setPartEditTarget] = useState<PartEditTarget | null>(null);
        host = document;
        const editing = useEditing({
          document,
          pushDocument: (next) => {
            mountedHost.pushDocument(next);
            setDocument(next);
          },
          partEditTarget,
          setPartEditTarget,
        });
        const io = useFileIO({
          pagedEditorRef,
          resolveImage: () => null,
          comments: [],
          documentName: undefined,
          onSave: undefined,
          downloadOnSave: false,
          onOpen: undefined,
          onError: (error) => errors.push(error),
          onPrint: undefined,
          onDocumentNameChange: undefined,
          loadBuffer: async () => {},
          focusActiveEditor: () => {},
        });
        return { editing, save: io.handleSave };
      });
      const workerSaves = () => workers.flatMap((worker) => worker.requests).filter((type) => type === 'save').length;
      const save = async (): Promise<ArrayBuffer> => {
        expect(editor.ref.current!.getEditorRef()!.getYrsSession()).toBe(session);
        expect(mountedHost.document).toBe(host);
        const saves = workerSaves();
        let saved: ArrayBuffer | null = null;
        await act(async () => {
          saved = await hook.result.current.save();
        });
        expect(errors.map(({ message }) => message)).toEqual([]);
        expect(saved).not.toBeNull();
        expect(workerSaves()).toBe(saves + 1);
        return saved!;
      };
      await act(async () => hook.result.current.editing.handleHeaderFooterDoubleClick('header', 1));
      expect(host.package.headers?.size).toBe(1);
      const firstSave = await save();
      expect(unzipContainer(new Uint8Array(firstSave))['word/header1.xml']).toBeDefined();
      await act(async () => hook.result.current.editing.handleRemoveHeaderFooter());
      expect(host.package.headers?.size).toBe(0);
      const lastSave = await save();
      for (const saved of [firstSave, lastSave]) {
        const parts = unzipContainer(new Uint8Array(saved));
        const rels = new DOMParser().parseFromString(xmlPart(parts, 'word/_rels/document.xml.rels'), 'application/xml');
        for (const entry of xmlElements(rels, RELS, 'Relationship')) {
          if (entry.getAttribute('TargetMode') === 'External') continue;
          const target = entry.getAttribute('Target')!;
          const name = new URL(target, 'https://package.test/word/document.xml').pathname.slice(1);
          expect(parts[name]).toBeDefined();
        }
        const types = new DOMParser().parseFromString(xmlPart(parts, '[Content_Types].xml'), 'application/xml');
        for (const entry of xmlElements(types, 'http://schemas.openxmlformats.org/package/2006/content-types', 'Override')) {
          expect(parts[entry.getAttribute('PartName')!.slice(1)]).toBeDefined();
        }
      }
      expect((await reopened(lastSave)).package.document.finalSectionProperties?.headerReferences ?? []).toEqual([]);
    } finally {
      captureHost.mockRestore();
    }
  });
});
