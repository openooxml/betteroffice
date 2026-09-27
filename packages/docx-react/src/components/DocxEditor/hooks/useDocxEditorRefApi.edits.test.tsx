import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef } from 'react';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type DocxEditRequest, type YrsSession } from '@betteroffice/docx/yrs';
import type { Comment } from '@betteroffice/docx/types/content';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Line one</w:t></w:r><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">target words and more</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function setup(options: { flush?: () => Promise<void> | void; mode?: EditorMode } = {}) {
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(fixture(), true);
  const events: string[] = [];
  const comments: Comment[] = [];
  const editor = {
    getYrsSession: () => session,
    flushPendingInput: async () => {
      events.push('flush');
      await options.flush?.();
    },
    syncYrsInputState: (docChanged: boolean, stories?: readonly string[]) => {
      events.push(`sync:${docChanged}:${stories?.join(',') ?? '*'}`);
      return true;
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  const modeRef = { current: options.mode ?? ('editing' as EditorMode) };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      ref,
      document: null,
      documentFromYrs: () => null,
      historyStateRef: { current: null },
      pagedEditorRef,
      handleSave: async () => null,
      zoom: 1,
      setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {},
      loadBuffer: async () => {},
      comments,
      setComments: (update) => {
        const next = typeof update === 'function' ? update(comments) : update;
        comments.splice(0, comments.length, ...next);
      },
      setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() },
      selectionChangeSubscribersRef: { current: new Set() },
      getCachedStyleResolver: (() => {
        throw new Error('unused');
      }) as never,
      commentIdAllocator: createCommentIdAllocator(),
      commands: UNAVAILABLE_DOCX_COMMANDS,
      modeRef,
    });
    return ref;
  });
  const api = () => {
    const current = hook.result.current.current;
    if (!current) throw new Error('the ref API is not mounted');
    return current;
  };
  return { session, events, editor, pagedEditorRef, modeRef, api, comments };
}

const target = { kind: 'search', text: 'Tail', within: { kind: 'paragraph', story: 'body', paraId: '00000002' }, view: 'accepted' } as const;

function request(session: YrsSession, text = 'Head'): DocxEditRequest {
  return {
    expectVersion: session.version(),
    steps: [{ op: 'replaceText', target, text }],
  };
}

function bodyTexts(session: YrsSession, view: 'accepted' | 'original' = 'accepted'): string[] {
  const read = session.readParagraphs({ view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((paragraph) => paragraph.text);
}

test('applyEdits flushes input first and publishes one applied batch', async () => {
  const { session, events, api } = await setup();
  const read = await api().readParagraphs({ view: 'accepted' });
  expect(read.ok).toBe(true);
  if (!read.ok) return;
  const result = await api().applyEdits({ expectVersion: read.version, steps: request(session).steps });
  expect(result).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
  expect(events).toEqual(['flush', 'flush', 'sync:true:body']);
  expect(bodyTexts(session)[1]).toBe('Head');
  const noOp = await api().applyEdits({
    expectVersion: session.version(),
    steps: [{ op: 'replaceText', target: { ...target, text: 'Head' }, text: 'Head' }],
  });
  expect(noOp).toMatchObject({ ok: true, applied: false });
  expect(events.filter((event) => event.startsWith('sync'))).toHaveLength(1);
});

test('a batch made stale by flushed typing refuses without undoing the typing', async () => {
  let session!: YrsSession;
  const state = await setup({
    flush: () => {
      session.insertText({ story: 'body', paraId: '00000002', offset: 4 }, ' typed');
    },
  });
  session = state.session;
  const stale = request(session);
  const result = await state.api().applyEdits(stale);
  expect(result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  expect(bodyTexts(session)[1]).toBe('Tail typed');
  expect(state.events).toEqual(['flush']);
  const validated = await state.api().validateEdits(stale);
  expect(validated).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
});

test('a handle rebuilt while input flushes keeps the batch; a replaced session aborts it', async () => {
  let rebuild = () => {};
  const layout = await setup({ flush: () => rebuild() });
  rebuild = () => {
    layout.pagedEditorRef.current = { ...layout.editor } as PagedEditorRef;
  };
  expect(await layout.api().applyEdits(request(layout.session))).toMatchObject({ ok: true, applied: true });
  expect(await layout.api().readParagraphs({ view: 'accepted' })).toMatchObject({ ok: true });
  await layout.api().flushPendingInput();
  expect(bodyTexts(layout.session)[1]).toBe('Head');
  expect(layout.events).toEqual(['flush', 'sync:true:body', 'flush', 'flush']);

  const other = await createYrsSession();
  sessions.push(other);
  let replace = () => {};
  const replaced = await setup({ flush: () => replace() });
  replace = () => {
    replaced.pagedEditorRef.current = { ...replaced.editor, getYrsSession: () => other } as PagedEditorRef;
  };
  const before = replaced.session.version();
  await expect(replaced.api().applyEdits(request(replaced.session))).rejects.toThrow('document changed');
  expect(replaced.session.version()).toBe(before);
  replaced.pagedEditorRef.current = replaced.editor;
  await expect(replaced.api().flushPendingInput()).rejects.toThrow('document changed');
});

test('read-only and suggesting modes gate batches before and after the flush', async () => {
  const viewing = await setup({ mode: 'viewing' });
  expect(await viewing.api().applyEdits(request(viewing.session))).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });
  expect(viewing.events).toEqual([]);

  const suggesting = await setup({ mode: 'suggesting' });
  expect(await suggesting.api().applyEdits(request(suggesting.session))).toMatchObject({
    ok: false,
    failure: { code: 'invalid-step', stepIndex: 0 },
  });
  const suggested = await suggesting.api().applyEdits({
    expectVersion: suggesting.session.version(),
    steps: [{ op: 'replaceText', target, text: 'Head', suggest: { author: 'Ann', date: '2026-09-24T00:00:00Z' } }],
  });
  expect(suggested).toMatchObject({ ok: true, applied: true });
  if (suggested.ok) expect(suggested.receipts[0]!.revisionIds).toHaveLength(1);
  expect(bodyTexts(suggesting.session, 'original')[1]).toBe('Tail');

  let lock = () => {};
  const switching = await setup({ flush: () => lock() });
  lock = () => {
    switching.modeRef.current = 'viewing';
  };
  expect(await switching.api().applyEdits(request(switching.session))).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });
  expect(bodyTexts(switching.session)[1]).toBe('Tail');
});

test('legacy helpers target the text after a hard break', async () => {
  const { session, api, comments, events } = await setup();
  expect(api().applyFormatting({ paraId: '00000001', search: 'words', marks: { bold: true } })).toBe(true);
  const bold = session
    .storySegments('body')
    .filter((segment) => segment.kind === 'text' && segment.attributes.bold === true)
    .map((segment) => (segment.kind === 'text' ? segment.text : ''));
  expect(bold).toEqual(['words']);

  const commentId = api().addComment({ paraId: '00000001', text: 'Check', author: 'Ann', search: 'and' });
  expect(commentId).not.toBeNull();
  expect(comments.map((comment) => comment.id)).toEqual([commentId!]);
  const [anchor] = session.resolveComment(String(commentId));
  expect([anchor!.start, anchor!.end]).toEqual([22, 25]);

  expect(
    api().proposeChange({ paraId: '00000001', search: 'target', replaceWith: 'chosen', author: 'Agent' })
  ).toBe(true);
  expect(bodyTexts(session)[0]).toBe('Line one\uFFFCchosen words and more');
  expect(bodyTexts(session, 'original')[0]).toBe('Line one\uFFFCtarget words and more');
  expect(
    api().proposeChange({ paraId: '00000001', search: 'chosen', replaceWith: 'again', author: 'Agent' })
  ).toBe(false);
  expect(api().proposeChange({ paraId: '00000001', search: 'absent', replaceWith: 'x', author: 'Agent' })).toBe(
    false
  );
  expect(events.filter((event) => event.startsWith('sync'))).toEqual([
    'sync:true:body',
    'sync:true:body',
    'sync:true:body',
  ]);

  session.setSelection(
    { story: 'body', paraId: '00000001', offset: 5 },
    { story: 'body', paraId: '00000001', offset: 28 }
  );
  expect(api().getSelectionInfo()).toMatchObject({
    paraId: '00000001',
    selectedText: 'one\uFFFCchosen words ',
    before: 'Line ',
    after: 'and more',
  });
});
