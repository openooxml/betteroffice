import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, useRef, useState, type RefObject } from 'react';
import { repackDocx } from '@betteroffice/docx/docx';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import { getLayoutKernelInputs } from '@betteroffice/docx/editor';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type DocxEditStep, type YrsSession } from '@betteroffice/docx/yrs';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import type { DocxEditorRef } from '../../DocxEditor';
import { PagedEditor, type PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import { useYrsCoreSession, type YrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const DOCX = resolve(import.meta.dir, '../../../../../docx/src');
const FONT = resolve(import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
let fontBytes: ArrayBuffer;

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(DOCX, 'wasm/generated/edit/docx_edit_bg.wasm'))));
  fontBytes = readFileSync(FONT).buffer as ArrayBuffer;
});
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function setup(options: { flush?: () => void; mode?: EditorMode } = {}) {
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(
    new Uint8Array(readFileSync(resolve(DOCX, 'yrs/__fixtures__/content-controls/template.docx'))),
    true
  );
  const events: string[] = [];
  const comments: Comment[] = [];
  const current = { session };
  const editor = {
    getYrsSession: () => current.session,
    flushPendingInput: async () => {
      events.push('flush');
      options.flush?.();
    },
    syncYrsInputState: (docChanged: boolean, stories?: readonly string[]) => {
      events.push(`sync:${docChanged}:${stories?.join(',') ?? '*'}`);
      return true;
    },
  } as unknown as PagedEditorRef;
  const modeRef = { current: options.mode ?? ('editing' as EditorMode) };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      ref,
      document: null,
      documentFromYrs: () => null,
      historyStateRef: { current: null },
      pagedEditorRef: { current: editor },
      handleSave: async () => null,
      handleDirectPrint: () => {},
      zoom: 1,
      setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {},
      loadBuffer: async () => {},
      comments,
      setComments: () => {},
      setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() },
      selectionChangeSubscribersRef: { current: new Set() },
      getCachedStyleResolver: (() => {
        throw new Error('unused');
      }) as never,
      commentIdAllocator: createCommentIdAllocator(),
      modeRef,
    });
    return ref;
  });
  const api = () => {
    const current = hook.result.current.current;
    if (!current) throw new Error('the ref API is not mounted');
    return current;
  };
  return { session, current, events, modeRef, api };
}

test('lists after flushing and fills through one published batch', async () => {
  const { events, api } = await setup();
  const listed = await api().listContentControls();
  if (!listed.ok) throw new Error(listed.failure.message);
  expect(events).toEqual(['flush']);
  const found = await api().findContentControls({ kind: 'tag', tag: 'customer.address' });
  if (!found.ok) throw new Error(found.failure.message);
  const address = found.content.controls[0]!;
  const name = listed.content.controls.find((control) => control.tag === 'customer.name')!;
  const result = await api().applyEdits({
    expectVersion: listed.version,
    steps: [
      { op: 'setContentControlText', target: { kind: 'id', controlId: name.controlId }, text: 'Ada' },
      { op: 'setContentControlText', target: { kind: 'id', controlId: address.controlId }, text: 'A\nB' },
    ],
  });
  expect(result).toMatchObject({ ok: true, applied: true, changedStories: ['body', 'body:sdt0'] });
  expect(events.filter((event) => event.startsWith('sync'))).toEqual(['sync:true:body,body:sdt0']);
  const again = await api().listContentControls({ stories: ['body'] });
  if (!again.ok) throw new Error(again.failure.message);
  expect(again.content.controls.find((control) => control.tag === 'customer.address')?.value).toEqual({
    kind: 'text',
    text: 'A\nB',
  });
});

test('typing flushed after the read makes the fill stale', async () => {
  let typed = false;
  const context: { session?: YrsSession } = {};
  const { session, api } = await setup({
    flush: () => {
      if (!typed || !context.session) return;
      const paragraph = context.session.paragraphs('body')[0]!;
      context.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'x');
    },
  });
  context.session = session;
  const listed = await api().listContentControls();
  if (!listed.ok) throw new Error(listed.failure.message);
  typed = true;
  const result = await api().applyEdits({
    expectVersion: listed.version,
    steps: [{ op: 'setContentControlText', target: { kind: 'tag', tag: 'document.title' }, text: 'x' }],
  });
  expect(result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
});

test('the editor mode governs fills', async () => {
  const { session, modeRef, api } = await setup({ mode: 'viewing' });
  const step = {
    op: 'setContentControlText',
    target: { kind: 'tag', tag: 'document.title' },
    text: 'x',
  } as const;
  expect(await api().applyEdits({ expectVersion: session.version(), steps: [step] })).toMatchObject({
    ok: false,
    failure: { code: 'read-only' },
  });
  modeRef.current = 'suggesting';
  expect(await api().applyEdits({ expectVersion: session.version(), steps: [step] })).toMatchObject({
    ok: false,
    failure: { code: 'invalid-step' },
  });
  expect(
    await api().applyEdits({
      expectVersion: session.version(),
      steps: [{ ...step, suggest: { author: 'Reviewer', date: '2026-09-25T00:00:00Z' } } as DocxEditStep],
    })
  ).toMatchObject({ ok: false, failure: { code: 'unsupported', reason: 'unsupported-suggestion' } });
});

test('discovery refuses when flushing replaces the document', async () => {
  const replacement = await createYrsSession();
  sessions.push(replacement);
  const context: { replace?: () => void } = {};
  const { session, current, api } = await setup({ flush: () => context.replace?.() });
  context.replace = () => {
    current.session = replacement;
  };
  await expect(api().listContentControls()).rejects.toThrow('The document changed while flushing input');
  current.session = session;
  await expect(api().findContentControls({ kind: 'tag', tag: 'document.title' })).rejects.toThrow(
    'The document changed while flushing input'
  );
});

test('a refusal after the flush publishes nothing and records no history', async () => {
  const context: { flip?: () => void } = {};
  const { session, events, modeRef, api } = await setup({ flush: () => context.flip?.() });
  const version = session.version();
  context.flip = () => {
    modeRef.current = 'viewing';
  };
  expect(
    await api().applyEdits({
      expectVersion: version,
      steps: [{ op: 'setContentControlText', target: { kind: 'tag', tag: 'document.title' }, text: 'x' }],
    })
  ).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  context.flip = undefined;
  modeRef.current = 'editing';
  expect(
    await api().applyEdits({
      expectVersion: version,
      steps: [{ op: 'setContentControlText', target: { kind: 'tag', tag: 'terms.standard' }, text: 'x' }],
    })
  ).toMatchObject({ ok: false, failure: { code: 'locked-target', reason: 'content-locked' } });
  expect(events).toEqual(['flush', 'flush']);
  expect(session.version()).toBe(version);
  expect(session.canUndo()).toBe(false);
});

function Mounted({
  apiRef,
  editorRef,
  coreRef,
  layouts,
}: {
  apiRef: RefObject<DocxEditorRef | null>;
  editorRef: RefObject<PagedEditorRef | null>;
  coreRef: { current: YrsCoreSession | null };
  layouts: Layout[];
}) {
  const [bytes] = useState(() =>
    new Uint8Array(readFileSync(resolve(DOCX, 'yrs/__fixtures__/content-controls/template.docx')))
  );
  const [host, setHost] = useState<Document | null>(null);
  const core = useYrsCoreSession(true, host, null, bytes, 0, undefined, {
    onHostDocument: (next) => setHost(next.document),
  });
  coreRef.current = host ? core : null;
  const modeRef = useRef<EditorMode>('editing');
  useDocxEditorRefApi({
    ref: apiRef,
    document: host,
    documentFromYrs: () => editorRef.current?.getDocument() ?? null,
    historyStateRef: { current: null },
    pagedEditorRef: editorRef,
    handleSave: async () => null,
    handleDirectPrint: () => {},
    zoom: 1,
    setZoom: () => {},
    scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
    loadParsedDocument: () => {},
    loadBuffer: async () => {},
    comments: [],
    setComments: () => {},
    setShowCommentsSidebar: () => {},
    contentChangeSubscribersRef: { current: new Set() },
    selectionChangeSubscribersRef: { current: new Set() },
    getCachedStyleResolver: (() => {
      throw new Error('unused');
    }) as never,
    commentIdAllocator: createCommentIdAllocator(),
    modeRef,
  });
  return (
    <PagedEditor
      ref={editorRef}
      document={host}
      yrsCore={core}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      onLayoutComputed={(layout) => {
        if (layout) layouts.push(layout);
      }}
    />
  );
}

async function until(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 5000;
  while (!done() && Date.now() < deadline) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

test('a fill through the mounted editor renders and saves the same text', async () => {
  const apiRef = createRef<DocxEditorRef>();
  const editorRef = createRef<PagedEditorRef>();
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  const layouts: Layout[] = [];
  const rendered = () => JSON.stringify(getLayoutKernelInputs(layouts.at(-1)!)?.measured ?? null);
  render(<Mounted apiRef={apiRef} editorRef={editorRef} coreRef={coreRef} layouts={layouts} />);
  await until(() => !!coreRef.current?.session && !!apiRef.current && layouts.length > 0);
  expect(rendered()).toContain('Click to enter a name.');
  const result = await act(() =>
    apiRef.current!.applyEdits({
      expectVersion: coreRef.current!.session!.version(),
      steps: [{ op: 'setContentControlText', target: { kind: 'tag', tag: 'customer.name' }, text: 'Ada Lovelace' }],
    })
  );
  expect(result).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
  await until(() => rendered().includes('Ada Lovelace'));
  expect(rendered()).not.toContain('Click to enter a name.');
  const document = editorRef.current!.getDocument();
  if (!document) throw new Error('the editor has no document to save');
  const parts = unzipContainer(new Uint8Array(await act(() => repackDocx(document))));
  const xml = new TextDecoder().decode(parts['word/document.xml'] as Uint8Array);
  expect(xml).toContain('Ada Lovelace');
  expect(xml).not.toContain('Click to enter a name.');
});
