import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef } from 'react';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import type { DocxHostSearch } from './useHostSearch';
import type { Comment } from '@betteroffice/docx/types/content';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

const ROOT = resolve(import.meta.dir, '../../../../../..');
const PAGES = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx'))
);
const FONT = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))
);
const MARKUP = { revisionView: 'markup' } as const;

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm')))
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/**
 * A session holding the fixture, the request its editor lays it out with, and a layout with a
 * request.
 */
async function openSession(): Promise<{
  session: YrsSession;
  request: string;
  layout: (request?: string) => void;
}> {
  const session = await createYrsSession();
  sessions.push(session);
  const { document } = session.openDocx(PAGES, true);
  const font = session.registerFont(FONT);
  const request = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(
    session.layoutFontRequirementsJson(JSON.stringify(request))
  ) as Array<{ key: string }>;
  request.measurement = {
    fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  const json = JSON.stringify(request);
  return {
    session,
    request: json,
    layout: (current = json) => void session.layoutDocumentWithRegionsRetainedJson(current),
  };
}

function edit(session: YrsSession, text: string) {
  const applied = session.applyEdits({
    expectVersion: session.version(),
    steps: [
      { op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId: '00000001' }, text },
    ],
  });
  if (!applied.ok) throw new Error(applied.failure.message);
}

async function setup(options: {
  session: YrsSession;
  request: () => string | null;
  flush?: () => void;
  relayout?: () => void;
  opening?: { current: boolean };
  comments?: Comment[];
  setComments?: () => void;
  save?: () => Promise<ArrayBuffer | null>;
  awaitingDocument?: () => boolean;
}) {
  const events: string[] = [];
  const editor = {
    getYrsSession: () => options.session,
    getLayoutRequest: options.request,
    readLayoutRequest: async () => options.request(),
    flushPendingInput: async () => {
      events.push('flush');
      options.flush?.();
    },
    relayout: (relayoutOptions?: { onHost?: boolean }) => {
      // An export lays out on this thread: a worker pass would leave the session's layout as it is.
      events.push(relayoutOptions?.onHost ? 'relayout' : 'relayout in the worker');
      options.relayout?.();
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      hostSearch: {} as DocxHostSearch,
      ref,
      document: null,
      documentFromYrs: () => null,
      historyStateRef: { current: null },
      pagedEditorRef,
      handleSave: options.save ?? (async () => null),
      zoom: 1,
      setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {},
      loadBuffer: async () => {},
      comments: options.comments ?? [],
      setComments: options.setComments ?? (() => {}),
      setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() },
      selectionChangeSubscribersRef: { current: new Set() },
      getCachedStyleResolver: (() => {
        throw new Error('unused');
      }) as never,
      commentIdAllocator: createCommentIdAllocator(),
      commands: UNAVAILABLE_DOCX_COMMANDS,
      modeRef: { current: 'editing' as EditorMode },
      openingRef: options.opening,
      allowHostProposalsRef: { current: false },
      awaitingDocument: options.awaitingDocument,
    });
    return ref;
  });
  const api = () => {
    const current = hook.result.current.current;
    if (!current) throw new Error('the ref API is not mounted');
    return current;
  };
  return { events, pagedEditorRef, api };
}

test('a partial layout has no page contents yet', async () => {
  const { session, request } = await openSession();
  const { layout } = JSON.parse(session.layoutDocumentWithRegionsRetainedJson(request)) as {
    layout: Layout;
  };
  const { pagedEditorRef, api } = await setup({ session, request: () => request });
  Object.assign(pagedEditorRef.current!, { getLayout: () => ({ ...layout, partial: true }) });
  expect(api().getPageContent(1)).toBeNull();
});

test('the page count is 0 until the loaded document is laid out in full', async () => {
  const { session, request } = await openSession();
  const { layout } = JSON.parse(session.layoutDocumentWithRegionsRetainedJson(request)) as {
    layout: Layout;
  };
  let awaiting = false;
  let current: Layout = { ...layout, partial: true };
  const { pagedEditorRef, api } = await setup({
    session,
    request: () => request,
    awaitingDocument: () => awaiting,
  });
  Object.assign(pagedEditorRef.current!, { getLayout: () => current });
  expect(api().getTotalPages()).toBe(0);
  current = layout;
  expect(api().getTotalPages()).toBe(layout.pages.length);
  awaiting = true;
  expect(api().getTotalPages()).toBe(0);
});

test('flushed input is laid out before its pages are exported', async () => {
  const { session, request, layout } = await openSession();
  layout();
  const { events, api } = await setup({
    session,
    request: () => request,
    flush: () => edit(session, 'Typed before the export'),
    relayout: layout,
  });
  const result = await api().exportStructuredWithPages(MARKUP);
  if (!result.ok) throw new Error(result.failure.message);
  expect(events).toEqual(['flush', 'relayout']);
  expect(result.version).toBe(session.version());
  expect(result.content.layout.documentVersion).toBe(session.version());
  const title = result.content.structured.stories[0]!.blocks[0]!;
  expect(title.kind === 'heading' && title.paragraph.inlines[0]).toMatchObject({
    kind: 'text',
    text: 'Typed before the export',
  });
});

test('while the document opens, the API reads, exports and changes nothing', async () => {
  const { session, request, layout } = await openSession();
  layout();
  const opening = { current: true };
  const commentChanges: string[] = [];
  const { events, api } = await setup({
    session,
    request: () => request,
    opening,
    comments: [{ id: 1, content: [], author: 'A' } as unknown as Comment],
    setComments: () => commentChanges.push('changed'),
    save: async () => {
      commentChanges.push('saved');
      return new ArrayBuffer(0);
    },
  });
  expect(api().getDocument()).toBeNull();
  expect(api().getComments()).toEqual([]);
  await expect(api().exportStructuredWithPages(MARKUP)).rejects.toThrow();
  await expect(api().readParagraphs({ story: 'body' } as never)).rejects.toThrow();
  expect(api().findInDocument('a')).toEqual([]);
  expect(api().replyToComment(1, 'reply', 'B')).toBeNull();
  api().resolveComment(1);
  expect(await api().save()).toBeNull();
  expect(commentChanges).toEqual([]);
  expect(events).toEqual([]);

  opening.current = false;
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  api().resolveComment(1);
  expect(commentChanges).toEqual(['changed']);
});

test('a current layout is exported without laying out again', async () => {
  const { session, request, layout } = await openSession();
  layout();
  const { events, api } = await setup({ session, request: () => request, relayout: layout });
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  expect(events).toEqual(['flush']);
});

test('a deferred layout is awaited', async () => {
  const { session, request, layout } = await openSession();
  const { events, api } = await setup({ session, request: () => request });
  setTimeout(layout, 50);
  const result = await api().exportStructuredWithPages(MARKUP);
  expect(result.ok).toBe(true);
  expect(events).toEqual(['flush', 'relayout']);
});

test('a layout from other inputs is laid out again with the current ones', async () => {
  const { session, request, layout } = await openSession();
  layout();
  const current = JSON.parse(request) as { renderEnv: Record<string, unknown> };
  current.renderEnv = { ...current.renderEnv, showHiddenText: true };
  const hidden = JSON.stringify(current);
  expect(session.exportStructuredWithPagesFor(MARKUP, hidden)).toMatchObject({
    ok: false,
    failure: { code: 'stale-layout' },
  });
  const { events, api } = await setup({
    session,
    request: () => hidden,
    relayout: () => layout(hidden),
  });
  const result = await api().exportStructuredWithPages(MARKUP);
  if (!result.ok) throw new Error(result.failure.message);
  expect(events).toEqual(['flush', 'relayout']);
});

test('a layout that previews decisions the editor no longer shows is laid out again', async () => {
  const { session, request, layout } = await openSession();
  const previewing = JSON.parse(request) as { renderEnv: Record<string, unknown> };
  previewing.renderEnv = { ...previewing.renderEnv, revisionPreview: { '1': 'accepted' } };
  layout(JSON.stringify(previewing));
  expect(session.exportStructuredWithPagesFor(MARKUP, request)).toMatchObject({
    ok: false,
    failure: { code: 'unsupported-revision-layout' },
  });
  const { events, api } = await setup({ session, request: () => request, relayout: layout });
  const result = await api().exportStructuredWithPages(MARKUP);
  if (!result.ok) throw new Error(result.failure.message);
  expect(events).toEqual(['flush', 'relayout']);
});

test('an export waits for the fonts the document needs', async () => {
  const { session, request, layout } = await openSession();
  layout();
  let ready = false;
  const { api } = await setup({ session, request: () => (ready ? request : null) });
  setTimeout(() => {
    ready = true;
  }, 50);
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
});

test('replacing the document while waiting for its layout throws', async () => {
  const { session, request } = await openSession();
  const replacement = await openSession();
  const { pagedEditorRef, api } = await setup({ session, request: () => request });
  setTimeout(() => {
    pagedEditorRef.current = {
      getYrsSession: () => replacement.session,
    } as unknown as PagedEditorRef;
  }, 50);
  await expect(api().exportStructuredWithPages(MARKUP)).rejects.toThrow(
    'The document changed while it was being laid out'
  );
});

test('an expected layout version is never laid out again', async () => {
  const { session, request, layout } = await openSession();
  layout();
  const first = await (
    await setup({ session, request: () => request })
  ).api().exportStructuredWithPages(MARKUP);
  if (!first.ok) throw new Error(first.failure.message);
  const { events, api } = await setup({
    session,
    request: () => request,
    flush: () => edit(session, 'Changed'),
    relayout: layout,
  });
  const result = await api().exportStructuredWithPages({
    ...MARKUP,
    expectLayoutVersion: first.content.layout.layoutVersion,
  });
  expect(result).toMatchObject({ ok: false, failure: { code: 'stale-document' } });
  expect(events).toEqual(['flush']);
});
