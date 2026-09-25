import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef } from 'react';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';

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
}) {
  const events: string[] = [];
  const editor = {
    getYrsSession: () => options.session,
    getLayoutRequest: options.request,
    flushPendingInput: async () => {
      events.push('flush');
      options.flush?.();
    },
    relayout: () => {
      events.push('relayout');
      options.relayout?.();
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      ref,
      document: null,
      documentFromYrs: () => null,
      historyStateRef: { current: null },
      pagedEditorRef,
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
      modeRef: { current: 'editing' as EditorMode },
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
