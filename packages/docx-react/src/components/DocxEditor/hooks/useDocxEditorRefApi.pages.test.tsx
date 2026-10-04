import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef } from 'react';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, decodeDocxHostJson, type ResidentDocumentRead, type ResidentEngineWorkerClient, type YrsSession } from '@betteroffice/docx/yrs';
import { createResidentEngineSession } from '@betteroffice/docx/yrs/residentEngineSession';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import type { DocxHostSearch } from './useHostSearch';
import type { Comment } from '@betteroffice/docx/types/content';
import { exportWorkerOpenPages, registerWorkerOpenExport, type WorkerOpenExport } from '../internals/workerOpenExport';
import { workerExportVersions } from '../internals/workerExportVersions';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const workers: Awaited<ReturnType<typeof createResidentEngineSession>>[] = [];

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
  for (const worker of workers.splice(0)) worker.destroy();
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

async function openWorkerSession(afterRead?: (value: string) => Promise<void>, bytes = PAGES) {
  const worker = await createResidentEngineSession();
  workers.push(worker);
  const { document } = decodeDocxHostJson(worker.openDocx(bytes), bytes);
  const font = worker.registerFont(FONT);
  const inputs = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(worker.layoutFontRequirementsJson(JSON.stringify(inputs))) as Array<{ key: string }>;
  inputs.measurement = {
    fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  const request = JSON.stringify(inputs);
  const peer = await createYrsSession();
  sessions.push(peer);
  peer.openDocx(bytes, false);
  peer.loadState(worker.encodeState());
  peer.registerFont(FONT);
  const owner = {};
  const reads: Array<{ options: Parameters<WorkerOpenExport['export']>[0]; version: string }> = [];
  const operation: WorkerOpenExport = {
    export: (options, context) => exportWorkerOpenPages(peer, options, context, {
      assertCurrent: () => {},
      catchUp: async () => {
        const vector = peer.encodeStateVector();
        worker.applyUpdate(peer.encodeStateAsUpdate(worker.encodeStateVector()));
        peer.applyLocalUpdate(worker.encodeStateAsUpdate(vector));
        return { P: peer.version(), W: worker.proposalEngine.version(), changed: false };
      },
      read: (async (read: ResidentDocumentRead, version: string) => {
        if (worker.proposalEngine.version() !== version) return { status: 'superseded' };
        if (read.kind !== 'exportStructuredWithPages') throw new Error('Expected paged export');
        reads.push({ options: read.options, version });
        const value = worker.exportStructuredWithPagesJson(read.options, read.currentRequest);
        await afterRead?.(value);
        return { status: 'ok', version, value };
      }) as ResidentEngineWorkerClient['documentReadAt'],
      versions: workerExportVersions(peer, owner, 1),
      serialize: (run) => run(),
    }),
  };
  return { session: peer, worker, operation, reads, request, layout: (current = request) => { worker.layoutDocumentWithRegionsRetainedJson(current); } };
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
  worker?: WorkerOpenExport;
  settle?: () => Promise<void>;
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
      events.push(relayoutOptions?.onHost ? 'relayout' : 'relayout in the worker');
      options.relayout?.();
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  if (options.worker) registerWorkerOpenExport(options.session, options.worker);
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      hostSearch: {} as DocxHostSearch,
      ref,
      document: null,
      documentFromYrs: () => null,
      historyStateRef: { current: null },
      pagedEditorRef,
      experimentalWorkerOpen: options.worker !== undefined,
      settledDisplayList: options.worker ? async () => {
        await options.settle?.();
        return {} as never;
      } : undefined,
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

test('worker editor export lays out flushed input and returns a peer edit version', async () => {
  const opened = await openWorkerSession();
  opened.layout();
  const peerExport = spyOn(opened.session, 'exportStructuredWithPagesFor');
  const { events, api } = await setup({
    session: opened.session, worker: opened.operation, request: () => opened.request,
    flush: () => edit(opened.session, 'Typed before the worker export'), relayout: opened.layout,
  });
  const result = await api().exportStructuredWithPages(MARKUP);
  if (!result.ok) throw new Error(result.failure.message);
  expect(events).toEqual(['flush', 'relayout in the worker']);
  expect(result.version).toBe(opened.session.version());
  expect(result.content.layout.documentVersion).toBe(result.version);
  const title = result.content.structured.stories[0]!.blocks[0]!;
  expect(title.kind === 'heading' && title.paragraph.inlines[0]).toMatchObject({ kind: 'text', text: 'Typed before the worker export' });
  const target = { kind: 'paragraph', story: 'body', paraId: '00000001' } as const;
  expect(opened.session.validateEdits({ expectVersion: result.version, steps: [{ op: 'replaceText', target, text: 'Next edit' }] }).ok).toBe(true);
  const applied = opened.session.applyEdits({ expectVersion: result.version, steps: [{ op: 'replaceText', target, text: 'Next edit' }] });
  expect(applied.ok).toBe(true);
  expect(opened.session.validateEdits({ expectVersion: result.version, steps: [{ op: 'replaceText', target, text: 'Stale edit' }] })).toMatchObject({ ok: false });
  expect(peerExport).not.toHaveBeenCalled();
  peerExport.mockRestore();
});

test('worker editor export reuses its retained layout and pinned tokens round-trip', async () => {
  const opened = await openWorkerSession();
  opened.layout();
  const { events, api } = await setup({ session: opened.session, worker: opened.operation, request: () => opened.request, relayout: opened.layout });
  const first = await api().exportStructuredWithPages(MARKUP);
  if (!first.ok) throw new Error(first.failure.message);
  expect(first.content.layout.layoutVersion.startsWith(`${opened.session.version()}:`)).toBe(true);
  const second = await api().exportStructuredWithPages({ ...MARKUP, expectLayoutVersion: first.content.layout.layoutVersion });
  expect(second).toEqual(first);
  expect(opened.reads[1]!.options.expectLayoutVersion).toBe(`${opened.reads[0]!.version}:${first.content.layout.layoutVersion.split(':').at(-1)}`);
  expect(events).toEqual(['flush', 'flush']);
});

test('worker editor export waits for deferred layout completion', async () => {
  const paragraphs = Array.from({ length: 64 }, (_, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Page ${index + 1}</w:t></w:r></w:p>`
  ).join('');
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraphs}</w:body></w:document>`)],
  ])));
  const opened = await openWorkerSession(undefined, bytes);
  const prefix = JSON.parse(opened.worker.layoutDocumentWithRegionsPrefixRetainedJson(opened.request, 1));
  expect(prefix).toMatchObject({ provisional: true, layout: { partial: true } });
  expect(JSON.parse(opened.worker.exportStructuredWithPagesJson(MARKUP, opened.request)))
    .toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
  let completed = false;
  const { events, api } = await setup({
    session: opened.session, worker: opened.operation, request: () => opened.request,
    settle: () => new Promise((resolve) => setTimeout(() => { opened.layout(); completed = true; resolve(); }, 20)),
  });
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  expect(completed).toBe(true);
  expect(events).toEqual(['flush', 'relayout in the worker']);
  expect(opened.reads).toHaveLength(2);
});

test('worker editor export lays out changed inputs and waits for fonts', async () => {
  const opened = await openWorkerSession();
  opened.layout();
  const inputs = JSON.parse(opened.request) as { renderEnv: Record<string, unknown> };
  inputs.renderEnv.showHiddenText = true;
  const changed = JSON.stringify(inputs);
  let fontsReady = false;
  const { events, api } = await setup({
    session: opened.session, worker: opened.operation, request: () => fontsReady ? changed : null,
    settle: async () => { fontsReady = true; opened.layout(changed); },
  });
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  expect(events).toEqual(['flush', 'relayout in the worker']);
});

test('worker editor export refreshes a layout made with other inputs', async () => {
  const opened = await openWorkerSession();
  opened.layout();
  const inputs = JSON.parse(opened.request) as { renderEnv: Record<string, unknown> };
  inputs.renderEnv.showHiddenText = true;
  const changed = JSON.stringify(inputs);
  const { events, api } = await setup({
    session: opened.session, worker: opened.operation, request: () => changed,
    relayout: () => opened.layout(changed),
  });
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  expect(events).toEqual(['flush', 'relayout in the worker']);
  expect(opened.reads).toHaveLength(2);
});

test('worker editor export refreshes a retained revision preview', async () => {
  const opened = await openWorkerSession();
  const inputs = JSON.parse(opened.request) as { renderEnv: Record<string, unknown> };
  inputs.renderEnv.revisionPreview = { '1': 'accepted' };
  opened.layout(JSON.stringify(inputs));
  const { events, api } = await setup({
    session: opened.session, worker: opened.operation, request: () => opened.request, relayout: opened.layout,
  });
  expect((await api().exportStructuredWithPages(MARKUP)).ok).toBe(true);
  expect(events).toEqual(['flush', 'relayout in the worker']);
});

test('a pinned worker editor export refuses changed input without relayout', async () => {
  const opened = await openWorkerSession();
  opened.layout();
  const { events, api } = await setup({ session: opened.session, worker: opened.operation, request: () => opened.request, relayout: opened.layout });
  const first = await api().exportStructuredWithPages(MARKUP);
  if (!first.ok) throw new Error(first.failure.message);
  edit(opened.session, 'Changed after capture');
  const refusal = await api().exportStructuredWithPages({ ...MARKUP, expectLayoutVersion: first.content.layout.layoutVersion });
  expect(refusal).toMatchObject({ ok: false, version: opened.session.version(), failure: { code: 'stale-document' } });
  expect(events).toEqual(['flush', 'flush']);
  expect(opened.reads).toHaveLength(2);
});

test('later typing does not appear in the captured worker export', async () => {
  let release!: () => void;
  let posted!: () => void;
  const pendingRead = new Promise<void>((resolve) => { release = resolve; });
  const reading = new Promise<void>((resolve) => { posted = resolve; });
  const opened = await openWorkerSession((value) => {
    if (!(JSON.parse(value) as { ok: boolean }).ok) return Promise.resolve();
    posted();
    return pendingRead;
  });
  const { api } = await setup({
    session: opened.session, worker: opened.operation, request: () => opened.request,
    flush: () => { edit(opened.session, 'Flushed text'); },
    relayout: opened.layout,
  });
  const pending = api().exportStructuredWithPages(MARKUP);
  await reading;
  const capturedVersion = opened.session.version();
  edit(opened.session, 'Later typing');
  release();
  const result = await pending;
  if (!result.ok) throw new Error(result.failure.message);
  expect(result.version).toBe(capturedVersion);
  expect(JSON.stringify(result.content.structured)).toContain('Flushed text');
  expect(JSON.stringify(result.content.structured)).not.toContain('Later typing');
  expect(opened.reads).toHaveLength(2);
});
