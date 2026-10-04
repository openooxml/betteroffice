import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import { createStyleResolver } from '@betteroffice/docx/styles';
import type { Document } from '@betteroffice/docx/types/document';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { layoutMetaSummary } from '@betteroffice/docx/yrs';
import type {
  DocxContentControlsResult,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxParagraphIdentitySnapshot,
  DocxProposalRequest,
  ResidentDocumentRead,
  ResidentEngineWorkerClient,
  ResidentProposalReply,
  YrsSession,
} from '@betteroffice/docx/yrs';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxDocumentChange, DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import type { EditorMode } from '../internals/editing-modes';
import { resetDeprecatedViewerMembersForTests } from '../internals/deprecatedViewerMembers';
import { markPresented, stampWorkerFrameVersion } from '../internals/layoutProvenance';
import { navigateViewer, readViewerSelectionInfo, type ViewerNavigationTarget, type ViewerRefReadAccess } from '../internals/viewerRefReads';
import { deferWorkerOpenReplica, requestWorkerOpenReplica } from '../internals/workerOpenReplica';
import { beginWorkerProposalHandover, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import { usePagedEditorRefApi } from './usePagedEditorRefApi';
import { useDocxCommandBinding, type DocxCommandInputs } from './useDocxCommands';
import { DocxAsyncOnlyError, DocxReplicaNotReadyError, routeViewerRefAccess, useDocxEditorRefApi } from './useDocxEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
beforeEach(resetDeprecatedViewerMembersForTests);
afterEach(() => { cleanup(); mock.restore(); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const INFO = { paraId: 'p', selectedText: 'hello', paragraphText: 'hello', before: '', after: '' };
const PAGE_OPTIONS = { revisionView: 'markup' } as const;
const PAGE_REQUEST = JSON.stringify({ renderEnv: {} });
const PAGE_EXPORT = {
  ok: true, version: 'worker-v',
  content: { structured: {}, layout: { documentVersion: 'worker-v', layoutVersion: 'layout-v', pages: [] } },
} as unknown as Awaited<ReturnType<DocxEditorRef['exportStructuredWithPages']>>;
const PAGE_REFUSAL = {
  ok: false, version: 'worker-v',
  failure: { code: 'stale-layout', target: null, message: 'The layout is stale.' },
} as const;
const MATCHES = [{ paraId: 'p', match: 'hello', before: '', after: '' }];
const CONTROLS: DocxContentControlsResult = {
  ok: true, version: 'worker-v',
  content: { schemaVersion: 1, anchorScope: 'session', includedStories: ['body'], controls: [], complete: true, diagnostics: [] },
};

function apiFor(viewer = false, pendingReplica = false, settledDisplayList?: Parameters<typeof useDocxEditorRefApi>[0]['settledDisplayList'], bindCommands = false, viewerSession = false) {
  const events: string[] = [];
  const document = { package: {} } as Document;
  const state = { viewer, version: 'v' };
  const modeRef = { current: viewer || viewerSession ? 'viewing' : 'editing' } as { current: EditorMode };
  const allowHostProposalsRef = { current: true };
  const sidebar = mock(() => {});
  const setComments = mock((..._args: Parameters<Parameters<typeof useDocxEditorRefApi>[0]['setComments']>) => {});
  const session = {
    version: () => state.version,
    storyIds: () => ['body'],
    paragraphs: () => state.viewer ? [] : [{ paraId: 'p', text: 'hello', properties: {} }],
    paragraphIdentities: mock(() => ({ sessionId: 'shell', packageSha256: null, paragraphs: [] })),
    locateParagraph: () => ({ start: 0, end: 5 }),
    selection: () => ({ anchor: { story: 'body', paraId: 'p', offset: 0 }, head: { story: 'body', paraId: 'p', offset: 5 } }),
    selectionText: () => { events.push('selection'); return INFO; },
    commentTextTarget: mock(() => ({ ok: true })),
    formatTextTarget: mock(() => ({ ok: true })),
    applyParagraphStyle: mock(() => {}),
    insertPageBreak: mock(() => {}),
    mirrorWorkerDocument: (mirror: { version: string } | null) => { if (mirror) state.version = mirror.version; },
    getProposals: () => ({ version: 'v', previewVersion: 0, proposals: [] }),
    applyEdits: mock(() => ({ ok: true, applied: true, changedStories: ['body'] })),
    validateEdits: mock(() => ({ ok: true, version: state.version })),
    findText: mock(() => ({ ok: true, version: state.version, matches: [], truncated: false })),
    listContentControls: mock(() => CONTROLS),
    findContentControls: mock(() => CONTROLS),
    canUndo: () => false,
    canRedo: () => false,
    onUpdate: () => () => {},
    exportStructuredWithPagesFor: mock((): Awaited<ReturnType<DocxEditorRef['exportStructuredWithPages']>> => PAGE_EXPORT),
    layoutFontRequirementsJson: mock(() => '[]'),
  } as unknown as YrsSession;
  const hydrate = mock(async () => () => {});
  const fallback = mock(() => {});
  const request = mock(() => {});
  const replica = pendingReplica ? deferWorkerOpenReplica(session, hydrate, fallback, () => {}, { active: () => true, request }) : null;
  const editor = {
    isWorkerViewer: () => state.viewer,
    getYrsSession: () => session,
    getDocument: () => document,
    flushPendingInput: mock(async () => { events.push('flush'); }),
    getPositionAtPoint: mock(() => { events.push('point'); return null; }),
    readPositionAtPoint: mock(async () => null),
    readViewerSelectionInfo: mock(async () => INFO),
    navigateViewer: mock(async (..._args: Parameters<PagedEditorRef['navigateViewer']>) => true),
    scrollToParaId: mock(() => { events.push('paragraph'); return true; }),
    scrollToCommentId: mock(() => { events.push('comment'); return false; }),
    scrollToChangeId: mock(() => { events.push('change'); return true; }),
    syncYrsInputState: () => { events.push('sync'); return true; },
    getLayout: (): Layout | null => null,
    getLayoutRequest: mock((): string | null => PAGE_REQUEST),
    readLayoutRequest: mock(async (): Promise<string | null> => PAGE_REQUEST),
    relayout: mock(() => {}),
  };
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const bridge = {
    session: () => session, rootStory: () => 'body', hasPendingInput: () => pendingReplica,
    subscribe: () => () => {}, toolbarSelection: () => null, hasSelection: () => false,
    runAfterPendingInput: mock(async () => { throw new Error('unexpected input admission'); }),
  };
  const subscribers = new Set<(change: DocxDocumentChange) => void>();
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    const commands = useDocxCommandBinding({
      session: bindCommands ? session : null, pagedEditorRef, bridgeRef: { current: bridge },
      document, readOnly: true, mode: 'viewing', experimentalWorkerOpen: pendingReplica, viewerSession,
    } as unknown as DocxCommandInputs);
    useDocxEditorRefApi({
      ref, document, documentFromYrs: () => document, historyStateRef: { current: document }, pagedEditorRef,
      experimentalWorkerOpen: pendingReplica, settledDisplayList, viewerSession,
      handleSave: async () => null, zoom: 1, setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {}, loadBuffer: async () => {},
      comments: [{ id: 1 } as never], setComments, setShowCommentsSidebar: sidebar,
      contentChangeSubscribersRef: { current: new Set() }, documentChangeSubscribersRef: { current: subscribers },
      selectionChangeSubscribersRef: { current: new Set() }, getCachedStyleResolver: createStyleResolver,
      commentIdAllocator: createCommentIdAllocator(), commands: bindCommands ? commands.controller.store : UNAVAILABLE_DOCX_COMMANDS,
      modeRef, allowHostProposalsRef,
      hostSearch: {
        search: async () => ({ query: '', options: { caseSensitive: false }, total: 0, current: -1 }),
        searchNext: () => null, searchPrevious: () => null, searchGoTo: () => null, clearSearch: () => {},
        getSearchState: () => null, onSearchChange: () => () => {},
      },
    });
    return ref;
  });
  return { api: hook.result.current.current!, editor, session, state, replica, hydrate, fallback, request, events, pagedEditorRef, subscribers, modeRef, allowHostProposalsRef, sidebar, setComments, bridge };
}

const WORKER_IDENTITIES: DocxParagraphIdentitySnapshot = {
  sessionId: 'worker-session', packageSha256: null,
  paragraphs: [
    { session: { kind: 'session', sessionId: 'worker-session', story: 'header', paraId: 'p' }, origin: 'authored', ooxmlParaId: null, idOrigin: null, persisted: null, source: null },
    { session: { kind: 'session', sessionId: 'worker-session', story: 'body:cell', paraId: 'p' }, origin: 'authored', ooxmlParaId: null, idOrigin: null, persisted: null, source: null },
  ],
};

function workerFor(
  host: ReturnType<typeof apiFor>,
  identities = WORKER_IDENTITIES,
  laidOut: () => Promise<void> = async () => {}
) {
  const snapshot: ResidentProposalReply = {
    mirror: { version: 'worker-v', proposals: { previewVersion: 0, entries: [] } },
    result: { ok: true, snapshot: { version: 'worker-v', previewVersion: 0, proposals: [] } },
    changedStories: [], geometry: { version: 'worker-v', previewVersion: 0, proposals: '', targets: {}, hidden: [] },
    updates: [], stateVector: new Uint8Array(),
  };
  const proposal = mock(async (_operation: Parameters<ResidentEngineWorkerClient['proposal']>[0]) => snapshot);
  const documentRead = mock(async (read: ResidentDocumentRead): Promise<{ version: string; value: unknown }> => ({
    version: 'worker-v', value: read.kind === 'paragraphIdentities' ? identities : MATCHES,
  }));
  const authority = registerWorkerProposalAuthority(host.session, {
    proposal,
    documentRead: documentRead as ResidentEngineWorkerClient['documentRead'],
    handOver: async () => ({ state: new Uint8Array(), version: 'worker-v', proposals: snapshot.mirror.proposals }),
  }, { relayout: () => {}, current: () => true, laidOut, adopted: () => {}, handedOver: () => {}, contentChanged: () => {} });
  return { authority, proposal, documentRead, snapshot };
}

function expectNoReplica(host: ReturnType<typeof apiFor>) {
  expect(host.replica!.started).toBe(false);
  expect(host.request).not.toHaveBeenCalled();
  expect(host.hydrate).not.toHaveBeenCalled();
  expect(host.fallback).not.toHaveBeenCalled();
  expect(host.editor.flushPendingInput).not.toHaveBeenCalled();
}

for (const search of ['hello', '']) {
  test(`viewer proposeChange queues ${search ? 'replacement' : 'append'} using worker identities and version`, async () => {
    spyOn(console, 'warn').mockImplementation(() => {});
    const host = apiFor(true, true);
    const worker = workerFor(host);
    const proposed = deferred<DocxProposalRequest>();
    const propose = worker.authority.propose;
    const call = spyOn(worker.authority, 'propose').mockImplementation((request, main) => {
      proposed.resolve(request);
      return propose(request, main);
    });
    expect(host.api.proposeChange({ paraId: 'p', search, replaceWith: 'world', author: 'Host' })).toBe(true);
    expectNoReplica(host);
    const request = await proposed.promise;
    await worker.authority.getProposals(async () => host.session.getProposals());
    expect(call).toHaveBeenCalledTimes(1);
    expect(request).toEqual({
      expectVersion: 'worker-v',
      proposals: [{
        id: expect.any(String),
        paragraph: { kind: 'session', sessionId: 'worker-session', story: 'body:cell', paraId: 'p' },
        suggest: { author: 'Host', date: expect.any(String) },
        ...(search ? { op: 'replaceText', search, replaceWith: 'world' } : { op: 'insertText', at: 'end', text: 'world' }),
      }],
    });
    expect(request.proposals[0]!.id.length).toBeGreaterThan(0);
    expect(new Date(request.proposals[0]!.suggest.date).toISOString()).toBe(request.proposals[0]!.suggest.date);
    expect(host.session.paragraphIdentities).not.toHaveBeenCalled();
    expect(host.session.applyEdits).not.toHaveBeenCalled();
    expect(host.sidebar).not.toHaveBeenCalled();
    expectNoReplica(host);
  });
}

test('viewer proposeChange returns false for empty input without starting a worker call', () => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(true, true);
  const worker = workerFor(host);
  expect(host.api.proposeChange({ paraId: 'p', search: '', replaceWith: '', author: 'Host' })).toBe(false);
  expect(worker.proposal).not.toHaveBeenCalled();
  expect(worker.documentRead).not.toHaveBeenCalled();
  expectNoReplica(host);
});

test('viewer proposeChange builds each queued request after the previous proposal settles', async () => {
  const host = apiFor(true, true);
  const identities = {
    ...WORKER_IDENTITIES,
    paragraphs: [...WORKER_IDENTITIES.paragraphs, {
      ...WORKER_IDENTITIES.paragraphs[1]!,
      session: { ...WORKER_IDENTITIES.paragraphs[1]!.session!, paraId: 'q' },
    }],
  };
  const worker = workerFor(host, identities);
  await worker.authority.initialize();
  const firstStarted = deferred<void>();
  const firstSettled = deferred<void>();
  const secondStarted = deferred<DocxProposalRequest>();
  let version = 'worker-v';
  spyOn(worker.authority, 'geometry').mockImplementation(() => ({ ...worker.snapshot.geometry, version }));
  const reads = spyOn(worker.authority, 'paragraphIdentities');
  const propose = spyOn(worker.authority, 'propose').mockImplementation(async (request) => {
    if (request.proposals[0]!.paragraph.kind !== 'session') throw new Error('Expected session paragraph');
    if (request.proposals[0]!.paragraph.paraId === 'p') {
      firstStarted.resolve();
      await firstSettled.promise;
      version = 'worker-v2';
    } else {
      secondStarted.resolve(request);
    }
    return { ok: true, snapshot: { version, previewVersion: 0, proposals: [] } };
  });
  for (const paraId of ['p', 'q']) {
    expect(host.api.proposeChange({ paraId, search: 'hello', replaceWith: 'world', author: 'Host' })).toBe(true);
  }
  await firstStarted.promise;
  await worker.authority.getProposals(async () => host.session.getProposals());
  expect(reads).toHaveBeenCalledTimes(1);
  expect(propose).toHaveBeenCalledTimes(1);
  firstSettled.resolve();
  const second = await secondStarted.promise;
  expect(reads).toHaveBeenCalledTimes(2);
  expect(propose).toHaveBeenCalledTimes(2);
  expect(second.expectVersion).toBe('worker-v2');
  expect(second.proposals[0]!.paragraph).toMatchObject({ paraId: 'q' });
  expectNoReplica(host);
});

for (const failure of ['missing paragraph', 'refused proposal', 'rejected proposal']) {
  test(`viewer proposeChange warns once and swallows a ${failure}`, async () => {
    const warned = deferred<void>();
    const warning = spyOn(console, 'warn').mockImplementation((message) => {
      if (message === '[DocxEditor] proposeChange:') warned.resolve();
    });
    const host = apiFor(true, true);
    const worker = workerFor(host, failure === 'missing paragraph' ? { ...WORKER_IDENTITIES, paragraphs: [] } : WORKER_IDENTITIES);
    if (failure === 'refused proposal') {
      worker.proposal.mockImplementation(async () => ({
        ...worker.snapshot, result: { ok: false, version: 'worker-v', failure: { code: 'read-only', message: 'Refused' } },
      }));
    } else if (failure === 'rejected proposal') {
      spyOn(worker.authority, 'propose').mockImplementation(async () => { throw new Error('Failed'); });
    }
    const options = { paraId: 'p', search: 'hello', replaceWith: 'world', author: 'Host' };
    expect(host.api.proposeChange(options)).toBe(true);
    expect(host.api.proposeChange(options)).toBe(true);
    await warned.promise;
    await worker.authority.getProposals(async () => host.session.getProposals());
    expect(warning.mock.calls.filter(([message]) => message === '[DocxEditor] proposeChange:')).toHaveLength(1);
    expectNoReplica(host);
  });
}

test('viewer proposeChange without host admission retains the replica gate', () => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(true, true);
  host.allowHostProposalsRef.current = false;
  expect(host.api.proposeChange({ paraId: 'p', search: 'hello', replaceWith: 'world', author: 'Host' })).toBe(false);
  expect(host.fallback).toHaveBeenCalledTimes(1);
});

test('editor proposeChange preserves synchronous edits, refresh and sidebar behavior', () => {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor();
  expect(host.api.proposeChange({ paraId: 'p', search: 'hello', replaceWith: 'world', author: 'Host' })).toBe(true);
  expect(host.session.applyEdits).toHaveBeenCalledWith({
    expectVersion: 'v', source: 'agent',
    steps: [{ op: 'replaceText', target: { kind: 'search', text: 'hello', within: { kind: 'paragraph', story: 'body', paraId: 'p' }, view: 'accepted' }, text: 'world', suggest: { author: 'Host', date: expect.any(String) } }],
  });
  expect(host.events).toEqual(['sync']);
  expect(host.sidebar).toHaveBeenCalledWith(true);
  expect(host.api.proposeChange({ paraId: 'p', search: '', replaceWith: '', author: 'Host' })).toBe(false);
  expect(warning).not.toHaveBeenCalled();
});

for (const member of ['applyEdits', 'validateEdits'] as const) {
  test(`viewer ${member} refuses before replica waiting with the mirrored worker version`, async () => {
    const host = apiFor(true, true);
    await workerFor(host).authority.initialize();
    host.state.version = 'worker-v~';
    expect(await host.api[member]({ expectVersion: 'worker-v', source: 'agent', steps: [] })).toEqual({
      ok: false, version: 'worker-v', failure: { code: 'read-only', message: 'The editor is read-only' },
    });
    expect(host.session[member]).not.toHaveBeenCalled();
    expectNoReplica(host);
  });
}

test('viewer findText reads through the worker authority without a replica', async () => {
  const host = apiFor(true, true);
  const worker = workerFor(host);
  const request: DocxFindTextRequest = { text: 'hello', within: { kind: 'story', story: 'body' }, view: 'accepted', limit: 2 };
  const result: DocxFindTextResult = { ok: true, version: 'worker-v', matches: [], truncated: false };
  worker.documentRead.mockImplementation(async () => ({ version: 'worker-v', value: result }));
  const call = spyOn(worker.authority, 'findText');
  expect(await host.api.findText(request)).toEqual(result);
  expect(call).toHaveBeenCalledWith(request, expect.any(Function));
  expect(worker.documentRead).toHaveBeenCalledWith({ kind: 'findText', request });
  expect(host.session.findText).not.toHaveBeenCalled();
  expectNoReplica(host);
});

function expectWorkerPageExport(host: ReturnType<typeof apiFor>) {
  expectNoReplica(host);
  expect(host.editor.getLayoutRequest).not.toHaveBeenCalled();
  expect(host.editor.relayout).not.toHaveBeenCalled();
  expect(host.session.exportStructuredWithPagesFor).not.toHaveBeenCalled();
  expect(host.session.layoutFontRequirementsJson).not.toHaveBeenCalled();
}

test('viewer content-control reads answer from the worker without loading the replica', async () => {
  const host = apiFor(true, true);
  const worker = workerFor(host);
  worker.documentRead.mockResolvedValue({ version: 'worker-v', value: CONTROLS });
  const options = { stories: ['body'], maxControls: 1 } as const;
  const query = { kind: 'ooxmlId', ooxmlId: '1' } as const;
  expect(await host.api.listContentControls(options)).toEqual(CONTROLS);
  expect(await host.api.findContentControls(query, options)).toEqual(CONTROLS);
  expect(await host.api.listContentControls()).toEqual(CONTROLS);
  expect(await host.api.findContentControls(query)).toEqual(CONTROLS);
  expect(worker.documentRead.mock.calls.map(([read]) => read)).toEqual([
    { kind: 'listContentControls', options },
    { kind: 'findContentControls', query, options },
    { kind: 'listContentControls', options: {} },
    { kind: 'findContentControls', query, options: {} },
  ]);
  expect(host.session.listContentControls).not.toHaveBeenCalled();
  expect(host.session.findContentControls).not.toHaveBeenCalled();
  expectNoReplica(host);
});

test('editor content-control reads still flush and read the main session', async () => {
  const host = apiFor(false, true);
  const worker = workerFor(host);
  const options = { stories: ['body'] } as const;
  const query = { kind: 'tag', tag: 'field' } as const;
  const listed = host.api.listContentControls(options);
  host.replica!.start();
  expect(await listed).toEqual(CONTROLS);
  expect(await host.api.findContentControls(query, options)).toEqual(CONTROLS);
  expect(host.session.listContentControls).toHaveBeenCalledWith(options);
  expect(host.session.findContentControls).toHaveBeenCalledWith(query, options);
  expect(host.editor.flushPendingInput).toHaveBeenCalledTimes(2);
  expect(worker.documentRead).not.toHaveBeenCalled();
});

test('viewer proposal decisions still use the worker after a refused revision command', async () => {
  const host = apiFor(true, true, undefined, true);
  const worker = workerFor(host);
  const refusal = await host.api.commands.execute('reviewAccept', { revisionId: 'revision' });
  expect(refusal).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  expect(host.bridge.runAfterPendingInput).not.toHaveBeenCalled();
  expectNoReplica(host);
  await worker.authority.initialize();
  const request = { expectVersion: host.session.version(), expectPreviewVersion: 0, changes: [] };
  expect(await host.api.setProposalStates(request)).toEqual(worker.snapshot.result!);
  expect(worker.proposal).toHaveBeenLastCalledWith({ kind: 'setStates', request });
  expectNoReplica(host);
});

test.each([false, true])('viewer paged export and synchronous refusal ignore the published layout (summary=%s)', async (summary) => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(true, true);
  const worker = workerFor(host);
  const layout: Layout = {
    pageSize: { w: 816, h: 1056 },
    pages: [{
      number: 1, size: { w: 816, h: 1056 }, margins: { top: 72, right: 72, bottom: 72, left: 72 },
      fragments: [{ kind: 'shape', blockId: 'shape', x: 72, y: 72, width: 40, height: 40 }],
    }],
  };
  const published = summary ? layoutMetaSummary({
    v: 1, layoutRevision: 1, pageCount: 1, partial: false, provisional: false,
    notesConverged: true, pageSizes: Float64Array.of(816, 1056), headersFootersEpoch: 1,
    layoutShell: JSON.stringify({ ...layout, pages: layout.pages.map((page) => ({ ...page, fragments: [] })) }),
  }) : layout;
  const getLayout = spyOn(host.editor, 'getLayout').mockReturnValue(published);
  let caught: unknown;
  try { host.api.getPageContent(1); } catch (error) { caught = error; }
  const expected = new DocxAsyncOnlyError('getPageContent', 'exportStructuredWithPages');
  expect(caught).toBeInstanceOf(DocxAsyncOnlyError);
  expect(caught).toMatchObject({ member: expected.member, use: expected.use, message: expected.message });
  worker.documentRead.mockResolvedValue({ version: 'worker-v', value: JSON.stringify(PAGE_EXPORT) });
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual(PAGE_EXPORT);
  expect(getLayout).not.toHaveBeenCalled();
  expect(worker.documentRead).toHaveBeenCalledWith({ kind: 'exportStructuredWithPages', options: PAGE_OPTIONS, currentRequest: PAGE_REQUEST });
  expect(host.editor.readLayoutRequest).toHaveBeenCalledTimes(1);
  expectWorkerPageExport(host);
});

test('viewer paged export refuses unavailable fonts after the poll budget', async () => {
  const host = apiFor(true, true);
  const worker = workerFor(host);
  await worker.authority.initialize();
  host.editor.readLayoutRequest.mockResolvedValue(null);
  let now = 0;
  spyOn(Date, 'now').mockImplementation(() => now += 1_000);
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual({
    ok: false, version: 'worker-v',
    failure: { code: 'layout-unavailable', target: null, message: 'The fonts this document uses are not loaded yet.' },
  });
  expect(host.editor.readLayoutRequest).toHaveBeenCalledTimes(3);
  expect(worker.documentRead).not.toHaveBeenCalled();
  expectWorkerPageExport(host);
});

for (const code of ['stale-document', 'stale-layout', 'layout-unavailable', 'unsupported-revision-layout'] as const) {
  test(`viewer paged export waits for its own layout after ${code}`, async () => {
    const waiting = deferred<DisplayList>();
    const posted = deferred<void>();
    const settled = mock(() => { posted.resolve(); return waiting.promise; });
    const host = apiFor(true, true, settled);
    const worker = workerFor(host);
    worker.documentRead
      .mockResolvedValueOnce({ version: 'worker-v', value: JSON.stringify({ ...PAGE_REFUSAL, failure: { ...PAGE_REFUSAL.failure, code } }) })
      .mockResolvedValueOnce({ version: 'worker-v', value: JSON.stringify(PAGE_EXPORT) });
    const result = host.api.exportStructuredWithPages(PAGE_OPTIONS);
    await posted.promise;
    expect(settled).toHaveBeenCalledWith(null, 60_000, 'window');
    expect(worker.documentRead).toHaveBeenCalledTimes(1);
    waiting.resolve({} as DisplayList);
    expect(await result).toEqual(PAGE_EXPORT);
    expect(worker.documentRead).toHaveBeenCalledTimes(2);
    expectWorkerPageExport(host);
  });
}

test('viewer paged export retries after its layout wait rejects', async () => {
  const settled = mock(async () => { throw new Error('Layout wait timed out'); });
  const host = apiFor(true, true, settled);
  const worker = workerFor(host);
  worker.documentRead
    .mockResolvedValueOnce({ version: 'worker-v', value: JSON.stringify(PAGE_REFUSAL) })
    .mockResolvedValueOnce({ version: 'worker-v', value: JSON.stringify(PAGE_EXPORT) });
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual(PAGE_EXPORT);
  expect(settled).toHaveBeenCalledWith(null, 60_000, 'window');
  expect(worker.documentRead).toHaveBeenCalledTimes(2);
  expect(host.editor.getLayoutRequest).not.toHaveBeenCalled();
  expect(host.editor.relayout).not.toHaveBeenCalled();
  expect(host.editor.flushPendingInput).not.toHaveBeenCalled();
});

test('viewer paged export does not retry an unsupported revision preview', async () => {
  const settled = mock(async () => ({} as DisplayList));
  const host = apiFor(true, true, settled);
  const worker = workerFor(host);
  host.editor.readLayoutRequest.mockResolvedValue(JSON.stringify({ renderEnv: { revisionPreview: { proposal: 'accept' } } }));
  const refusal = { ...PAGE_REFUSAL, failure: { ...PAGE_REFUSAL.failure, code: 'unsupported-revision-layout' as const } };
  worker.documentRead.mockResolvedValue({ version: 'worker-v', value: JSON.stringify(refusal) });
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual(refusal);
  expect(worker.documentRead).toHaveBeenCalledTimes(1);
  expect(settled).not.toHaveBeenCalled();
  expectWorkerPageExport(host);
});

test('viewer paged export refuses when the worker never lays the document out in time', async () => {
  const host = apiFor(true, true);
  const worker = workerFor(host, WORKER_IDENTITIES, () => new Promise<void>(() => {}));
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void) => {
    callback();
    return 0;
  }) as unknown as typeof setTimeout);
  try {
    expect(await host.api.exportStructuredWithPages({ ...PAGE_OPTIONS, expectLayoutVersion: 'layout-v' })).toEqual({
      ok: false, version: 'v',
      failure: { code: 'layout-unavailable', target: null, message: 'The document is not laid out yet.' },
    });
  } finally {
    timeout.mockRestore();
  }
  expect(worker.documentRead).not.toHaveBeenCalled();
  expect(host.editor.readLayoutRequest).not.toHaveBeenCalled();
  expectWorkerPageExport(host);
});

test('viewer paged export with an expected layout version makes exactly one worker read', async () => {
  const settled = mock(async () => ({} as DisplayList));
  const host = apiFor(true, true, settled);
  const worker = workerFor(host);
  worker.documentRead.mockResolvedValue({ version: 'worker-v', value: JSON.stringify(PAGE_REFUSAL) });
  const options = { ...PAGE_OPTIONS, expectLayoutVersion: 'layout-v' };
  expect(await host.api.exportStructuredWithPages(options)).toEqual(PAGE_REFUSAL);
  expect(worker.documentRead).toHaveBeenCalledWith({ kind: 'exportStructuredWithPages', options, currentRequest: PAGE_REQUEST });
  expect(worker.documentRead).toHaveBeenCalledTimes(1);
  expect(settled).not.toHaveBeenCalled();
  expectWorkerPageExport(host);
});

test('viewer paged export rejects a session replacement while waiting for layout', async () => {
  const host = apiFor(true, true, async () => {
    host.pagedEditorRef.current = null;
    return {} as DisplayList;
  });
  const worker = workerFor(host);
  worker.documentRead.mockResolvedValue({ version: 'worker-v', value: JSON.stringify(PAGE_REFUSAL) });
  await expect(host.api.exportStructuredWithPages(PAGE_OPTIONS)).rejects.toThrow('The document changed while it was being laid out');
  expect(worker.documentRead).toHaveBeenCalledTimes(1);
  expectWorkerPageExport(host);
});

test('viewer paged export returns the main result after hand-over without more worker attempts', async () => {
  const settled = mock(async () => ({} as DisplayList));
  const host = apiFor(true, true, settled);
  const worker = workerFor(host);
  await worker.authority.initialize();
  host.hydrate.mockImplementation(async () => {
    const handover = await beginWorkerProposalHandover(host.session)!;
    return () => { host.state.version = 'main-v'; handover.complete(); };
  });
  const ready = requestWorkerOpenReplica(host.session)!;
  spyOn(host.session, 'exportStructuredWithPagesFor').mockReturnValue({
    ...PAGE_REFUSAL, failure: { ...PAGE_REFUSAL.failure, code: 'unsupported-revision-layout' },
  });
  const result = host.api.exportStructuredWithPages(PAGE_OPTIONS);
  await ready;
  expect(await result).toEqual({ ...PAGE_REFUSAL, failure: { ...PAGE_REFUSAL.failure, code: 'unsupported-revision-layout' } });
  expect(worker.documentRead).not.toHaveBeenCalled();
  expect(host.editor.readLayoutRequest).not.toHaveBeenCalled();
  expect(host.session.exportStructuredWithPagesFor).toHaveBeenCalledTimes(2);
  expect(host.editor.flushPendingInput).toHaveBeenCalledTimes(1);
  expect(settled).not.toHaveBeenCalled();
});

test('viewer paged export uses a main-thread copy that is already loaded', async () => {
  const host = apiFor(true);
  const worker = workerFor(host);
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual(PAGE_EXPORT);
  expect(host.session.exportStructuredWithPagesFor).toHaveBeenCalledWith(PAGE_OPTIONS, PAGE_REQUEST);
  expect(host.editor.readLayoutRequest).not.toHaveBeenCalled();
  expect(worker.documentRead).not.toHaveBeenCalled();
});

test('editor paged export still flushes and reads the main session layout', async () => {
  const host = apiFor();
  const worker = workerFor(host);
  expect(await host.api.exportStructuredWithPages(PAGE_OPTIONS)).toEqual(PAGE_EXPORT);
  expect(host.session.exportStructuredWithPagesFor).toHaveBeenCalledWith(PAGE_OPTIONS, PAGE_REQUEST);
  expect(host.editor.getLayoutRequest).toHaveBeenCalledTimes(1);
  expect(host.editor.readLayoutRequest).not.toHaveBeenCalled();
  expect(worker.documentRead).not.toHaveBeenCalled();
  expect(host.events).toEqual(['flush']);
});

test('editor findText still flushes and reads the main session', async () => {
  const host = apiFor();
  const request: DocxFindTextRequest = { text: 'hello', within: { kind: 'story', story: 'body' }, view: 'accepted' };
  expect(await host.api.findText(request)).toEqual(host.session.findText(request));
  expect(host.events).toEqual(['flush']);
});

for (const [member, args, use] of [
  ['getDocument', [], 'readParagraphs or exportStructuredWithPages'],
  ['getPageContent', [1], 'exportStructuredWithPages'],
  ['findInDocument', ['hello'], 'findParagraphs'],
] as const) {
  test(`${member} refuses worker viewers without requesting a replica and evaluates the session at call time`, () => {
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const host = apiFor(true, true);
    for (let attempt = 0; attempt < 2; attempt += 1) {
      let caught: unknown;
      try { Reflect.apply(host.api[member], host.api, args); } catch (error) { caught = error; }
      expect(caught).toBeInstanceOf(DocxAsyncOnlyError);
      expect((caught as DocxAsyncOnlyError).member).toBe(member);
      expect((caught as DocxAsyncOnlyError).use).toBe(use);
      expect((caught as Error).message).toBe(`${member} cannot read the document synchronously in a viewer session; use ${use}`);
    }
    expect(warning).toHaveBeenCalledTimes(1);
    expect(host.replica!.started).toBe(false);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.fallback).not.toHaveBeenCalled();
    expect(host.request).not.toHaveBeenCalled();
    host.state.viewer = false;
    Reflect.apply(host.api[member], host.api, args);
    expect(host.fallback).toHaveBeenCalledTimes(1);
    expect(warning).toHaveBeenCalledTimes(1);
  });
}

test('a host retrying DocxReplicaNotReadyError after flushPendingInput gets DocxAsyncOnlyError from a viewer at once', async () => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(true);
  for (const [member, args] of [['getDocument', []], ['getPageContent', [1]], ['findInDocument', ['hello']]] as const) {
    let attempts = 0;
    let surfaced: unknown;
    while (surfaced === undefined && attempts < 5) {
      attempts += 1;
      try {
        Reflect.apply(host.api[member], host.api, args);
      } catch (error) {
        if (error instanceof DocxReplicaNotReadyError) await host.api.flushPendingInput();
        else surfaced = error;
      }
    }
    expect(surfaced).toBeInstanceOf(DocxAsyncOnlyError);
    expect(attempts).toBe(1);
  }
  expect(host.editor.flushPendingInput).not.toHaveBeenCalled();
});

const NAVIGATION = [
  ['scrollToParaId', 'scrollToParagraph', ['p', { highlight: { color: 'red' } }]],
  ['scrollToCommentId', 'scrollToComment', [1]],
  ['scrollToChangeId', 'scrollToChange', [2]],
] as const;
for (const [member, twin, args] of NAVIGATION) {
  test(`${member} starts ${twin} immediately, returns true and drops a rejection`, async () => {
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const legacy = mock(() => false);
    const asyncCall = mock(async () => { throw new Error('gone'); });
    const api = routeViewerRefAccess({ [member]: legacy, [twin]: asyncCall } as unknown as DocxEditorRef, () => true);
    expect(Reflect.apply(api[member], api, args)).toBe(true);
    expect(asyncCall).toHaveBeenCalledWith(...args);
    expect(legacy).not.toHaveBeenCalled();
    Reflect.apply(api[member], api, args);
    await Promise.resolve();
    expect(warning).toHaveBeenCalledTimes(1);
  });
}

const PASS_THROUGH = [
  'getPositionAtPoint', 'getSelectionInfo', 'proposeChange', 'highlightRange',
  'getComments', 'readParagraphs', 'onDocumentChange',
] as const;
for (const member of PASS_THROUGH) {
  test(`${member} passes through the gated implementation in both session kinds`, () => {
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const value = {};
    const call = mock(() => value);
    let viewer = false;
    const api = routeViewerRefAccess({ [member]: call } as unknown as DocxEditorRef, () => viewer);
    expect(Reflect.apply(api[member] as Function, api, ['argument'])).toBe(value);
    expect(warning).not.toHaveBeenCalled();
    viewer = true;
    expect(Reflect.apply(api[member] as Function, api, ['argument'])).toBe(value);
    expect(Reflect.apply(api[member] as Function, api, ['argument'])).toBe(value);
    expect(call).toHaveBeenCalledTimes(3);
    const deprecated = !['highlightRange', 'getComments', 'resolveComment', 'readParagraphs', 'onDocumentChange'].includes(member);
    expect(warning).toHaveBeenCalledTimes(deprecated ? 1 : 0);
  });
}

const REFUSALS = [
  ['getEditorRef', [], null, true],
  ['setParagraphStyle', [{ paraId: 'p', styleId: 'Normal' }], false, true],
  ['applyFormatting', [{ paraId: 'p', search: 'hello', marks: { bold: true } }], false, true],
  ['insertBreak', [{ paraId: 'p', type: 'page' }], false, true],
  ['addComment', [{ paraId: 'p', search: 'hello', text: 'Comment', author: 'Author' }], null, true],
  ['replyToComment', [1, 'Reply', 'Author'], null, true],
  ['insertComment', [{ paraId: 'p', search: 'hello', text: 'Comment', author: 'Author' }], null, false],
  ['insertCommentReply', [1, 'Reply', 'Author'], null, false],
  ['resolveComment', [1], undefined, false],
] as const;

for (const [member, args, value, deprecated] of REFUSALS) {
  test(`viewer ${member} refuses before the gated implementation and warns once per page`, async () => {
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const gated = mock(() => 'editor-result');
    let viewer = false;
    const api = routeViewerRefAccess({ [member]: gated } as unknown as DocxEditorRef, () => viewer);
    expect(Reflect.apply(api[member] as Function, api, args)).toBe('editor-result');
    expect(gated).toHaveBeenCalledWith(...args);
    expect(warning).not.toHaveBeenCalled();
    viewer = true;
    const second = routeViewerRefAccess({ [member]: gated } as unknown as DocxEditorRef, () => true);
    for (const ref of [api, api, second]) {
      const result = Reflect.apply(ref[member] as Function, ref, args);
      if (member === 'insertComment' || member === 'insertCommentReply') expect(result).toBeInstanceOf(Promise);
      else expect(result).toBe(value);
      expect(await result).toBe(value);
    }
    expect(gated).toHaveBeenCalledTimes(1);
    expect(warning).toHaveBeenCalledTimes(deprecated ? 1 : 0);
    if (deprecated) {
      expect(warning.mock.calls[0]![0]).toContain(`[DocxEditor] ${member} is deprecated; returns ${String(value)} in viewer sessions.`);
    }
    if (member === 'getEditorRef') {
      for (const twin of ['getParagraphIdentities', 'resolveParagraphAnchors', 'readParagraphs', 'readSelectionInfo', 'findParagraphs', 'exportStructuredWithPages', 'proposeChanges']) {
        expect(warning.mock.calls[0]![0]).toContain(twin);
      }
    }
  });

  test(`viewer ${member} never hydrates or mutates the main-thread document`, async () => {
    spyOn(console, 'warn').mockImplementation(() => {});
    const host = apiFor(true, true);
    const result = Reflect.apply(host.api[member] as Function, host.api, args);
    if (member !== 'insertComment' && member !== 'insertCommentReply') expect(result).toBe(value);
    expect(await result).toBe(value);
    expect(host.setComments).not.toHaveBeenCalled();
    expect(host.sidebar).not.toHaveBeenCalled();
    for (const mutation of ['commentTextTarget', 'formatTextTarget', 'applyParagraphStyle', 'insertPageBreak'] as const) {
      expect(host.session[mutation]).not.toHaveBeenCalled();
    }
    expect(host.events).toEqual([]);
    expectNoReplica(host);
  });
}

test('viewer ref refusals remain active after worker read routing ends', async () => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(false, true, undefined, false, true);
  for (const [member, args, value] of REFUSALS) {
    expect(await Reflect.apply(host.api[member] as Function, host.api, args)).toBe(value);
  }
  expect(host.setComments).not.toHaveBeenCalled();
  expect(host.events).toEqual([]);
  expectNoReplica(host);
});

test('editor refusal members preserve synchronous edits and comment state', () => {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor();
  expect(host.api.getEditorRef() as unknown).toBe(host.editor);
  expect(host.api.setParagraphStyle({ paraId: 'p', styleId: 'Normal' })).toBe(true);
  expect(host.api.applyFormatting({ paraId: 'p', search: 'hello', marks: { bold: true } })).toBe(true);
  expect(host.api.insertBreak({ paraId: 'p', type: 'page' })).toBe(true);
  expect(host.api.addComment({ paraId: 'p', search: 'hello', text: 'Comment', author: 'Author' })).toBeNumber();
  expect(host.api.replyToComment(1, 'Reply', 'Author')).toBeNumber();
  host.api.resolveComment(1);
  expect(host.session.applyParagraphStyle).toHaveBeenCalledTimes(1);
  expect(host.session.formatTextTarget).toHaveBeenCalledTimes(1);
  expect(host.session.insertPageBreak).toHaveBeenCalledTimes(1);
  expect(host.session.commentTextTarget).toHaveBeenCalledTimes(1);
  expect(host.setComments).toHaveBeenCalledTimes(3);
  const resolve = host.setComments.mock.calls[2]![0] as unknown as (comments: Array<{ id: number }>) => unknown;
  expect(resolve([{ id: 1 }, { id: 2 }])).toEqual([{ id: 1, done: true }, { id: 2 }]);
  expect(warning).not.toHaveBeenCalled();
});

test('editor twins flush before reading or navigating and preserve synchronous results', async () => {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor();
  expect(host.api.getDocument()).toBe(host.pagedEditorRef.current!.getDocument());
  expect(host.api.getPageContent(1)).toBeNull();
  expect(await host.api.readSelectionInfo()).toEqual(host.api.getSelectionInfo());
  expect(host.events.slice(0, 2)).toEqual(['flush', 'selection']);
  expect(await host.api.findParagraphs('hello')).toEqual(host.api.findInDocument('hello'));
  for (const [member, twin, args] of NAVIGATION) {
    host.events.length = 0;
    expect(await Reflect.apply(host.api[twin], host.api, args)).toBe(Reflect.apply(host.api[member], host.api, args));
    expect(host.events[0]).toBe('flush');
  }
  host.events.length = 0;
  expect(await host.api.insertComment({ paraId: 'p', text: 'Comment', author: 'Author' })).toBeNumber();
  expect(host.events).toEqual(['flush', 'sync']);
  host.events.length = 0;
  expect(await host.api.insertCommentReply(1, 'Reply', 'Author')).toBeNumber();
  expect(host.events).toEqual(['flush']);
  expect(warning).not.toHaveBeenCalled();
});

test('unavailable editor twins have nothing to flush', async () => {
  const host = apiFor();
  host.pagedEditorRef.current = null;
  expect(await host.api.readSelectionInfo()).toBeNull();
  expect(await host.api.findParagraphs('hello')).toEqual([]);
  expect(await host.api.scrollToParagraph('p')).toBe(false);
  expect(await host.api.scrollToComment(1)).toBe(false);
  expect(await host.api.scrollToChange(1)).toBe(false);
  expect(await host.api.insertComment({ paraId: 'p', text: '', author: '' })).toBeNull();
});

test('viewer synchronous selection stays null without asking for a replica', () => {
  spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor(true, true);
  expect(host.api.getSelectionInfo()).toBeNull();
  expect(host.request).not.toHaveBeenCalled();
  expect(host.hydrate).not.toHaveBeenCalled();
});

test('editor twins propagate failed input flushes', async () => {
  const host = apiFor();
  host.editor.flushPendingInput.mockImplementation(async () => { throw new Error('input failed'); });
  await expect(host.api.readSelectionInfo()).rejects.toThrow('input failed');
  await expect(host.api.findParagraphs('hello')).rejects.toThrow('input failed');
  await expect(host.api.scrollToParagraph('p')).rejects.toThrow('input failed');
  expect(host.editor.scrollToParaId).not.toHaveBeenCalled();
});

test('worker twins read through PagedEditor and authority without requesting a replica', async () => {
  const host = apiFor(true, true);
  const reads: ResidentDocumentRead[] = [];
  const snapshot: ResidentProposalReply = {
    mirror: { version: 'v', proposals: { previewVersion: 0, entries: [] } },
    changedStories: [], geometry: { version: 'v', previewVersion: 0, proposals: '', targets: {}, hidden: [] },
    updates: [], stateVector: new Uint8Array(),
  };
  registerWorkerProposalAuthority(host.session, {
    proposal: async () => snapshot,
    documentRead: (async (read: ResidentDocumentRead) => { reads.push(read); return { version: 'v', value: MATCHES }; }) as ResidentEngineWorkerClient['documentRead'],
    handOver: async () => ({ state: new Uint8Array(), version: 'v', proposals: snapshot.mirror.proposals }),
  }, { relayout: () => {}, current: () => true, laidOut: async () => {}, adopted: () => {}, handedOver: () => {}, contentChanged: () => {} });
  expect(await host.api.findParagraphs('hello', { limit: 3 })).toEqual(MATCHES);
  expect(reads).toEqual([{ kind: 'findParagraphs', query: 'hello', limit: 3 }]);
  expect(await host.api.readSelectionInfo()).toEqual(INFO);
  expect(await host.api.readPositionAtPoint(1, 2)).toBeNull();
  await host.api.scrollToParagraph('p', { highlight: { color: 'red' } });
  await host.api.scrollToComment(1);
  await host.api.scrollToChange(2);
  expect(host.editor.navigateViewer.mock.calls).toEqual([
    [{ kind: 'paragraphTarget', paraId: 'p' }, { highlight: { color: 'red' } }],
    [{ kind: 'commentTarget', commentId: '1' }],
    [{ kind: 'revisionTarget', revisionId: '2' }],
  ]);
  expect(await host.api.insertCommentReply(1, 'Reply', 'Author')).toBeNull();
  expect(host.setComments).not.toHaveBeenCalled();
  expect(host.events).toEqual([]);
  expect(host.request).not.toHaveBeenCalled();
  expect(host.replica!.started).toBe(false);
});

test('content registration warns once per page and version subscriptions unsubscribe', () => {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  apiFor(false).api.onContentChange(() => {});
  expect(warning).not.toHaveBeenCalled();
  const first = apiFor(true);
  first.api.onContentChange(() => {});
  const second = apiFor(true);
  second.api.onContentChange(() => {});
  expect(warning).toHaveBeenCalledTimes(1);
  const listener = mock(() => {});
  const unsubscribe = first.api.onDocumentChange(listener);
  for (const subscriber of first.subscribers) subscriber({ version: 'v2' });
  expect(listener).toHaveBeenCalledWith({ version: 'v2' });
  unsubscribe();
  expect(first.subscribers.size).toBe(0);
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function readAccess(answer: (read: ResidentDocumentRead) => Promise<{ version: string; value: unknown }>) {
  const host = document.createElement('div');
  const frame = (version: string) => {
    const queries = { displayList: {} } as DisplayListQueries;
    stampWorkerFrameVersion(queries, version);
    markPresented(host, queries.displayList);
    return queries;
  };
  const state = { queries: frame('v'), current: true, selection: { anchor: 1, head: 1 } as { anchor: number; head: number } | null };
  const reads: ResidentDocumentRead[] = [];
  const access: ViewerRefReadAccess = {
    read: (async (read: ResidentDocumentRead) => { reads.push(read); return answer(read); }) as ResidentEngineWorkerClient['documentRead'],
    story: 'body', host: () => host, queries: () => state.queries,
    awaitFrame: async () => { state.queries = frame('v2'); return state.queries; },
    current: () => state.current, selection: () => state.selection,
  };
  return { access, state, reads, frame };
}

test('selection info includes a collapsed caret and avoids reading a cleared selection', async () => {
  const host = readAccess(async () => ({ version: 'v', value: INFO }));
  expect(await readViewerSelectionInfo(host.access)).toEqual(INFO);
  expect(host.reads).toEqual([{ kind: 'selectionInfo', story: 'body', anchor: 1, head: 1, expectVersion: 'v' }]);
  host.state.selection = null;
  expect(await readViewerSelectionInfo(host.access)).toBeNull();
  expect(host.reads.length).toBe(1);
});

test('selection info re-reads the newly presented version after supersession', async () => {
  const host = readAccess(async (read) => ({ version: 'expectVersion' in read && read.expectVersion === 'v' ? 'superseded' : 'v2', value: INFO }));
  expect(await readViewerSelectionInfo(host.access)).toEqual(INFO);
  expect(host.reads.map((read) => 'expectVersion' in read && read.expectVersion)).toEqual(['v', 'v2']);
});

for (const cleared of [false, true]) {
  test(`selection info ${cleared ? 'returns null when cleared' : 'reads a changed selection'} during the await`, async () => {
    const pending = deferred<{ version: string; value: unknown }>();
    const host = readAccess(async () => host.reads.length === 1 ? pending.promise : { version: 'v', value: INFO });
    const result = readViewerSelectionInfo(host.access);
    host.state.selection = cleared ? null : { anchor: 2, head: 5 };
    pending.resolve({ version: 'v', value: INFO });
    expect(await result).toEqual(cleared ? null : INFO);
    if (!cleared) expect(host.reads[1]).toEqual({ kind: 'selectionInfo', story: 'body', anchor: 2, head: 5, expectVersion: 'v' });
  });
}

const TARGETS: ViewerNavigationTarget[] = [
  { kind: 'paragraphTarget', paraId: 'p' }, { kind: 'commentTarget', commentId: '1' }, { kind: 'revisionTarget', revisionId: '2' },
];
for (const target of TARGETS) {
  test(`${target.kind} re-reads superseded frames, then selects and reveals`, async () => {
    const host = readAccess(async (read) => ({ version: 'expectVersion' in read && read.expectVersion === 'v' ? 'old' : 'v2', value: { anchor: 2, head: 5 } }));
    const apply = mock(() => {});
    expect(await navigateViewer(host.access, target, apply, { highlight: { color: 'red' } })).toBe(true);
    expect(host.reads).toEqual([{ ...target, story: 'body', expectVersion: 'v' }, { ...target, story: 'body', expectVersion: 'v2' }]);
    expect(apply).toHaveBeenCalledWith({ anchor: 2, head: 5 }, { highlight: { color: 'red' } });
  });
  test(`${target.kind} drops a reply when its navigation context is invalidated`, async () => {
    const pending = deferred<{ version: string; value: unknown }>();
    const host = readAccess(() => pending.promise);
    const apply = mock(() => {});
    const result = navigateViewer(host.access, target, apply);
    host.state.current = false;
    pending.resolve({ version: 'v', value: { anchor: 1, head: 5 } });
    expect(await result).toBe(false);
    expect(apply).not.toHaveBeenCalled();
  });
  test(`${target.kind} re-reads a reply whose frame was replaced while it waited`, async () => {
    const pending = deferred<{ version: string; value: unknown }>();
    const host = readAccess(async () => host.reads.length === 1 ? pending.promise : { version: 'v2', value: { anchor: 4, head: 8 } });
    const apply = mock(() => {});
    const result = navigateViewer(host.access, target, apply);
    host.state.queries = host.frame('v2');
    pending.resolve({ version: 'v', value: { anchor: 1, head: 5 } });
    expect(await result).toBe(true);
    expect(apply).toHaveBeenCalledWith({ anchor: 4, head: 8 }, undefined);
  });
  test(`${target.kind} returns false for a missing target`, async () => {
    const host = readAccess(async () => ({ version: 'v', value: null }));
    expect(await navigateViewer(host.access, target, () => { throw new Error('unexpected navigation'); })).toBe(false);
  });
}

test('superseded reads are bounded', async () => {
  const host = readAccess(async () => ({ version: 'old', value: null }));
  expect(await navigateViewer(host.access, TARGETS[0]!, () => {})).toBe(false);
  expect(host.reads.length).toBe(5);
});


test('a newer paged navigation invalidates a pending worker navigation across ref rebuilds', async () => {
  const pending = deferred<void>();
  const bump = mock(() => {});
  const hook = renderHook(() => {
    const ref = useRef<PagedEditorRef>(null);
    usePagedEditorRefApi({
      ref, viewerSelection: true, bumpInputEpoch: bump, yrsInputRef: { current: null },
      layout: null, yrsSession: {} as YrsSession, documentFromYrs: () => null,
      runLayoutPipeline: () => {}, getLayoutRequest: () => null, readLayoutRequest: async () => null,
      scrollToPositionImpl: () => {}, revealPositionImpl: () => 'scrolled',
      scrollToParaIdImpl: () => false, scrollToPageImpl: () => {},
      setIsFocused: () => {}, onReadyRef: { current: undefined },
      yrsLocToDisplayPosition: () => null, syncYrsInputState: () => true,
      applyYrsFormatting: () => false, applyYrsCommand: () => false,
      getYrsPositionProjection: () => null, displayPositionToYrsLoc: () => null,
      getPositionAtPoint: () => null,
      navigateViewer: async (_target, _options, current) => {
        await pending.promise;
        return current?.() ?? false;
      },
    });
    return ref;
  });
  const first = hook.result.current.current!.navigateViewer(TARGETS[0]!);
  hook.rerender();
  hook.result.current.current!.scrollToPage(2);
  pending.resolve();
  expect(await first).toBe(false);
  const second = hook.result.current.current!.navigateViewer(TARGETS[1]!);
  expect(await second).toBe(true);
  expect(bump).not.toHaveBeenCalled();
});
