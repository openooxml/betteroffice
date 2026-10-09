import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef, type ReactNode } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsSession,
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  yrsToDocument,
  proposalSetIdentity,
  type ResidentProposalReply,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { Document } from '@betteroffice/docx/types/document';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { YrsInput, type YrsInputRef } from '../YrsInput';
import { createCommentIdAllocator } from '../commentFactories';
import * as workerOpenReplica from '../internals/workerOpenReplica';
import { deferWorkerOpenReplica, holdWorkerOpenDocument, type WorkerOpenFallbackReason } from '../internals/workerOpenReplica';
import { beginWorkerProposalHandover, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { EditorMode } from '../internals/editing-modes';
import { DOCX_REF_ASYNC_TWINS, DOCX_REF_REPLICA_ACCESS, DOCX_REF_REPLICA_LOADING_MUTATIONS, DocxAsyncOnlyError, DocxReplicaNotReadyError, useDocxEditorRefApi } from './useDocxEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const ROOT = resolve(import.meta.dir, '../../../../../..');
const bytes = new Uint8Array(readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx')));
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))));
afterEach(() => {
  cleanup();
  mock.restore();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function apiFor(
  session: YrsSession,
  document: Document,
  mode: EditorMode = 'viewing',
  replicaReadyRef?: { current: boolean },
  viewerSession = false
) {
  const events: string[] = [];
  const inputRef = { current: null as YrsInputRef | null };
  const openingRef = { current: false };
  const project = () => {
    const base = session.materializeDocx();
    return base ? yrsToDocument(session, base) : null;
  };
  const editor = {
    isWorkerViewer: () => viewerSession,
    getYrsSession: () => session,
    getDocument: project,
    flushPendingInput: async () => {
      events.push('flush');
      await inputRef.current?.flushPendingInput();
    },
    insertText: (text: string) => inputRef.current?.insertText(text),
    syncYrsInputState: () => { events.push('sync'); return true; },
    getLayout: () => null,
    getLayoutRequest: () => null, readLayoutRequest: async () => null,
    scrollToPosition: () => {},
    getPositionAtPoint: () => null,
    displayPositionToYrsLoc: () => null,
    scrollToParaId: () => session.hasStory('body'),
    scrollToCommentId: () => false,
    scrollToChangeId: () => false,
    highlightRange: () => {},
    focus: () => {},
  } satisfies Partial<PagedEditorRef>;
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      experimentalWorkerOpen: true,
      viewerSession,
      ref,
      document,
      documentFromYrs: project,
      historyStateRef: { current: document },
      pagedEditorRef,
      openingRef,
      handleSave: async () => {
        events.push('save');
        return new ArrayBuffer(0);
      },
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
      getCachedStyleResolver: createStyleResolver,
      hostSearch: {
        search: async () => ({ query: '', options: { caseSensitive: false }, total: 0, current: -1 }),
        searchNext: () => null,
        searchPrevious: () => null,
        searchGoTo: () => null,
        clearSearch: () => {},
        getSearchState: () => null,
        onSearchChange: () => () => {},
      },
      commentIdAllocator: createCommentIdAllocator(),
      commands: UNAVAILABLE_DOCX_COMMANDS,
      modeRef: { current: mode },
      allowHostProposalsRef: { current: false },
      settledDisplayList: async () => ({ pages: [] }),
    });
    return ref;
  }, {
    wrapper: replicaReadyRef ? ({ children }: { children: ReactNode }) => {
      const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
      return (
        <div>
          {children}
          <YrsInput
            ref={inputRef}
            enabled
            readOnly={mode === 'viewing'}
            replicaReadyRef={replicaReadyRef}
            session={session}
            inputPositionMap={map}
            displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
            locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
            onStateChange={() => {}}
            onDirectInput={() => {}}
          />
        </div>
      );
    } : undefined,
  });
  const api = hook.result.current.current;
  if (!api) throw new Error('The ref API is not mounted');
  return { api, events, pagedEditorRef, openingRef };
}

async function pendingReplica(mode: EditorMode = 'viewing', mountInput = false, waitForLayout = false) {
  const worker = await createYrsSession();
  const session = await createYrsSession();
  sessions.push(worker, session);
  const { document } = worker.openDocx(bytes, true);
  const state = worker.encodeState();
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  const opens: boolean[] = [];
  const fallbackReasons: WorkerOpenFallbackReason[] = [];
  const readiness = { current: false };
  const replica = deferWorkerOpenReplica(
    session,
    async () => {
      const handover = await beginWorkerProposalHandover(session);
      await held;
      return () => {
        opens.push(false);
        session.openDocx(bytes, false);
        session.loadState(handover?.state ?? state);
        handover?.complete();
      };
    },
    (reason) => {
      fallbackReasons.push(reason);
      opens.push(true);
      session.openDocx(bytes, true);
    },
    () => { readiness.current = true; },
    { current: () => true, cancel: () => {}, waitForLayout }
  );
  const mounted = apiFor(session, document, mode, mountInput ? readiness : undefined);
  return { ...mounted, session, worker, replica, release, opens, fallbackReasons };
}

function expectLoadingMutations(api: DocxEditorRef, editor: PagedEditorRef) {
  const session = editor.getYrsSession()!;
  const started = workerOpenReplica.workerOpenReplicaStarted(session);
  const request = spyOn(workerOpenReplica, 'requestWorkerOpenReplica');
  const requests = request.mock.calls.length;
  const ensure = spyOn(workerOpenReplica, 'ensureWorkerOpenReplica');
  const admissions = ensure.mock.calls.length;
  for (const member of DOCX_REF_REPLICA_LOADING_MUTATIONS) {
    const call = SYNC_REPLICA_CALLS.find(([name]) => name === member);
    if (!call) throw new Error(`Missing loading mutation: ${member}`);
    expect(() => Reflect.apply(api[member] as Function, api, call[1])).toThrow(DocxReplicaNotReadyError);
  }
  const navigate = spyOn(editor, 'scrollToParaId');
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(api.findInDocument('Page')).toEqual([]);
  expect(api.scrollToParaId('00000001')).toBe(false);
  expect(navigate).not.toHaveBeenCalled();
  expect(workerOpenReplica.workerOpenReplicaPending(session)).toBe(true);
  expect(workerOpenReplica.workerOpenReplicaStarted(session)).toBe(started);
  expect(request.mock.calls).toHaveLength(requests);
  expect(ensure.mock.calls).toHaveLength(admissions);
}

function expectReadyMutations(api: DocxEditorRef, session: YrsSession, editor: PagedEditorRef) {
  expect(api.getDocument()).not.toBeNull();
  expect(api.getEditorRef()).toBe(editor);
  const first = session.paragraphs('body')[0]!;
  expect(api.findInDocument(first.text)).toContainEqual(expect.objectContaining({ paraId: first.paraId, match: first.text }));
  expect(api.scrollToParaId(first.paraId)).toBe(true);
  expect(editor.scrollToParaId).toHaveBeenCalledWith(first.paraId, undefined);
  const last = session.paragraphs('body').at(-1)!;
  const search = 'Replica contract';
  let paraId!: string;
  act(() => {
    const span = session.locateParagraph('body', last.paraId);
    paraId = session.splitParagraph({ story: 'body', paraId: last.paraId, offset: span.end - span.start }).secondParaId;
    session.insertText({ story: 'body', paraId, offset: 0 }, search);
    session.setParagraphAttr(paraId, 'pStyle', 'Heading1');
  });
  act(() => {
    const comment = api.addComment({ paraId, search, text: 'Check', author: 'Ann' });
    expect(comment).toEqual(expect.any(Number));
    expect(session.resolveComment(String(comment))).toContainEqual(expect.objectContaining({ story: 'body' }));
    expect(api.applyFormatting({ paraId, search, marks: { italic: true } })).toBe(true);
    expect(session.storySegments('body')).toContainEqual(expect.objectContaining({
      kind: 'text', text: search, attributes: expect.objectContaining({ italic: true }),
    }));
    expect(api.setParagraphStyle({ paraId, styleId: 'Normal' })).toBe(true);
    expect(session.paragraphs('body').find((entry) => entry.paraId === paraId)!.properties.pStyle).toBe('Normal');
    expect(api.proposeChange({ paraId, search: '', replaceWith: 'Replica ready', author: 'Host' })).toBe(true);
    expect(session.listRevisions()).toContainEqual(expect.objectContaining({
      kind: 'insertion', author: 'Host', preview: 'Replica ready',
    }));
    const breaks = session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak').length;
    expect(api.insertBreak({ paraId, type: 'page' })).toBe(true);
    expect(session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak')).toHaveLength(breaks + 1);
  });
}

async function pendingWorkerProposalReplica(mode: EditorMode = 'viewing') {
  const pending = await pendingReplica(mode, false);
  const { session, worker } = pending;
  let previewVersion = 0;
  const reply = (): ResidentProposalReply => {
    const snapshot = { version: worker.version(), previewVersion, proposals: [] };
    return {
      mirror: { version: snapshot.version, proposals: { previewVersion, entries: [] } },
      result: { ok: true, snapshot },
      changedStories: [],
      geometry: {
        version: snapshot.version, previewVersion,
        proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [],
      },
      updates: [],
      stateVector: new Uint8Array(),
    };
  };
  const transport = {
    proposal: mock(async (op: { kind: string }) => {
      if (op.kind === 'propose') previewVersion += 1;
      return reply();
    }),
    documentRead: async () => { throw new Error('unexpected worker read'); },
    handOver: mock(async () => ({
      state: worker.encodeState(), version: worker.version(), proposals: reply().mirror.proposals,
    })),
  };
  const authority = registerWorkerProposalAuthority(session, transport, {
    relayout: () => {},
    current: () => true,
    laidOut: async () => {},
    adopted: () => {},
    handedOver: () => {},
    contentChanged: () => {},
  });
  await authority.propose({ expectVersion: worker.version(), proposals: [] }, async () => {
    throw new Error('unexpected main proposal');
  });
  expect(authority.holdsWorkerState()).toBe(true);
  expect(pending.replica.started).toBe(false);
  expect(pending.opens).toEqual([]);
  return { ...pending, authority, transport };
}

test('every public ref API is classified for replica access', async () => {
  const { api } = await pendingReplica();
  expect(Object.keys(api).sort()).toEqual(Object.keys(DOCX_REF_REPLICA_ACCESS).sort());
  expect(Object.keys(DOCX_REF_REPLICA_ACCESS).sort()).toEqual([
    'addComment', 'applyEdits', 'applyFormatting', 'clearSearch', 'commands', 'exportStructuredWithPages',
    'findContentControls', 'findInDocument', 'findText', 'flushPendingInput', 'focus',
    'getComments', 'getCurrentPage', 'getDocument', 'getEditorRef', 'getMemoryStats', 'getPageContent', 'getParagraphIdentities',
    'getPositionAtPoint', 'getProposals', 'getSearchState', 'getSelectionInfo', 'getTotalPages', 'getZoom',
    'highlightRange', 'insertBreak', 'listContentControls', 'loadDocument', 'loadDocumentBuffer',
    'onContentChange', 'onSearchChange', 'onSelectionChange', 'openPrintPreview', 'print', 'proposeChange',
    'proposeChanges', 'readParagraphs', 'readPositionAtPoint', 'replyToComment', 'resolveComment', 'resolveParagraphAnchors', 'save',
    'search', 'searchGoTo', 'searchNext', 'searchPrevious',
    'scrollToChangeId', 'scrollToCommentId', 'scrollToPage', 'scrollToParaId', 'scrollToPosition',
    'setParagraphStyle', 'setProposalStates', 'setZoom', 'validateEdits', 'whenLayoutComplete',
    'withdrawProposals', 'readSelectionInfo', 'findParagraphs', 'scrollToParagraph', 'scrollToComment',
    'scrollToChange', 'insertComment', 'insertCommentReply', 'onDocumentChange',
  ].sort());
});

test('async reads, exports and write refusals wait for the main replica while save stays independent', async () => {
  const { api, events, session, replica, release, opens } = await pendingReplica();
  const search = { text: 'Page', within: { kind: 'story', story: 'body' }, view: 'accepted' } as const;
  const edits = { expectVersion: 'before-ready', steps: [] };
  const proposals = { expectVersion: 'before-ready', proposals: [] };
  const states = { expectVersion: 'before-ready', expectPreviewVersion: 0, changes: [] };
  const withdrawal = { expectVersion: 'before-ready', ids: [] };
  const pending = Promise.all([
    api.readParagraphs({ view: 'accepted' }),
    api.listContentControls(),
    api.findContentControls({ kind: 'tag', tag: 'missing' }),
    api.findText(search),
    api.getProposals(),
    api.validateEdits(edits),
    api.applyEdits(edits),
    api.proposeChanges(proposals),
    api.setProposalStates(states),
    api.withdrawProposals(withdrawal),
    api.save(),
    api.flushPendingInput(),
    api.exportStructuredWithPages({ revisionView: 'markup', expectLayoutVersion: 'before-ready' }),
    api.whenLayoutComplete(),
  ]);
  const completed = { value: false };
  void pending.then(() => { completed.value = true; });
  replica.start();
  await act(async () => {});
  expect(completed.value).toBe(false);
  expect(session.storyIds()).toEqual([]);
  expect(events).toEqual(['save']);
  expect(opens).toEqual([]);
  let values!: Awaited<typeof pending>;
  await act(async () => {
    release();
    values = await pending;
  });
  expect(opens).toEqual([false]);
  expect(values[0]).toEqual(session.readParagraphs({ view: 'accepted' }));
  expect(values[1]).toEqual(session.listContentControls());
  expect(values[2]).toEqual(session.findContentControls({ kind: 'tag', tag: 'missing' }));
  expect(values[3]).toEqual(session.findText(search));
  expect(values[4]).toEqual(session.getProposals());
  for (const result of values.slice(5, 10)) {
    expect(result).toMatchObject({ ok: false, version: session.version(), failure: { code: 'read-only' } });
  }
  expect(new TextDecoder().decode(values[10]!)).toBe('');
  expect(values[12]).toMatchObject({ ok: false, version: session.version(), failure: { code: 'layout-unavailable' } });
  expect(values[13]).toBe(0);
});

test.each(['viewing', 'editing'] as const)('ref save in %s mode leaves the held viewer document or deferred editor replica unloaded', async (mode) => {
  if (mode === 'viewing') {
    const host = await heldViewer();
    expect(await host.api.save()).toBeInstanceOf(ArrayBuffer);
    expect(host.events).toEqual(['save']);
    expect(workerOpenReplica.workerOpenDocumentHeld(host.session)).toBe(true);
    expect(workerOpenReplica.workerOpenReplicaPending(host.session)).toBe(true);
    expect(workerOpenReplica.workerOpenReplicaStarted(host.session)).toBe(false);
    expect(host.session.storyIds()).toEqual([]);
    expect(host.release).not.toHaveBeenCalled();
    for (const helper of host.helpers) expect(helper).not.toHaveBeenCalled();
    return;
  }
  const { api, events, session, replica, opens } = await pendingReplica(mode);
  expect(await api.save()).toBeInstanceOf(ArrayBuffer);
  expect(events).toEqual(['save']);
  expect(replica.started).toBe(false);
  expect(replica.pending).toBe(true);
  expect(session.storyIds()).toEqual([]);
  expect(opens).toEqual([]);
});

test('synchronous reads return loading answers until the owner opens the editor replica', async () => {
  const { api, session, replica, release, opens, fallbackReasons, pagedEditorRef } = await pendingReplica('editing');
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(api.findInDocument('Page')).toEqual([]);
  expect(api.scrollToParaId('00000001')).toBe(false);
  expectLoadingMutations(api, pagedEditorRef.current!);
  await act(async () => {});
  expect(opens).toEqual([]);
  expect(replica.started).toBe(false);
  expect(replica.pending).toBe(true);
  replica.start();
  await act(async () => { release(); await replica.ready; });
  expectReadyMutations(api, session, pagedEditorRef.current!);
  expect(api.getDocument()).not.toBeNull();
  expect(api.getEditorRef()?.getYrsSession()).toBe(session);
  const first = session.paragraphs('body')[0]!;
  expect(api.findInDocument(first.text)).toContainEqual({
    paraId: first.paraId, match: first.text, before: '', after: '',
  });
  expect(api.findInDocument('Page')).toContainEqual({
    paraId: first.paraId, match: 'Page', before: '', after: first.text.slice(4),
  });
  expect(api.scrollToParaId(first.paraId)).toBe(true);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
});

test('a batch chained from an early read edits the hydrated document', async () => {
  const { api, session, replica, release } = await pendingReplica('editing');
  const edited = api.readParagraphs({ view: 'accepted' }).then((read) => {
    if (!read.ok) throw new Error(read.failure.message);
    return api.applyEdits({
      expectVersion: read.version,
      steps: [{
        op: 'replaceText',
        target: { kind: 'paragraph', story: 'body', paraId: read.paragraphs[0]!.paraId },
        text: 'Written after opening',
      }],
    });
  });
  replica.start();
  await act(async () => { release(); expect(await edited).toMatchObject({ ok: true, applied: true }); });
  expect(session.paragraphs('body')[0]!.text).toBe('Written after opening');
});

test('insertBreak preserves paragraph text and inserts a lowerable page break', async () => {
  const { api, session, replica, release } = await pendingReplica('editing');
  replica.start();
  await act(async () => { release(); await replica.ready; });
  const before = session.paragraphs('body');
  const original = before[0]!;
  expect(original.text.length).toBeGreaterThan(0);
  const pageBreakCount = () => session.storySegments('body').filter(
    (segment) => segment.kind === 'embed' && segment.embedKind === 'pageBreak'
  ).length;
  const previousPageBreaks = pageBreakCount();

  act(() => expect(api.insertBreak({ paraId: original.paraId, type: 'page' })).toBe(true));

  expect(() => session.yrsBlocksForStory('body')).not.toThrow();
  const after = session.paragraphs('body');
  expect(after).toHaveLength(before.length + 1);
  expect(after[0]!.paraId).toBe(original.paraId);
  expect(after.slice(0, 2).map((paragraph) => paragraph.text)).toEqual([original.text, '']);
  expect(pageBreakCount()).toBe(previousPageBreaks + 1);
});

test('a synchronous write during an in-flight editor handoff throws until readiness', async () => {
  const { api, session, replica, release, opens, fallbackReasons, pagedEditorRef } = await pendingReplica('editing');
  const request = spyOn(workerOpenReplica, 'requestWorkerOpenReplica');
  replica.start();
  await act(async () => {});
  act(() => expect(() => api.insertBreak({ paraId: '00000001', type: 'page' })).toThrow(DocxReplicaNotReadyError));
  expectLoadingMutations(api, pagedEditorRef.current!);
  expect(request).not.toHaveBeenCalled();
  expect(replica.started).toBe(true);
  expect(opens).toEqual([]);
  expect(session.storyIds()).toEqual([]);
  await act(async () => { release(); await replica.ready; });
  expectReadyMutations(api, session, pagedEditorRef.current!);
  expect(opens).toEqual([false]);
  const readyVersion = session.version();
  act(() => expect(api.insertBreak({ paraId: '00000001', type: 'page' })).toBe(true));
  expect(session.version()).not.toBe(readyVersion);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
});

test('replacing the document while a ref waits rejects the pending call', async () => {
  const { api, replica, release, pagedEditorRef } = await pendingReplica();
  const read = api.readParagraphs({ view: 'accepted' });
  const rejected = read.then(
    () => { throw new Error('The pending read should reject'); },
    (error: unknown) => error
  );
  replica.start();
  pagedEditorRef.current = null;
  await act(async () => {
    release();
    const error = await rejected;
    expect(error).toBeInstanceOf(Error);
    expect(error).toMatchObject({ message: expect.stringContaining('document changed') });
  });
});

test('a layout deadline covers the wait for the main replica', async () => {
  const { api, opens } = await pendingReplica();
  const error = await api.whenLayoutComplete({ timeoutMs: 20 }).then(() => null, (failure: unknown) => failure);
  expect(error).toMatchObject({ message: 'The document did not finish rendering' });
  expect(opens).toEqual([]);
});

test('getProposals stays worker-served while proposals are held in the worker', async () => {
  const { api, events, worker, opens, replica, release, transport } = await pendingWorkerProposalReplica();
  let read!: ReturnType<DocxEditorRef['getProposals']>;
  expect(() => { read = api.getProposals(); }).not.toThrow();
  expect(read).toBeInstanceOf(Promise);
  expect(await read).toEqual({ version: worker.version(), previewVersion: 1, proposals: [] });
  expect(events).toEqual([]);
  expect(transport.proposal.mock.calls.map(([op]) => op.kind)).toEqual(['snapshot', 'propose']);
  expect(opens).toEqual([]);
  expect(replica.started).toBe(false);
  await act(async () => { release(); });
  expect(transport.handOver).not.toHaveBeenCalled();
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
});

test('independent APIs do not start a replica open', async () => {
  const { api, opens, replica } = await pendingReplica();
  api.setZoom(2);
  expect(api.getZoom()).toBe(1);
  expect(api.getCurrentPage()).toBe(1);
  expect(api.getTotalPages()).toBe(0);
  expect(api.getComments()).toEqual([]);
  const unsubscribe = api.onContentChange(() => {});
  unsubscribe();
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
});

test('getEditorRef returns null until readiness and then immediately inserts text', async () => {
  const { api, session, replica, release, opens, fallbackReasons, pagedEditorRef } = await pendingReplica('editing', true);
  expect(replica.pending).toBe(true);
  expect(api.getEditorRef()).toBeNull();
  expectLoadingMutations(api, pagedEditorRef.current!);
  expect(replica.started).toBe(false);
  expect(opens).toEqual([]);
  replica.start();
  await act(async () => { release(); await replica.ready; });
  expectReadyMutations(api, session, pagedEditorRef.current!);
  await act(async () => {
    api.getEditorRef()!.insertText('Immediate ');
    await api.flushPendingInput();
  });
  expect(replica.pending).toBe(false);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
  expect(session.paragraphs('body')[0]!.text).toStartWith('Immediate ');
});

async function heldViewer() {
  const worker = await createYrsSession();
  const session = await createYrsSession();
  sessions.push(worker, session);
  const { document } = worker.openDocx(bytes, true);
  const release = mock(() => { throw new Error('unexpected viewer release'); });
  holdWorkerOpenDocument(session, release);
  const mounted = apiFor(session, document, 'viewing', undefined, true);
  const helpers = [
    spyOn(workerOpenReplica, 'requestWorkerOpenReplica'),
    spyOn(workerOpenReplica, 'awaitWorkerOpenReplica'),
    spyOn(workerOpenReplica, 'ensureWorkerOpenReplica'),
  ];
  return { ...mounted, session, release, helpers };
}

test('held viewer save delegates without generic replica admission', async () => {
  const host = await heldViewer();
  expect(await host.api.save()).toBeInstanceOf(ArrayBuffer);
  expect(host.events).toEqual(['save']);
  expect(workerOpenReplica.workerOpenDocumentHeld(host.session)).toBe(true);
  expect(host.release).not.toHaveBeenCalled();
  for (const helper of host.helpers) expect(helper).not.toHaveBeenCalled();
});

test('held viewer layout completion and input flush resolve without a replica', async () => {
  const host = await heldViewer();
  expect(await host.api.whenLayoutComplete({ timeoutMs: 20 })).toBe(0);
  await host.api.flushPendingInput();
  expect(host.events).toEqual(['flush']);
  expect(host.release).not.toHaveBeenCalled();
  for (const helper of host.helpers) expect(helper).not.toHaveBeenCalled();
});

test('held viewer synchronous reads and edit refusals never reach replica helpers', async () => {
  const host = await heldViewer();
  spyOn(console, 'warn').mockImplementation(() => {});
  expect(() => host.api.getDocument()).toThrow(DocxAsyncOnlyError);
  expect(() => host.api.getPageContent(1)).toThrow(DocxAsyncOnlyError);
  expect(() => host.api.findInDocument('text')).toThrow(DocxAsyncOnlyError);
  expect(host.api.getEditorRef()).toBeNull();
  expect(host.api.getSelectionInfo()).toBeNull();
  expect(host.api.applyFormatting({ paraId: 'p', search: 'text', marks: { bold: true } })).toBe(false);
  expect(host.api.proposeChange({ paraId: 'p', search: 'text', replaceWith: 'next', author: 'Host' })).toBe(false);
  host.api.focus();
  host.api.scrollToPosition(0);
  host.api.highlightRange(0, 1);
  expect(host.api.getPositionAtPoint(0, 0)).toBeNull();
  expect(host.release).not.toHaveBeenCalled();
  for (const helper of host.helpers) expect(helper).not.toHaveBeenCalled();
});

const SYNC_REPLICA_CALLS = [
  ['getDocument', [], null, (result: unknown) => expect(result).not.toBeNull()],
  ['getEditorRef', [], null, (result: unknown) => expect(result).not.toBeNull()],
  ['scrollToParaId', ['00000001'], false, (result: unknown) => expect(result).toBe(true)],
  ['scrollToCommentId', [-1], false, (result: unknown) => expect(result).toBe(true)],
  ['scrollToChangeId', [-1], false, (result: unknown) => expect(result).toBe(true)],
  ['findInDocument', ['map'], [], (result: unknown) => expect(result).toContainEqual({
    paraId: '00000001', match: 'map', before: 'Page ', after: '',
  })],
  ['getPageContent', [1], null, (result: unknown) => expect(result).toEqual({
    pageNumber: 1,
    text: '[00000001] Page map',
    paragraphs: [{ paraId: '00000001', text: 'Page map', styleId: 'Heading1' }],
  })],
  ['getSelectionInfo', [], null, (result: unknown) => expect(result).toMatchObject({
    paraId: '00000001', selectedText: 'Page map', paragraphText: 'Page map',
  })],
  ['addComment', [{
    paraId: '00000001', search: 'map', text: 'Check', author: 'Ann',
  }], DocxReplicaNotReadyError, (result: unknown) => expect(result).toEqual(expect.any(Number))],
  ['proposeChange', [{
    paraId: '00000001', search: 'map', replaceWith: 'plan', author: 'Agent',
  }], DocxReplicaNotReadyError, (result: unknown) => expect(result).toBe(true)],
  ['applyFormatting', [{
    paraId: '00000001', search: 'map', marks: { italic: true },
  }], DocxReplicaNotReadyError, (result: unknown) => expect(result).toBe(true)],
  ['setParagraphStyle', [{
    paraId: '00000001', styleId: 'Normal',
  }], DocxReplicaNotReadyError, (result: unknown) => expect(result).toBe(true)],
  ['insertBreak', [{
    paraId: '00000001', type: 'page',
  }], DocxReplicaNotReadyError, (result: unknown) => expect(result).toBe(true)],
] as const;


function prepareSyncRead(editor: PagedEditorRef) {
  editor.getLayout = () => ({
    pages: [{ fragments: [{ kind: 'paragraph', pmStart: 0 }] }],
  }) as unknown as ReturnType<PagedEditorRef['getLayout']>;
  editor.displayPositionToYrsLoc = () => ({ story: 'body', paraId: '00000001', offset: 0 });
  editor.scrollToCommentId = () => true;
  editor.scrollToChangeId = () => true;
}

test('ref flush bypasses the layout wait and drains queued input', async () => {
  const { api, session, replica, release, pagedEditorRef, fallbackReasons } = await pendingReplica('editing', true, true);
  act(() => pagedEditorRef.current!.insertText('Q'));
  replica.start();
  await act(async () => release());
  const before = session.paragraphs('body')[0]!.text;
  expect(replica.pending).toBe(true);
  expect(api.getDocument()).toBeNull();
  await act(async () => api.flushPendingInput());
  expect(replica.pending).toBe(false);
  expect(session.paragraphs('body')[0]!.text).toBe(`Q${before}`);
  expect(api.getDocument()).not.toBeNull();
  expect(fallbackReasons).toEqual([]);
});

test('editor document access returns loading answers and paragraph styling throws until the peer is ready', async () => {
  const { api, session, replica, release, opens, fallbackReasons, pagedEditorRef } = await pendingReplica('editing');
  const getDocument = spyOn(pagedEditorRef.current!, 'getDocument');
  const paragraphs = spyOn(session, 'paragraphs');
  const open = spyOn(session, 'openDocx');
  const style = spyOn(session, 'applyParagraphStyle');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const options = { paraId: '00000001', styleId: 'Normal' };
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(() => api.setParagraphStyle(options)).toThrow(DocxReplicaNotReadyError);
  await act(async () => {});
  expect(replica.started).toBe(false);
  expect(getDocument).not.toHaveBeenCalled();
  expect(paragraphs).not.toHaveBeenCalled();
  expect(open).not.toHaveBeenCalled();
  expect(style).not.toHaveBeenCalled();
  expect(warning).not.toHaveBeenCalled();
  expect(opens).toEqual([]);
  replica.start();
  await act(async () => { release(); await replica.ready; });
  expect(api.getDocument()).toEqual(pagedEditorRef.current!.getDocument());
  expect(api.getDocument()).not.toBeNull();
  expect(api.getEditorRef()).toBe(pagedEditorRef.current);
  act(() => { expect(api.setParagraphStyle(options)).toBe(true); });
  expect(session.paragraphs('body')[0]!.properties.pStyle).toBe('Normal');
  expect(open).toHaveBeenCalledTimes(1);
  expect(open).toHaveBeenCalledWith(bytes, false);
  expect(style).toHaveBeenCalledTimes(1);
  expect(warning).not.toHaveBeenCalled();
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
});

test.each(['focus', 'scrollToPosition', 'print', 'openPrintPreview', 'highlightRange', 'getPositionAtPoint'] as const)(
  '%s calls through without starting a pending editor replica', async (method) => {
    const { api, opens, replica, pagedEditorRef } = await pendingReplica('editing');
    if (method === 'print' || method === 'openPrintPreview') {
      const execute = spyOn(UNAVAILABLE_DOCX_COMMANDS, 'execute');
      api[method]();
      expect(execute).toHaveBeenCalledWith('print', null);
    } else if (method === 'getPositionAtPoint') {
      const position = {} as NonNullable<ReturnType<PagedEditorRef['getPositionAtPoint']>>;
      const point = spyOn(pagedEditorRef.current!, method).mockReturnValue(position);
      expect(api.getPositionAtPoint(0, 0)).toBe(position);
      expect(point).toHaveBeenCalledWith(0, 0);
    } else {
      const call = spyOn(pagedEditorRef.current!, method);
      if (method === 'scrollToPosition') api.scrollToPosition(0);
      else if (method === 'highlightRange') api.highlightRange(0, 1);
      else api.focus();
      expect(call).toHaveBeenCalledTimes(1);
    }
    await act(async () => {});
    expect(opens).toEqual([]);
    expect(replica.started).toBe(false);
    expect(replica.pending).toBe(true);
  }
);

test.each(SYNC_REPLICA_CALLS.filter(([member]) => DOCX_REF_REPLICA_LOADING_MUTATIONS.has(member)))(
  '%s throws DocxReplicaNotReadyError while loading and applies after readiness',
  async (member, args, _loading, check) => {
    const { api, session, replica, release, opens, fallbackReasons, events } = await pendingReplica('editing');
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const version = session.version();
    let caught: unknown;
    try { Reflect.apply(api[member], api, args); } catch (error) { caught = error; }
    expect(caught).toBeInstanceOf(DocxReplicaNotReadyError);
    expect((caught as DocxReplicaNotReadyError).member).toBe(member);
    expect((caught as Error).message).toContain(member);
    expect((caught as Error).message).toContain('flushPendingInput()');
    if (member in DOCX_REF_ASYNC_TWINS) {
      const twin = DOCX_REF_ASYNC_TWINS[member as keyof typeof DOCX_REF_ASYNC_TWINS];
      expect((caught as Error).message).toContain(typeof twin === 'string' ? twin : twin.join(' or '));
    } else {
      expect((caught as Error).message).not.toContain(', or use ');
    }
    await act(async () => {});
    expect(session.version()).toBe(version);
    expect(session.storyIds()).toEqual([]);
    expect(replica.started).toBe(false);
    expect(replica.pending).toBe(true);
    expect(opens).toEqual([]);
    expect(events).toEqual([]);
    expect(fallbackReasons).toEqual([]);
    expect(warning).not.toHaveBeenCalled();
    replica.start();
    await act(async () => { release(); await replica.ready; });
    const readyVersion = session.version();
    act(() => { check(Reflect.apply(api[member], api, args)); });
    expect(session.version()).not.toBe(readyVersion);
    expect(replica.pending).toBe(false);
    expect(opens).toEqual([false]);
    expect(fallbackReasons).toEqual([]);
    expect(warning).not.toHaveBeenCalled();
  }
);

test.each(SYNC_REPLICA_CALLS.filter(([member]) => DOCX_REF_REPLICA_LOADING_MUTATIONS.has(member)))(
  '%s throws DocxReplicaNotReadyError during the preview-to-full handoff and applies after readiness',
  async (member, args, _loading, check) => {
    const { api, session, replica, release, opens, fallbackReasons, events, openingRef } = await pendingReplica('editing');
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const version = session.version();
    openingRef.current = true;
    expect(api.getDocument()).toBeNull();
    expect(api.getEditorRef()).toBeNull();
    expect(api.findInDocument('map')).toEqual([]);
    expect(api.scrollToParaId('00000001')).toBe(false);
    const call = () => Reflect.apply(api[member], api, args);
    let caught: unknown;
    try { call(); } catch (error) { caught = error; }
    expect(caught).toBeInstanceOf(DocxReplicaNotReadyError);
    expect((caught as DocxReplicaNotReadyError).member).toBe(member);
    expect((caught as Error).message).toContain(member);
    expect((caught as Error).message).toContain('await flushPendingInput()');
    if (member in DOCX_REF_ASYNC_TWINS) {
      const twin = DOCX_REF_ASYNC_TWINS[member as keyof typeof DOCX_REF_ASYNC_TWINS];
      expect((caught as Error).message).toContain(typeof twin === 'string' ? twin : twin.join(' or '));
    } else {
      expect((caught as Error).message).not.toContain(', or use ');
    }
    await act(async () => {});
    expect(session.version()).toBe(version);
    expect(session.storyIds()).toEqual([]);
    expect(replica.started).toBe(false);
    expect(replica.pending).toBe(true);
    expect(opens).toEqual([]);
    expect(events).toEqual([]);
    expect(fallbackReasons).toEqual([]);
    expect(warning).not.toHaveBeenCalled();
    replica.start();
    await act(async () => {});
    expect(call).toThrow(DocxReplicaNotReadyError);
    expect(session.version()).toBe(version);
    expect(session.storyIds()).toEqual([]);
    expect(opens).toEqual([]);
    expect(events).toEqual([]);
    await act(async () => {
      release();
      await replica.ready;
      openingRef.current = false;
    });
    const readyVersion = session.version();
    act(() => { check(call()); });
    expect(session.version()).not.toBe(readyVersion);
    expect(replica.pending).toBe(false);
    expect(opens).toEqual([false]);
    expect(fallbackReasons).toEqual([]);
    expect(warning).not.toHaveBeenCalled();
  }
);

test.each(SYNC_REPLICA_CALLS)(
  '%s preserves replica state during an editor replica load',
  async (method, args, loading, check) => {
    const { api, opens, replica, release, session, fallbackReasons, pagedEditorRef } =
      await pendingReplica('editing');
    prepareSyncRead(pagedEditorRef.current!);
    if (loading === DocxReplicaNotReadyError) expect(() => Reflect.apply(api[method], api, args)).toThrow(loading);
    else expect(Reflect.apply(api[method], api, args)).toEqual(loading);
    await act(async () => {});
    expect(opens).toEqual([]);
    expect(fallbackReasons).toEqual([]);
    expect(replica.started).toBe(false);
    expect(replica.pending).toBe(true);
    replica.start();
    await act(async () => { release(); await replica.ready; });
    session.setSelection(
      { story: 'body', paraId: '00000001', offset: 0 },
      { story: 'body', paraId: '00000001', offset: 8 }
    );
    act(() => { check(Reflect.apply(api[method], api, args)); });
    expect(opens).toEqual([false]);
    expect(fallbackReasons).toEqual([]);
  }
);

test.each(SYNC_REPLICA_CALLS)(
  '%s leaves worker proposal hand-over to the replica owner', async (method, args, loading, check) => {
    const { api, opens, replica, release, session, pagedEditorRef, authority, transport, fallbackReasons } =
      await pendingWorkerProposalReplica('editing');
    prepareSyncRead(pagedEditorRef.current!);
    if (loading === DocxReplicaNotReadyError) expect(() => Reflect.apply(api[method], api, args)).toThrow(loading);
    else expect(Reflect.apply(api[method], api, args)).toEqual(loading);
    expectLoadingMutations(api, pagedEditorRef.current!);
    await act(async () => {});
    expect(opens).toEqual([]);
    expect(fallbackReasons).toEqual([]);
    expect(replica.started).toBe(false);
    expect(replica.pending).toBe(true);
    expect(authority.holdsWorkerState()).toBe(true);
    expect(transport.handOver).not.toHaveBeenCalled();
    const ready = api.flushPendingInput();
    await act(async () => {});
    expect(replica.started).toBe(false);
    expect(transport.handOver).not.toHaveBeenCalled();
    replica.start();
    await act(async () => { release(); await ready; });
    expectReadyMutations(api, session, pagedEditorRef.current!);
    expect(transport.handOver).toHaveBeenCalledTimes(1);
    expect(opens).toEqual([false]);
    expect(replica.pending).toBe(false);
    expect(authority.holdsWorkerState()).toBe(false);
    session.setSelection(
      { story: 'body', paraId: '00000001', offset: 0 },
      { story: 'body', paraId: '00000001', offset: 8 }
    );
    act(() => { check(Reflect.apply(api[method], api, args)); });
    expect(opens).toEqual([false]);
    expect(fallbackReasons).toEqual([]);
  }
);

test.each([
  ['an editor replica load', () => pendingReplica('editing')],
  ['worker proposals held in the worker', () => pendingWorkerProposalReplica('editing')],
] as const)('missing comment and change ids navigate to nothing after %s', async (setup, createReplica) => {
  const { api, opens, replica, release, fallbackReasons } = await createReplica();
  expect(api.scrollToCommentId(-1)).toBe(false);
  expect(api.scrollToChangeId(-1)).toBe(false);
  const ready = setup === 'worker proposals held in the worker' ? api.flushPendingInput() : replica.ready;
  await act(async () => {});
  replica.start();
  await act(async () => { release(); await ready; });
  expect(api.scrollToCommentId(-1)).toBe(false);
  expect(api.scrollToChangeId(-1)).toBe(false);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
  expect(replica.pending).toBe(false);
});
