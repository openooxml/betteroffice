import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
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
import { deferWorkerOpenReplica, workerOpenReplicaOnDemand, type WorkerOpenFallbackReason } from '../internals/workerOpenReplica';
import { beginWorkerProposalHandover, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { EditorMode } from '../internals/editing-modes';
import { DOCX_REF_REPLICA_ACCESS, DocxReplicaNotReadyError, useDocxEditorRefApi } from './useDocxEditorRefApi';

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
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function apiFor(
  session: YrsSession,
  document: Document,
  mode: EditorMode = 'viewing',
  replicaReadyRef?: { current: boolean }
) {
  const events: string[] = [];
  const inputRef = { current: null as YrsInputRef | null };
  const project = () => {
    const base = session.materializeDocx();
    return base ? yrsToDocument(session, base) : null;
  };
  const editor = {
    getYrsSession: () => session,
    getDocument: project,
    flushPendingInput: async () => {
      events.push('flush');
      await inputRef.current?.flushPendingInput();
    },
    insertText: (text: string) => inputRef.current?.insertText(text),
    syncYrsInputState: () => { events.push('sync'); return true; },
    getLayout: () => null,
    getLayoutRequest: () => null,
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
      ref,
      document,
      documentFromYrs: project,
      historyStateRef: { current: document },
      pagedEditorRef,
      handleSave: async () => {
        events.push('save');
        return new TextEncoder().encode(session.paragraphs('body').map((paragraph) => paragraph.text).join('\n')).buffer;
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
  return { api, events, pagedEditorRef };
}

async function pendingReplica(mode: EditorMode = 'viewing', mountInput = false, hydrateOnDemand = false) {
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
    { active: () => hydrateOnDemand, request: () => replica.start() }
  );
  const mounted = apiFor(session, document, mode, mountInput ? readiness : undefined);
  return { ...mounted, session, worker, replica, release, opens, fallbackReasons };
}

async function pendingWorkerProposalReplica() {
  const pending = await pendingReplica('viewing', false, true);
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
    'proposeChanges', 'readParagraphs', 'replyToComment', 'resolveComment', 'resolveParagraphAnchors', 'save',
    'search', 'searchGoTo', 'searchNext', 'searchPrevious',
    'scrollToChangeId', 'scrollToCommentId', 'scrollToPage', 'scrollToParaId', 'scrollToPosition',
    'setParagraphStyle', 'setProposalStates', 'setZoom', 'validateEdits', 'whenLayoutComplete',
    'withdrawProposals',
  ].sort());
});

test('async reads, save, exports and write refusals wait for the main replica', async () => {
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
  expect(events).toEqual([]);
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
  expect(new TextDecoder().decode(values[10]!)).toBe(session.paragraphs('body').map((paragraph) => paragraph.text).join('\n'));
  expect(values[12]).toMatchObject({ ok: false, version: session.version(), failure: { code: 'layout-unavailable' } });
  expect(values[13]).toBe(0);
});

test('synchronous reads finish the main open without changing their return types', async () => {
  const { api, session, replica, opens, fallbackReasons } = await pendingReplica();
  let document: Document | null = null;
  act(() => { document = api.getDocument(); });
  expect(document).not.toBeNull();
  expect(opens).toEqual([true]);
  expect(replica.pending).toBe(false);
  expect(api.getEditorRef()?.getYrsSession()).toBe(session);
  const first = session.paragraphs('body')[0]!;
  expect(api.findInDocument(first.text)).toContainEqual({
    paraId: first.paraId, match: first.text, before: '', after: '',
  });
  expect(api.scrollToParaId(first.paraId)).toBe(true);
  expect(fallbackReasons).toEqual([{ syncAccess: 'getDocument' }]);
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

test('a synchronous write during an in-flight handoff is not overwritten by its reply', async () => {
  const { api, session, replica, release, opens } = await pendingReplica();
  replica.start();
  await act(async () => {});
  act(() => expect(api.insertBreak({ paraId: '00000001', type: 'page' })).toBe(true));
  const edited = session.encodeState();
  await act(async () => { release(); await replica.ready; });
  expect(opens).toEqual([true]);
  expect(session.encodeState()).toEqual(edited);
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

test('layout completion leaves an on-demand replica pending while async reads start it', async () => {
  const { api, opens, replica, session, release } = await pendingReplica('viewing', false, true);
  expect(await api.whenLayoutComplete({ timeoutMs: 20 })).toBe(0);
  expect(workerOpenReplicaOnDemand(session)).toBe(true);
  expect(opens).toEqual([]);
  const read = api.readParagraphs({ view: 'accepted' });
  await act(async () => {
    release();
    expect(await read).toMatchObject({ ok: true });
  });
  expect(opens).toEqual([false]);
  expect(replica.pending).toBe(false);
});

test.each(['focus', 'scrollToPosition', 'print', 'openPrintPreview', 'highlightRange'] as const)(
  '%s leaves an on-demand replica pending', async (method) => {
    const { api, opens, replica, release } = await pendingReplica('viewing', false, true);
    act(() => {
      if (method === 'scrollToPosition') api.scrollToPosition(0);
      else if (method === 'highlightRange') api.highlightRange(0, 1);
      else api[method]();
    });
    expect(opens).toEqual([]);
    expect(replica.pending).toBe(true);
    await act(async () => { release(); });
    expect(opens).toEqual([]);
    expect(replica.pending).toBe(true);
  }
);

test('getSelectionInfo returns null without starting an on-demand replica', async () => {
  const { api, opens, replica, release } = await pendingReplica('viewing', false, true);
  expect(api.getSelectionInfo()).toBeNull();
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
  await act(async () => { release(); });
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
});

test('getPositionAtPoint requests an on-demand replica and returns null', async () => {
  const { api, opens, replica, release } = await pendingReplica('viewing', false, true);
  let requests = 0;
  replica.onDemand!.request = () => { requests += 1; replica.start(); };
  expect(api.getPositionAtPoint(0, 0)).toBeNull();
  expect(requests).toBe(1);
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
  await act(async () => { release(); await replica.ready; });
  expect(opens).toEqual([false]);
  expect(replica.pending).toBe(false);
});

test.each([
  ['getDocument', (api: DocxEditorRef) => expect(api.getDocument()).not.toBeNull()],
  ['getEditorRef', (api: DocxEditorRef) => expect(api.getEditorRef()).not.toBeNull()],
  ['scrollToParaId', (api: DocxEditorRef) => expect(api.scrollToParaId('00000001')).toBe(true)],
  ['scrollToCommentId', (api: DocxEditorRef) => expect(api.scrollToCommentId(-1)).toBe(false)],
  ['scrollToChangeId', (api: DocxEditorRef) => expect(api.scrollToChangeId(-1)).toBe(false)],
  ['findInDocument', (api: DocxEditorRef) => expect(api.findInDocument('map')).toContainEqual({
    paraId: '00000001', match: 'map', before: 'Page ', after: '',
  })],
  ['getPageContent', (api: DocxEditorRef) => expect(api.getPageContent(1)).toEqual({
    pageNumber: 1,
    text: '[00000001] Page map',
    paragraphs: [{ paraId: '00000001', text: 'Page map', styleId: 'Heading1' }],
  })],
  ['addComment', (api: DocxEditorRef) => expect(api.addComment({
    paraId: '00000001', search: 'map', text: 'Check', author: 'Ann',
  })).toEqual(expect.any(Number))],
  ['proposeChange', (api: DocxEditorRef) => expect(api.proposeChange({
    paraId: '00000001', search: 'map', replaceWith: 'plan', author: 'Agent',
  })).toBe(true)],
  ['applyFormatting', (api: DocxEditorRef) => expect(api.applyFormatting({
    paraId: '00000001', search: 'map', marks: { bold: true },
  })).toBe(true)],
  ['setParagraphStyle', (api: DocxEditorRef) => expect(api.setParagraphStyle({
    paraId: '00000001', styleId: 'Normal',
  })).toBe(true)],
  ['insertBreak', (api: DocxEditorRef) => expect(api.insertBreak({
    paraId: '00000001', type: 'page',
  })).toBe(true)],
] as const)('%s synchronously opens an on-demand replica and returns its result', async (method, check) => {
  const mode = ['addComment', 'proposeChange', 'applyFormatting', 'setParagraphStyle', 'insertBreak']
    .includes(method) ? 'editing' : 'viewing';
  const { api, opens, replica, pagedEditorRef } = await pendingReplica(mode, false, true);
  if (method === 'getPageContent') {
    pagedEditorRef.current!.getLayout = () => ({
      pages: [{ fragments: [{ kind: 'paragraph', pmStart: 0 }] }],
    }) as unknown as ReturnType<PagedEditorRef['getLayout']>;
    pagedEditorRef.current!.displayPositionToYrsLoc = () => ({
      story: 'body', paraId: '00000001', offset: 0,
    });
  }
  expect(replica.pending).toBe(true);
  act(() => { check(api); });
  expect(opens).toEqual([true]);
  expect(replica.pending).toBe(false);
});

test.each([
  ['getDocument', (api: DocxEditorRef) => expect(api.getDocument()).not.toBeNull()],
  ['getEditorRef', (api: DocxEditorRef) => expect(api.getEditorRef()).not.toBeNull()],
  ['scrollToParaId', (api: DocxEditorRef) => expect(api.scrollToParaId('00000001')).toBe(true)],
  ['scrollToCommentId', (api: DocxEditorRef) => expect(api.scrollToCommentId(-1)).toBe(false)],
  ['scrollToChangeId', (api: DocxEditorRef) => expect(api.scrollToChangeId(-1)).toBe(false)],
  ['findInDocument', (api: DocxEditorRef) => expect(api.findInDocument('map')).toContainEqual({
    paraId: '00000001', match: 'map', before: 'Page ', after: '',
  })],
  ['getPageContent', (api: DocxEditorRef) => expect(api.getPageContent(1)).toEqual({
    pageNumber: 1,
    text: '[00000001] Page map',
    paragraphs: [{ paraId: '00000001', text: 'Page map', styleId: 'Heading1' }],
  })],
  ['addComment', (api: DocxEditorRef) => expect(api.addComment({
    paraId: '00000001', search: 'map', text: 'Check', author: 'Ann',
  })).toEqual(expect.any(Number))],
  ['proposeChange', (api: DocxEditorRef) => expect(api.proposeChange({
    paraId: '00000001', search: 'map', replaceWith: 'plan', author: 'Agent',
  })).toBe(true)],
  ['applyFormatting', (api: DocxEditorRef) => expect(api.applyFormatting({
    paraId: '00000001', search: 'map', marks: { bold: true },
  })).toBe(true)],
  ['setParagraphStyle', (api: DocxEditorRef) => expect(api.setParagraphStyle({
    paraId: '00000001', styleId: 'Normal',
  })).toBe(true)],
  ['insertBreak', (api: DocxEditorRef) => expect(api.insertBreak({
    paraId: '00000001', type: 'page',
  })).toBe(true)],
] as const)('%s requests hand-over and throws while proposals are held in the worker', async (method, check) => {
  const { api, opens, replica, release, pagedEditorRef, authority, transport } =
    await pendingWorkerProposalReplica();
  if (method === 'getPageContent') {
    pagedEditorRef.current!.getLayout = () => ({
      pages: [{ fragments: [{ kind: 'paragraph', pmStart: 0 }] }],
    }) as unknown as ReturnType<PagedEditorRef['getLayout']>;
    pagedEditorRef.current!.displayPositionToYrsLoc = () => ({
      story: 'body', paraId: '00000001', offset: 0,
    });
  }
  let error: unknown;
  act(() => {
    try { check(api); } catch (failure) { error = failure; }
  });
  expect(error).toBeInstanceOf(DocxReplicaNotReadyError);
  expect((error as DocxReplicaNotReadyError).member).toBe(method);
  expect((error as Error).message).toContain(method);
  expect((error as Error).message).toContain('flushPendingInput()');
  expect(opens).toEqual([]);
  expect(replica.started).toBe(true);
  expect(replica.pending).toBe(true);
  expect(authority.holdsWorkerState()).toBe(true);
  const ready = api.flushPendingInput();
  await act(async () => { release(); await ready; });
  expect(transport.handOver).toHaveBeenCalledTimes(1);
  expect(opens).toEqual([false]);
  expect(replica.pending).toBe(false);
  expect(authority.holdsWorkerState()).toBe(false);
  act(() => { check(api); });
  expect(opens).toEqual([false]);
});

test('selection and point reads return null while proposals are held in the worker', async () => {
  const { api, opens, replica, release, transport } = await pendingWorkerProposalReplica();
  let requests = 0;
  replica.onDemand!.request = () => { requests += 1; replica.start(); };
  expect(api.getSelectionInfo()).toBeNull();
  expect(requests).toBe(0);
  expect(replica.started).toBe(false);
  expect(api.getPositionAtPoint(0, 0)).toBeNull();
  expect(requests).toBe(1);
  expect(replica.started).toBe(true);
  expect(replica.pending).toBe(true);
  expect(opens).toEqual([]);
  await act(async () => { release(); await replica.ready; });
  expect(transport.handOver).toHaveBeenCalledTimes(1);
  expect(opens).toEqual([false]);
  expect(replica.pending).toBe(false);
});

test('focus stays direct while proposals are held in the worker', async () => {
  const { api, opens, replica, release, pagedEditorRef, transport } = await pendingWorkerProposalReplica();
  const focus = mock(() => {});
  pagedEditorRef.current!.focus = focus;
  act(() => { expect(() => api.focus()).not.toThrow(); });
  expect(focus).toHaveBeenCalledTimes(1);
  expect(opens).toEqual([]);
  expect(replica.started).toBe(false);
  await act(async () => { release(); });
  expect(transport.handOver).not.toHaveBeenCalled();
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
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

test('getEditorRef immediately inserts text after synchronously finishing the replica', async () => {
  const { api, session, replica, opens, fallbackReasons } = await pendingReplica('editing', true);
  expect(replica.pending).toBe(true);
  await act(async () => {
    api.getEditorRef()!.insertText('Immediate ');
    await api.flushPendingInput();
  });
  expect(replica.pending).toBe(false);
  expect(opens).toEqual([true]);
  expect(fallbackReasons).toEqual([{ syncAccess: 'getEditorRef' }]);
  expect(session.paragraphs('body')[0]!.text).toStartWith('Immediate ');
});
