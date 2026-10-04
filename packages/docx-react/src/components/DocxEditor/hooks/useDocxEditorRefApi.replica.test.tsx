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
import { deferWorkerOpenReplica, workerOpenReplicaOnDemand } from '../internals/workerOpenReplica';
import { beginWorkerProposalHandover, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { EditorMode } from '../internals/editing-modes';
import { DOCX_REF_REPLICA_ACCESS, DOCX_REF_REPLICA_LOADING_ANSWERS, useDocxEditorRefApi } from './useDocxEditorRefApi';

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
  const fallbackReasons: string[] = [];
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
    () => {
      fallbackReasons.push('failure');
      opens.push(true);
      session.openDocx(bytes, true);
    },
    () => { readiness.current = true; },
    { active: () => hydrateOnDemand, request: () => replica.start() }
  );
  const mounted = apiFor(session, document, mode, mountInput ? readiness : undefined);
  return { ...mounted, session, worker, replica, release, opens, fallbackReasons };
}

async function pendingWorkerProposalReplica(mode: EditorMode = 'viewing', hydrateOnDemand = true) {
  const pending = await pendingReplica(mode, false, hydrateOnDemand);
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
  const direct = new Set([
    'focus', 'scrollToPosition', 'openPrintPreview', 'print', 'highlightRange', 'getPositionAtPoint',
  ]);
  expect(Object.keys(DOCX_REF_REPLICA_LOADING_ANSWERS).sort()).toEqual(
    Object.entries(DOCX_REF_REPLICA_ACCESS)
      .filter(([member, access]) => access === 'sync' && !direct.has(member))
      .map(([member]) => member).sort()
  );
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
  await act(async () => {});
  expect(replica.started).toBe(false);
  expect(completed.value).toBe(false);
  expect(events).toEqual([]);
  expect(opens).toEqual([]);
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

test('synchronous reads return loading answers until the owner opens the editor replica', async () => {
  const { api, session, replica, release, opens, fallbackReasons } = await pendingReplica('editing');
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(api.findInDocument('Page')).toEqual([]);
  expect(api.scrollToParaId('00000001')).toBe(false);
  await act(async () => {});
  expect(opens).toEqual([]);
  expect(replica.started).toBe(false);
  expect(replica.pending).toBe(true);
  replica.start();
  await act(async () => { release(); await replica.ready; });
  expect(api.getDocument()).not.toBeNull();
  expect(api.getEditorRef()?.getYrsSession()).toBe(session);
  const first = session.paragraphs('body')[0]!;
  expect(api.findInDocument(first.text)).toContainEqual({
    paraId: first.paraId, match: first.text, before: '', after: '',
  });
  expect(api.scrollToParaId(first.paraId)).toBe(true);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
});

test('editor document access and paragraph styling return loading answers until the peer is ready', async () => {
  const { api, session, replica, release, opens, fallbackReasons, pagedEditorRef } = await pendingReplica('editing');
  const getDocument = spyOn(pagedEditorRef.current!, 'getDocument');
  const paragraphs = spyOn(session, 'paragraphs');
  const open = spyOn(session, 'openDocx');
  const style = spyOn(session, 'applyParagraphStyle');
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const options = { paraId: '00000001', styleId: 'Normal' };
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(api.setParagraphStyle(options)).toBe(false);
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

test('a synchronous write during an in-flight editor handoff returns false until readiness', async () => {
  const { api, session, replica, release, opens, fallbackReasons } = await pendingReplica('editing');
  replica.start();
  await act(async () => {});
  act(() => expect(api.insertBreak({ paraId: '00000001', type: 'page' })).toBe(false));
  expect(opens).toEqual([]);
  expect(session.storyIds()).toEqual([]);
  await act(async () => { release(); await replica.ready; });
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

test('getPositionAtPoint answers without starting an on-demand replica', async () => {
  const { api, opens, replica, release } = await pendingReplica('viewing', false, true);
  let requests = 0;
  replica.onDemand!.request = () => { requests += 1; replica.start(); };
  expect(api.getPositionAtPoint(0, 0)).toBeNull();
  expect(requests).toBe(0);
  expect(replica.started).toBe(false);
  await act(async () => { release(); });
  expect(opens).toEqual([]);
  expect(replica.pending).toBe(true);
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
  }], null, (result: unknown) => expect(result).toEqual(expect.any(Number))],
  ['proposeChange', [{
    paraId: '00000001', search: 'map', replaceWith: 'plan', author: 'Agent',
  }], false, (result: unknown) => expect(result).toBe(true)],
  ['applyFormatting', [{
    paraId: '00000001', search: 'map', marks: { bold: true },
  }], false, (result: unknown) => expect(result).toBe(true)],
  ['setParagraphStyle', [{
    paraId: '00000001', styleId: 'Normal',
  }], false, (result: unknown) => expect(result).toBe(true)],
  ['insertBreak', [{
    paraId: '00000001', type: 'page',
  }], false, (result: unknown) => expect(result).toBe(true)],
] as const;

function prepareSyncRead(editor: PagedEditorRef) {
  editor.getLayout = () => ({
    pages: [{ fragments: [{ kind: 'paragraph', pmStart: 0 }] }],
  }) as unknown as ReturnType<PagedEditorRef['getLayout']>;
  editor.displayPositionToYrsLoc = () => ({ story: 'body', paraId: '00000001', offset: 0 });
  editor.scrollToCommentId = () => true;
  editor.scrollToChangeId = () => true;
}

for (const onDemand of [false, true]) {
  test.each(SYNC_REPLICA_CALLS)(
    `%s returns its loading answer without opening an ${onDemand ? 'on-demand' : 'editor'} replica`,
    async (method, args, loading, check) => {
      const { api, opens, replica, release, session, fallbackReasons, pagedEditorRef } =
        await pendingReplica('editing', false, onDemand);
      prepareSyncRead(pagedEditorRef.current!);
      expect(Reflect.apply(api[method], api, args)).toEqual(loading);
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
}

test.each(SYNC_REPLICA_CALLS)(
  '%s returns its loading answer without requesting worker proposal hand-over', async (method, args, loading, check) => {
    const { api, opens, replica, release, session, pagedEditorRef, authority, transport, fallbackReasons } =
      await pendingWorkerProposalReplica('editing', false);
    prepareSyncRead(pagedEditorRef.current!);
    expect(Reflect.apply(api[method], api, args)).toEqual(loading);
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

test('selection and point reads return null while proposals are held in the worker', async () => {
  const { api, opens, replica, release, transport } = await pendingWorkerProposalReplica();
  let requests = 0;
  replica.onDemand!.request = () => { requests += 1; replica.start(); };
  expect(api.getSelectionInfo()).toBeNull();
  expect(requests).toBe(0);
  expect(replica.started).toBe(false);
  expect(api.getPositionAtPoint(0, 0)).toBeNull();
  expect(requests).toBe(0);
  expect(replica.started).toBe(false);
  await act(async () => { release(); });
  expect(transport.handOver).not.toHaveBeenCalled();
  expect(opens).toEqual([]);
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

test('getEditorRef returns null until readiness and then immediately inserts text', async () => {
  const { api, session, replica, release, opens, fallbackReasons } = await pendingReplica('editing', true);
  expect(replica.pending).toBe(true);
  expect(api.getEditorRef()).toBeNull();
  expect(replica.started).toBe(false);
  expect(opens).toEqual([]);
  replica.start();
  await act(async () => { release(); await replica.ready; });
  await act(async () => {
    api.getEditorRef()!.insertText('Immediate ');
    await api.flushPendingInput();
  });
  expect(replica.pending).toBe(false);
  expect(opens).toEqual([false]);
  expect(fallbackReasons).toEqual([]);
  expect(session.paragraphs('body')[0]!.text).toStartWith('Immediate ');
});
