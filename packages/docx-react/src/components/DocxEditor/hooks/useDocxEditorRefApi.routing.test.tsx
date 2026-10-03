import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import { createStyleResolver } from '@betteroffice/docx/styles';
import type { Document } from '@betteroffice/docx/types/document';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { ResidentDocumentRead, ResidentEngineWorkerClient, ResidentProposalReply, YrsSession } from '@betteroffice/docx/yrs';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxDocumentChange, DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import { resetDeprecatedViewerMembersForTests } from '../internals/deprecatedViewerMembers';
import { markPresented, stampWorkerFrameVersion } from '../internals/layoutProvenance';
import { navigateViewer, readViewerSelectionInfo, type ViewerNavigationTarget, type ViewerRefReadAccess } from '../internals/viewerRefReads';
import { deferWorkerOpenReplica } from '../internals/workerOpenReplica';
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import { usePagedEditorRefApi } from './usePagedEditorRefApi';
import { DocxAsyncOnlyError, DocxReplicaNotReadyError, routeViewerRefAccess, useDocxEditorRefApi } from './useDocxEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
beforeEach(resetDeprecatedViewerMembersForTests);
afterEach(() => { cleanup(); mock.restore(); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const INFO = { paraId: 'p', selectedText: 'hello', paragraphText: 'hello', before: '', after: '' };
const MATCHES = [{ paraId: 'p', match: 'hello', before: '', after: '' }];

function apiFor(viewer = false, pendingReplica = false) {
  const events: string[] = [];
  const document = {} as Document;
  const session = {
    version: () => 'v',
    storyIds: () => ['body'],
    paragraphs: () => [{ paraId: 'p', text: 'hello', properties: {} }],
    locateParagraph: () => ({ start: 0, end: 5 }),
    selection: () => ({ anchor: { story: 'body', paraId: 'p', offset: 0 }, head: { story: 'body', paraId: 'p', offset: 5 } }),
    selectionText: () => { events.push('selection'); return INFO; },
    commentTextTarget: () => ({ ok: true }),
    mirrorWorkerDocument: () => {},
    getProposals: () => ({ version: 'v', previewVersion: 0, proposals: [] }),
  } as unknown as YrsSession;
  const hydrate = mock(async () => () => {});
  const fallback = mock(() => {});
  const request = mock(() => {});
  const replica = pendingReplica ? deferWorkerOpenReplica(session, hydrate, fallback, () => {}, { active: () => true, request }) : null;
  const state = { viewer };
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
    getLayout: () => null,
  };
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const subscribers = new Set<(change: DocxDocumentChange) => void>();
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      ref, document, documentFromYrs: () => document, historyStateRef: { current: document }, pagedEditorRef,
      experimentalWorkerOpen: pendingReplica,
      handleSave: async () => null, zoom: 1, setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {}, loadBuffer: async () => {},
      comments: [{ id: 1 } as never], setComments: () => {}, setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() }, documentChangeSubscribersRef: { current: subscribers },
      selectionChangeSubscribersRef: { current: new Set() }, getCachedStyleResolver: createStyleResolver,
      commentIdAllocator: createCommentIdAllocator(), commands: UNAVAILABLE_DOCX_COMMANDS,
      modeRef: { current: 'editing' }, allowHostProposalsRef: { current: true },
      hostSearch: {
        search: async () => ({ query: '', options: { caseSensitive: false }, total: 0, current: -1 }),
        searchNext: () => null, searchPrevious: () => null, searchGoTo: () => null, clearSearch: () => {},
        getSearchState: () => null, onSearchChange: () => () => {},
      },
    });
    return ref;
  });
  return { api: hook.result.current.current!, editor, session, state, replica, hydrate, fallback, request, events, pagedEditorRef, subscribers };
}

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
      expect(caught).toBeInstanceOf(DocxReplicaNotReadyError);
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
  'getPositionAtPoint', 'getSelectionInfo', 'getEditorRef', 'addComment', 'replyToComment',
  'proposeChange', 'setParagraphStyle', 'applyFormatting', 'insertBreak', 'highlightRange',
  'getComments', 'resolveComment', 'readParagraphs', 'onDocumentChange',
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

test('editor twins flush before reading or navigating and preserve synchronous results', async () => {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const host = apiFor();
  expect(host.api.getDocument()).toBe(host.pagedEditorRef.current!.getDocument());
  expect(host.api.getPageContent(1)).toBeNull();
  expect(await host.api.readSelectionInfo()).toEqual(host.api.getSelectionInfo());
  expect(host.events.slice(0, 2)).toEqual(['flush', 'selection']);
  host.events.length = 0;
  expect(await host.api.readPositionAtPoint(1, 2)).toBe(host.api.getPositionAtPoint(1, 2));
  expect(host.events.slice(0, 2)).toEqual(['flush', 'point']);
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
  expect(await host.api.insertCommentReply(1, 'Reply', 'Author')).toBeNumber();
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
      runLayoutPipeline: () => {}, getLayoutRequest: () => null,
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
  expect(bump).toHaveBeenCalledTimes(2);
});
