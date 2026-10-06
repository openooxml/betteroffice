import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type ResidentProposalReply, type YrsSession, type YrsStickyPosition } from '@betteroffice/docx/yrs';
import { readResidentSearch } from '@betteroffice/docx/yrs/residentSearch';
import type { ResidentDocumentRead } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import type { WorkerOpenedDocument } from './useDisplayList';
import type { PagedEditorRef } from '../PagedEditor';
import { createYrsPositionProjection } from '../internals/yrsPositionProjection';
import { stampSourceVersion } from '../internals/layoutProvenance';
import { beginWorkerProposalHandover, registerWorkerProposalAuthority, workerProposalAuthority } from '../internals/workerProposalAuthority';
import { deferWorkerOpenReplica, requestWorkerOpenReplica } from '../internals/workerOpenReplica';
import { yrsCellStory } from '../yrsCommands';
import { topPageInView, useHostSearch, type DocxSearchState } from './useHostSearch';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

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

/**
 * Body "The cat and the dog", a 1x2 table "the | THE", then "the end"; a header with "the".
 * `repeat` has the second page paint the first page's positions again, as a repeated table
 * header does.
 */
async function mount(page = 1, repeat = false) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId: first } = session.createStory('body', 'The cat and the dog');
  const { secondParaId: last } = session.splitParagraph({ story: 'body', paraId: first, offset: 19 });
  const { table } = session.insertTable({ story: 'body', paraId: last, offset: 0 }, 1, 2);
  for (const [column, text] of [
    [0, 'the'],
    [1, 'THE'],
  ] as const) {
    const story = yrsCellStory(session, { ...table, row: 0, column })!;
    session.insertText({ story, paraId: session.paragraphs(story)[0].paraId, offset: 0 }, text);
  }
  session.insertText({ story: 'body', paraId: last, offset: 1 }, 'the end');
  session.createStory('hdr1', 'the header');

  const reveals: number[] = [];
  const signals: AbortSignal[] = [];
  let placeable = () => true;
  const editor = {
    getYrsSession: () => session,
    hasPendingInput: () => false,
    flushPendingInput: async () => {},
    yrsLocToDisplayPosition: (loc: Parameters<PagedEditorRef['yrsLocToDisplayPosition']>[0]) =>
      createYrsPositionProjection(session, 'body')!.positionForLoc(loc),
    revealDisplayPosition: (position: number, signal?: AbortSignal) => {
      reveals.push(position);
      if (signal) signals.push(signal);
      return placeable() ? ('scrolled' as const) : ('unsupported' as const);
    },
  } as unknown as PagedEditorRef;
  // two pages: the first paragraph, then the table and the rest
  const firstCell = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  const splitAt = editor.yrsLocToDisplayPosition({
    story: firstCell,
    paraId: session.paragraphs(firstCell)[0].paraId,
    offset: 0,
  })!;
  // one display list per version, as the renderer publishes them
  const lists = new Map<number, DisplayListQueries>();
  const queries = (version: number) => {
    if (!lists.has(version)) {
      lists.set(version, {
        displayList: {
          pages: [
            { pageIndex: 0, primitives: [{ kind: 'text', docStart: 0, docEnd: splitAt - 1 }] },
            {
              pageIndex: 1,
              primitives: [
                ...(repeat ? [{ kind: 'text', docStart: 0, docEnd: splitAt - 3 }] : []),
                { kind: 'text', docStart: splitAt, docEnd: 100000 },
              ],
            },
          ],
        },
        anchorRect: (position: number) => ({
          pageIndex: position < splitAt ? 0 : 1,
          x: 0,
          y: 0,
          width: 1,
          height: 1,
        }),
      } as unknown as DisplayListQueries);
    }
    return lists.get(version)!;
  };
  const pagedEditorRef: { current: PagedEditorRef | null } = { current: editor };
  // pages 1 and 2 painted, `page` at the top of the window
  const host = document.createElement('div');
  for (const index of [0, 1]) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    const top = (index + 1 - page) * 1000 + 10;
    canvas.getBoundingClientRect = () => new DOMRect(0, top, 800, 990);
    host.append(canvas);
  }
  const hook = renderHook(
    ({ version }) =>
      useHostSearch({
        pagedEditorRef,
        displayListQueries: queries(version),
        canvasHostRef: { current: host },
      }),
    { initialProps: { version: 0 } }
  );
  const events: Array<DocxSearchState | null> = [];
  hook.result.current.api.onSearchChange((state) => events.push(state));
  return {
    session,
    first,
    last,
    hook,
    reveals,
    signals,
    events,
    pagedEditorRef,
    host,
    queries,
    layOut: (canPlace: () => boolean) => {
      placeable = canPlace;
    },
    stamp: (version: number, sourceVersion: string) => stampSourceVersion(queries(version), sourceVersion),
  };
}

async function mountText(text: string) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', text);
  const reveals: number[] = [];
  let placeable = () => true;
  const pagedEditorRef = { current: {
    getYrsSession: () => session,
    hasPendingInput: () => false,
    yrsLocToDisplayPosition: (loc: Parameters<PagedEditorRef['yrsLocToDisplayPosition']>[0]) =>
      createYrsPositionProjection(session, 'body')!.positionForLoc(loc),
    revealDisplayPosition: (position: number) => {
      reveals.push(position);
      return placeable() ? 'scrolled' : 'unsupported';
    },
  } as unknown as PagedEditorRef };
  const hook = renderHook(() => useHostSearch({
    pagedEditorRef,
    displayListQueries: null,
    canvasHostRef: { current: null },
  }));
  return {
    session, paraId, reveals, pagedEditorRef, hook,
    layOut: (canPlace: () => boolean) => { placeable = canPlace; },
  };
}

async function workerSearch(h: Pick<Awaited<ReturnType<typeof mount>>, 'session' | 'pagedEditorRef'>) {
  const session = await createYrsSession();
  sessions.push(session);
  session.loadState(h.session.encodeState());
  const snapshot = (): ResidentProposalReply => ({
    mirror: { version: session.version(), proposals: { previewVersion: 0, entries: [] } },
    geometry: { version: session.version(), previewVersion: 0, proposals: '[]', targets: {}, hidden: [] },
    changedStories: [], updates: [], stateVector: new Uint8Array(),
  });
  let blocked: Promise<void> | null = null;
  let blockedAnchor: Promise<void> | null = null;
  const reads: Extract<ResidentDocumentRead, { kind: 'searchText' }>[] = [];
  const anchorReads: Extract<ResidentDocumentRead, { kind: 'stickyAnchors' }>[] = [];
  const anchors: Array<YrsStickyPosition | null> = [];
  const worker = {
    proposal: async () => snapshot(),
    documentRead: async (read: ResidentDocumentRead) => {
      if (read.kind === 'stickyAnchors') {
        anchorReads.push(read);
        const waiting = blockedAnchor;
        blockedAnchor = null;
        const value = read.locs.map((loc) => {
          try {
            return session.encodeStickyPosition(loc);
          } catch {
            return null;
          }
        });
        anchors.push(value[0] ?? null);
        const version = session.version();
        await waiting;
        return { version, value };
      }
      if (read.kind !== 'searchText') throw new Error('unexpected read');
      reads.push(read);
      const waiting = blocked;
      blocked = null;
      const value = readResidentSearch(session, read.query, read.caseSensitive, read.carry);
      const version = session.version();
      await waiting;
      return { version, value };
    },
    handOver: async () => ({
      state: session.encodeState(), version: session.version(), proposals: snapshot().mirror.proposals,
    }),
  };
  registerWorkerProposalAuthority(h.session, worker as unknown as WorkerOpenedDocument, {
    relayout: () => {}, current: () => h.pagedEditorRef.current?.getYrsSession() === h.session,
    laidOut: async () => {}, adopted: () => {}, handedOver: () => {}, contentChanged: () => {},
  });
  let replicaRequests = 0;
  deferWorkerOpenReplica(h.session, async () => {
    replicaRequests += 1;
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.session.loadState(handover.state); handover.complete(); };
  }, () => { throw new Error('unexpected main fallback'); }, () => {});
  return {
    session,
    reads,
    anchorReads,
    anchors,
    replicaRequests: () => replicaRequests,
    mirror: () => h.session.mirrorWorkerDocument(snapshot().mirror),
    holdRead: () => {
      let release!: () => void;
      blocked = new Promise<void>((resolve) => { release = resolve; });
      return release;
    },
    holdAnchor: () => {
      let release!: () => void;
      blockedAnchor = new Promise<void>((resolve) => { release = resolve; });
      return release;
    },
  };
}

test('worker search starts in view and navigates cached body and table ranges', async () => {
  const h = await mount(2);
  const worker = await workerSearch(h);
  const mainReads = [
    spyOn(h.session, 'searchText'), spyOn(h.session, 'encodeStickyPosition'),
    spyOn(h.session, 'resolveStickyPosition'), spyOn(h.pagedEditorRef.current!, 'yrsLocToDisplayPosition'),
    spyOn(h.pagedEditorRef.current!, 'flushPendingInput'),
  ];
  try {
    const api = h.hook.result.current.api;
    await act(async () => {
      expect(await api.search('the')).toMatchObject({ total: 5, current: 2 });
    });
    expect(h.hook.result.current.highlight!.matches.map(({ displayFrom, displayTo }) => ({
      displayFrom, displayTo,
    }))).toEqual(readResidentSearch(worker.session, 'the', false).matches.map(({ displayFrom, displayTo }) => ({
      displayFrom, displayTo,
    })));
    await act(async () => {
      expect(api.searchGoTo(4)?.current).toBe(4);
      expect(api.searchNext()?.current).toBe(0);
      expect(api.searchPrevious()?.current).toBe(4);
    });
    expect(worker.reads).toHaveLength(1);
    expect(worker.anchorReads.map(({ locs }) => locs)).toEqual([2, 4, 0, 4].map((index) =>
      [index, (index + 4) % 5, (index + 1) % 5].map((i) => {
        const match = readResidentSearch(worker.session, 'the', false).matches[i];
        return { story: match.story, paraId: match.paraId, offset: match.start };
      })
    ));
    expect(worker.anchorReads.every((read) => read.version === worker.session.version())).toBe(true);
    for (const read of mainReads) expect(read).not.toHaveBeenCalled();
  } finally {
    for (const read of mainReads) read.mockRestore();
  }
  expect(worker.replicaRequests()).toBe(0);
});

test('worker refresh survives ref replacement, coalesces reads and retries stale versions', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  h.layOut(() => false);
  await act(async () => { api.searchGoTo(3); api.searchGoTo(4); });
  const before = h.hook.result.current.highlight!.matches[4].displayFrom;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  const release = worker.holdRead();
  act(() => {
    expect(api.getSearchState()).toMatchObject({ total: 5, current: 4 });
    api.getSearchState();
  });
  h.stamp(1, worker.session.version());
  h.hook.rerender({ version: 1 });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  h.pagedEditorRef.current = { ...h.pagedEditorRef.current! };
  expect(h.hook.result.current.highlight?.matches).toHaveLength(5);
  const published = h.events.length;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  h.stamp(2, worker.session.version());
  h.hook.rerender({ version: 2 });
  expect(worker.reads).toHaveLength(2);
  await act(async () => { release(); });
  await waitFor(() => expect(h.hook.result.current.highlight?.matches).toHaveLength(7));
  expect(api.getSearchState()).toMatchObject({ total: 7, current: 6 });
  expect(h.hook.result.current.highlight!.matches[6].displayFrom).toBe(before + 8);
  expect(worker.reads).toHaveLength(3);
  expect(h.events.slice(published).map((state) => state?.total)).toEqual([7]);
  expect(h.reveals.at(-1)).toBe(h.hook.result.current.highlight!.matches[6].displayFrom);
  act(() => { expect(api.searchPrevious()).toMatchObject({ total: 7, current: 5 }); });
  expect(h.reveals.at(-1)).toBe(h.hook.result.current.highlight!.matches[5].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test('a newer search and clear discard in-flight worker refreshes', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  const release = worker.holdRead();
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  let searched!: Promise<DocxSearchState>;
  act(() => { searched = api.search('dog'); });
  await act(async () => { release(); await searched; });
  expect(api.getSearchState()).toMatchObject({ query: 'dog', total: 1, current: 0 });
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'dog ');
  worker.mirror();
  const clearRelease = worker.holdRead();
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(4));
  act(() => api.clearSearch());
  const events = h.events.length;
  await act(async () => { clearRelease(); });
  expect(api.getSearchState()).toBeNull();
  expect(h.hook.result.current.highlight).toBeNull();
  expect(h.events).toHaveLength(events);
  expect(worker.replicaRequests()).toBe(0);
});

test('initial worker search retries a version change and clearing cancels a pending search', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  const release = worker.holdRead();
  let searched!: Promise<DocxSearchState>;
  act(() => { searched = api.search('the'); });
  await waitFor(() => expect(worker.reads).toHaveLength(1));
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  await act(async () => { release(); await searched; });
  expect(worker.reads).toHaveLength(2);
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 0 });
  const clearRelease = worker.holdRead();
  act(() => { searched = api.search('dog'); });
  await waitFor(() => expect(worker.reads).toHaveLength(3));
  act(() => api.clearSearch());
  await act(async () => { clearRelease(); await searched; });
  expect(api.getSearchState()).toBeNull();
  const replacedRelease = worker.holdRead();
  act(() => { searched = api.search('the'); });
  await waitFor(() => expect(worker.reads).toHaveLength(4));
  h.pagedEditorRef.current = { ...h.pagedEditorRef.current! };
  await act(async () => { replacedRelease(); await searched; });
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 0 });
  expect(h.hook.result.current.highlight!.matches).toEqual(readResidentSearch(worker.session, 'the', false).matches);
  act(() => { expect(api.searchGoTo(5)).toMatchObject({ total: 6, current: 5 }); });
  expect(h.reveals.at(-1)).toBe(h.hook.result.current.highlight!.matches[5].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test('initial worker search refreshes after two edits make both search reads stale', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  const firstRelease = worker.holdRead();
  let searched!: Promise<DocxSearchState>;
  act(() => { searched = api.search('the'); });
  await waitFor(() => expect(worker.reads).toHaveLength(1));
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  const secondRelease = worker.holdRead();
  await act(async () => { firstRelease(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  await act(async () => { secondRelease(); await searched; });
  const expected = readResidentSearch(worker.session, 'the', false).matches;
  await waitFor(() => expect(h.hook.result.current.highlight!.matches).toEqual(expected));
  expect(h.events.at(-1)).toMatchObject({ total: 7 });
  expect(worker.reads).toHaveLength(3);
  expect(worker.replicaRequests()).toBe(0);
});

test('same-index navigation during the first worker publish requests an anchor and carries it', async () => {
  const h = await mount(2);
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  let navigated = false;
  api.onSearchChange((state) => {
    if (!state || navigated) return;
    navigated = true;
    api.searchGoTo(state.current);
  });
  await act(async () => { expect(await api.search('the')).toMatchObject({ current: 2 }); });
  expect(worker.anchorReads).toHaveLength(1);
  const before = h.hook.result.current.highlight!.matches[2].displayFrom;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  h.stamp(1, worker.session.version());
  h.hook.rerender({ version: 1 });
  await waitFor(() => expect(h.hook.result.current.highlight?.matches).toHaveLength(6));
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 3 });
  expect(h.hook.result.current.highlight!.current).toBe(3);
  expect(h.hook.result.current.highlight!.matches[3].displayFrom).toBe(before + 4);
  expect(worker.reads[1].carry).toEqual(worker.anchors[0]);
  expect(worker.reads[1].carry).not.toBeNull();
  expect(worker.replicaRequests()).toBe(0);
});

test('worker refresh waits for the navigated match anchor before carrying it through an edit', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  const release = worker.holdAnchor();
  act(() => { api.searchGoTo(3); });
  await waitFor(() => expect(worker.anchorReads).toHaveLength(2));
  const before = h.hook.result.current.highlight!.matches[3].displayFrom;
  const events = h.events.length;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  h.stamp(1, worker.session.version());
  h.hook.rerender({ version: 1 });
  expect(worker.reads).toHaveLength(1);
  await act(async () => { release(); });
  await waitFor(() => expect(h.hook.result.current.highlight?.matches).toHaveLength(6));
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 4 });
  expect(h.hook.result.current.highlight!.matches[4].displayFrom).toBe(before + 4);
  expect(worker.reads[1].carry).toEqual(worker.anchors[1]);
  expect(worker.reads[1].carry).not.toBeNull();
  expect(h.events.slice(events).map((state) => state?.current)).toEqual([4]);
  expect(worker.replicaRequests()).toBe(0);
});

test('navigation during a held worker refresh encodes the old match at the current state and carries it', async () => {
  const h = await mount(2);
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  const release = worker.holdRead();
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  h.layOut(() => false);
  act(() => { expect(api.searchGoTo(4)?.current).toBe(4); });
  const before = h.hook.result.current.highlight!.matches[4].displayFrom;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  h.stamp(1, worker.session.version());
  h.hook.rerender({ version: 1 });
  h.layOut(() => true);
  await act(async () => { release(); });
  await waitFor(() => expect(h.hook.result.current.highlight?.matches).toHaveLength(7));
  expect(worker.anchors[1]).not.toBeNull();
  expect(worker.reads.at(-1)!.carry).toEqual(worker.anchors[1]);
  expect(worker.reads.at(-1)!.carry).not.toBeNull();
  expect(api.getSearchState()).toMatchObject({ total: 7, current: 6 });
  expect(h.hook.result.current.highlight!.current).toBe(6);
  expect(h.hook.result.current.highlight!.matches[6].displayFrom).toBe(before + 8);
  expect(h.hook.result.current.highlight!.matches).toEqual(readResidentSearch(worker.session, 'the', false).matches);
  expect(h.reveals.at(-1)).toBe(h.hook.result.current.highlight!.matches[6].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test.each(['next', 'goTo', 'previous'] as const)('stale %s navigation carries an exact neighbour through a prefix deletion', async (navigation) => {
  const h = await mountText('the A the B the C');
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { expect(await api.search('the')).toMatchObject({ total: 3, current: 0 }); });
  const index = navigation === 'previous' ? 2 : 1;
  const carry = worker.session.encodeStickyPosition({ story: 'body', paraId: h.paraId, offset: index * 6 });
  expect(worker.anchorReads).toEqual([{
    kind: 'stickyAnchors', version: worker.session.version(),
    locs: [0, 12, 6].map((offset) => ({ story: 'body', paraId: h.paraId, offset })),
  }]);
  const release = worker.holdRead();
  worker.session.deleteRange({
    story: 'body', start: { paraId: h.paraId, offset: 0 }, end: { paraId: h.paraId, offset: 6 },
  });
  worker.mirror();
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  act(() => {
    const state = navigation === 'next' ? api.searchNext() :
      navigation === 'previous' ? api.searchPrevious() : api.searchGoTo(1);
    expect(state).toMatchObject({ total: 3, current: index });
  });
  expect(worker.anchorReads).toHaveLength(1);
  await act(async () => { release(); });
  await waitFor(() => expect(api.getSearchState()).toMatchObject({ total: 2, current: index - 1 }));
  const matches = readResidentSearch(worker.session, 'the', false).matches;
  expect(worker.reads.at(-1)!.carry).toEqual(carry);
  expect(h.hook.result.current.highlight).toEqual({ matches, current: index - 1 });
  expect(h.reveals.at(-1)).toBe(matches[index - 1].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test.each([3, 5])('two next steps during a held refresh use cached anchors when available with %i matches', async (count) => {
  const h = await mountText(['the A', 'the B', 'the C', 'the D', 'the E'].slice(0, count).join(' '));
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { expect(await api.search('the')).toMatchObject({ total: count, current: 0 }); });
  const loc = { story: 'body', paraId: h.paraId, offset: 12 };
  const exact = worker.session.encodeStickyPosition(loc);
  const release = worker.holdRead();
  worker.session.deleteRange({
    story: 'body', start: { paraId: h.paraId, offset: 0 }, end: { paraId: h.paraId, offset: 6 },
  });
  worker.mirror();
  const carry = count === 3 ? exact : worker.session.encodeStickyPosition(loc);
  const current = count === 3 ? 1 : 2;
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  act(() => {
    expect(api.searchNext()?.current).toBe(1);
    expect(api.searchNext()?.current).toBe(2);
  });
  expect(worker.anchorReads).toHaveLength(1);
  await act(async () => { release(); });
  await waitFor(() => expect(api.getSearchState()).toMatchObject({ total: count - 1, current }));
  const matches = readResidentSearch(worker.session, 'the', false).matches;
  expect(worker.reads.at(-1)!.carry).toEqual(carry);
  if (count === 5) {
    expect(carry).not.toEqual(exact);
    expect(worker.anchors[1]).toEqual(carry);
  }
  expect(h.hook.result.current.highlight).toEqual({ matches, current });
  expect(h.reveals.at(-1)).toBe(matches[current].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test('an approximate neighbour prefetch preserves the exact current anchor', async () => {
  const h = await mountText('the A the B the C the D the E');
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  const version = worker.session.version();
  const carry = worker.session.encodeStickyPosition({ story: 'body', paraId: h.paraId, offset: 6 });
  h.layOut(() => false);
  await act(async () => {
    expect(api.searchNext()?.current).toBe(1);
    worker.session.deleteRange({
      story: 'body', start: { paraId: h.paraId, offset: 0 }, end: { paraId: h.paraId, offset: 6 },
    });
    worker.mirror();
    api.getSearchState();
    h.layOut(() => true);
  });
  await waitFor(() => expect(api.getSearchState()).toMatchObject({ total: 4, current: 0 }));
  const matches = readResidentSearch(worker.session, 'the', false).matches;
  expect(worker.anchorReads[1].version).toBe(version);
  expect(worker.anchors[1]).not.toEqual(carry);
  expect(worker.reads[1].carry).toEqual(carry);
  expect(h.hook.result.current.highlight).toEqual({ matches, current: 0 });
  expect(h.reveals.at(-1)).toBe(matches[0].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test('navigation successfully revealed during a held refresh reveals the carried match after it moves', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  h.stamp(0, worker.session.version());
  await act(async () => { await api.search('the'); api.searchGoTo(4); });
  const before = h.hook.result.current.highlight!.matches[4].displayFrom;
  worker.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  worker.mirror();
  const release = worker.holdRead();
  act(() => { api.getSearchState(); });
  await waitFor(() => expect(worker.reads).toHaveLength(2));
  act(() => { expect(api.searchGoTo(4)?.current).toBe(4); });
  expect(h.reveals.at(-1)).toBe(before);
  await act(async () => { release(); });
  await waitFor(() => expect(h.hook.result.current.highlight?.matches).toHaveLength(6));
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 5 });
  expect(h.hook.result.current.highlight!.current).toBe(5);
  expect(h.hook.result.current.highlight!.matches[5].displayFrom).toBe(before + 4);
  expect(h.reveals.at(-1)).toBe(h.hook.result.current.highlight!.matches[5].displayFrom);
  expect(worker.replicaRequests()).toBe(0);
});

test('worker display placement retries a pending reveal without another document read', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  h.stamp(0, 'older');
  h.layOut(() => false);
  await act(async () => { await api.search('the'); });
  const matches = h.hook.result.current.highlight!.matches;
  const events = h.events.length;
  h.stamp(1, 'older');
  h.hook.rerender({ version: 1 });
  expect(h.events).toHaveLength(events);
  expect(h.reveals).toHaveLength(2);
  h.layOut(() => true);
  h.stamp(2, worker.session.version());
  h.hook.rerender({ version: 2 });
  expect(h.events).toHaveLength(events + 1);
  expect(h.reveals).toHaveLength(3);
  h.stamp(3, worker.session.version());
  h.hook.rerender({ version: 3 });
  expect(h.reveals).toHaveLength(3);
  expect(worker.reads).toHaveLength(1);
  expect(worker.anchorReads).toHaveLength(1);
  expect(h.hook.result.current.highlight!.matches).toBe(matches);
  expect(worker.replicaRequests()).toBe(0);
});

test('worker search carries portable anchors through hand-over and subsequent main edits', async () => {
  const h = await mount();
  const worker = await workerSearch(h);
  const api = h.hook.result.current.api;
  await act(async () => { await api.search('the'); });
  act(() => api.searchGoTo(3));
  const before = h.hook.result.current.highlight!.matches[3].displayFrom;
  expect(worker.replicaRequests()).toBe(0);
  await act(async () => { await requestWorkerOpenReplica(h.session); });
  expect(worker.replicaRequests()).toBe(1);
  expect(workerProposalAuthority(h.session)).toBeNull();
  h.session.insertText({ story: 'body', paraId: h.first, offset: 0 }, 'the ');
  h.hook.rerender({ version: 1 });
  expect(api.getSearchState()).toMatchObject({ total: 6, current: 4 });
  expect(h.hook.result.current.highlight!.matches[4].displayFrom).toBe(before + 4);
  expect(worker.reads).toHaveLength(1);
});

test('finds every body and table match in reading order and walks them', async () => {
  const { hook, reveals, events } = await mount();
  const api = () => hook.result.current.api;
  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await api().search('the');
  });
  expect(state).toEqual({ query: 'the', options: { caseSensitive: false }, total: 5, current: 0 });
  const matches = hook.result.current.highlight!.matches;
  expect(matches).toHaveLength(5);
  expect(matches.every((match, index) => index === 0 || match.displayFrom > matches[index - 1].displayFrom)).toBe(true);
  expect(reveals).toEqual([matches[0].displayFrom]);

  act(() => {
    api().searchPrevious();
  });
  expect(api().getSearchState()?.current).toBe(4);
  act(() => {
    api().searchNext();
  });
  expect(api().getSearchState()?.current).toBe(0);
  act(() => {
    api().searchGoTo(2);
  });
  expect(hook.result.current.highlight?.current).toBe(2);
  expect(reveals.at(-1)).toBe(matches[2].displayFrom);
  expect(api().searchGoTo(9)?.current).toBe(4);
  expect(api().searchGoTo(-1)?.current).toBe(4);
  act(() => {
    api().searchGoTo(2);
  });

  await act(async () => {
    state = await api().search('the', { caseSensitive: true });
  });
  expect(state!.total).toBe(3);

  act(() => api().clearSearch());
  expect(hook.result.current.highlight).toBeNull();
  expect(api().getSearchState()).toBeNull();
  expect(api().searchNext()).toBeNull();
  expect(events.map((event) => event?.current ?? null)).toEqual([0, 4, 0, 2, 4, 4, 2, 0, null]);
});

test('starts at the first match on the page in view', async () => {
  const { hook } = await mount(2);
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  expect(hook.result.current.api.getSearchState()?.current).toBe(2);
});

test('a repeated header on the page in view does not send the search back', async () => {
  const { hook, reveals } = await mount(2, true);
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  expect(hook.result.current.api.getSearchState()?.current).toBe(2);
  expect(reveals).toEqual([hook.result.current.highlight!.matches[2].displayFrom]);
});

test('starts at the nearest later page beyond the look-ahead with non-monotonic match pages', async () => {
  const { first, hook, reveals, pagedEditorRef, host, queries } = await mount(3);
  const farMatch = pagedEditorRef.current!.yrsLocToDisplayPosition({
    story: 'body',
    paraId: first,
    offset: 12,
  })!;
  const displayListQueries = queries(0);
  Object.assign(displayListQueries, {
    displayList: {
      pages: Array.from({ length: 41 }, (_, pageIndex) => ({
        pageIndex,
        primitives: pageIndex === 0
          ? [
              { kind: 'text', docStart: 0, docEnd: farMatch },
              { kind: 'text', docStart: farMatch + 3, docEnd: 100000 },
            ]
          : pageIndex === 40
            ? [{ kind: 'text', docStart: farMatch, docEnd: farMatch + 3 }]
            : [],
      })),
    },
    anchorRect: (position: number) => ({
      pageIndex: position === farMatch ? 40 : 0,
      x: 0,
      y: 0,
      width: 1,
      height: 1,
    }),
  });
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '2';
  canvas.getBoundingClientRect = () => new DOMRect(0, 10, 800, 990);
  host.append(canvas);
  expect(topPageInView(host)).toBe(2);

  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await hook.result.current.api.search('the');
  });
  const matches = hook.result.current.highlight!.matches;
  expect(matches.map((match) => displayListQueries.anchorRect(match.displayFrom)?.pageIndex)).toEqual([0, 40, 0, 0, 0]);
  expect(state).toMatchObject({ total: 5, current: 1 });
  expect(hook.result.current.highlight!.current).toBe(1);
  expect(reveals).toEqual([matches[1].displayFrom]);
});

test('a page in view past every match wraps to the first without reading every match page', async () => {
  const { hook, host, queries } = await mount(3);
  const displayListQueries = queries(0);
  let pageReads = 0;
  Object.assign(displayListQueries, {
    displayList: {
      pages: Array.from({ length: 41 }, (_, pageIndex) => ({
        pageIndex,
        primitives:
          pageIndex === 0
            ? [{ kind: 'text', docStart: 0, docEnd: 100000 }]
            : pageIndex === 40
              ? [{ kind: 'text', docStart: 100001, docEnd: 100010 }]
              : [],
      })),
    },
    anchorRect: () => {
      pageReads += 1;
      return { pageIndex: 0, x: 0, y: 0, width: 1, height: 1 };
    },
  });
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '40';
  canvas.getBoundingClientRect = () => new DOMRect(0, 10, 800, 990);
  host.append(canvas);
  expect(topPageInView(host)).toBe(40);

  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await hook.result.current.api.search('the');
  });
  expect(state).toMatchObject({ total: 5, current: 0 });
  expect(pageReads).toBe(0);
});

test('a document change re-runs the search and keeps the current match', async () => {
  const { session, first, hook, events } = await mount();
  const api = () => hook.result.current.api;
  await act(async () => {
    await api().search('the');
  });
  act(() => {
    api().searchGoTo(3);
  });
  session.insertText({ story: 'body', paraId: first, offset: 0 }, 'the ');
  hook.rerender({ version: 1 });
  expect(api().getSearchState()).toMatchObject({ total: 6, current: 4 });
  expect(events.at(-1)).toMatchObject({ total: 6, current: 4 });

  // a new match lands at the current match's old offset in its own paragraph
  act(() => {
    api().searchGoTo(1);
  });
  const current = hook.result.current.highlight!.matches[1];
  session.insertText({ story: 'body', paraId: first, offset: 0 }, 'the ');
  hook.rerender({ version: 3 });
  expect(api().getSearchState()).toMatchObject({ total: 7, current: 2 });
  expect(hook.result.current.highlight!.matches[2].displayFrom).toBe(current.displayFrom + 4);
  events.splice(0, events.length - 1);

  act(() => {
    api().searchGoTo(0);
  });
  hook.rerender({ version: 4 });
  expect(events).toHaveLength(2);
});

test('a match past the laid-out pages is revealed once the layout reaches it', async () => {
  const { hook, reveals, layOut } = await mount();
  layOut(() => false);
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  expect(reveals).toHaveLength(1);
  hook.rerender({ version: 1 });
  expect(reveals).toHaveLength(2);
  layOut(() => true);
  hook.rerender({ version: 2 });
  hook.rerender({ version: 3 });
  expect(reveals).toHaveLength(3);
});

test('an edit before the layout reaches the match keeps revealing it', async () => {
  const { session, first, hook, reveals, layOut } = await mount();
  layOut(() => false);
  await act(async () => {
    await hook.result.current.api.search('dog');
  });
  session.insertText({ story: 'body', paraId: first, offset: 0 }, 'A ');
  hook.rerender({ version: 1 });
  const moved = hook.result.current.highlight!.matches[0].displayFrom;
  expect(reveals.at(-1)).toBe(moved);
  layOut(() => true);
  hook.rerender({ version: 2 });
  expect(reveals.at(-1)).toBe(moved);
  hook.rerender({ version: 3 });
  expect(reveals).toHaveLength(3);
});

test('clearing wins over a search still flushing input', async () => {
  const { hook, pagedEditorRef } = await mount();
  let release!: () => void;
  const flushed = new Promise<void>((resolve) => {
    release = resolve;
  });
  const editor = pagedEditorRef.current!;
  pagedEditorRef.current = {
    ...editor,
    hasPendingInput: () => true,
    flushPendingInput: () => flushed,
  } as PagedEditorRef;
  let searched!: Promise<unknown>;
  act(() => {
    searched = hook.result.current.api.search('the');
  });
  act(() => hook.result.current.api.clearSearch());
  await act(async () => {
    release();
    await searched;
  });
  expect(hook.result.current.api.getSearchState()).toBeNull();
  expect(hook.result.current.highlight).toBeNull();
});

test('unmounting the editor ends its search', async () => {
  const { hook, events } = await mount();
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  hook.unmount();
  expect(events.at(-1)).toBeNull();
});

test('a search ends with its editor or document', async () => {
  const { hook, pagedEditorRef, events } = await mount();
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  pagedEditorRef.current = null;
  expect(hook.result.current.api.searchNext()).toBeNull();
  expect(events.at(-1)).toBeNull();
});

test('a search without matches still owns the highlights', async () => {
  const { hook } = await mount();
  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await hook.result.current.api.search('zebra');
  });
  expect(state).toMatchObject({ total: 0, current: -1 });
  expect(hook.result.current.highlight).toEqual({ matches: [], current: -1 });
  expect(hook.result.current.api.searchNext()).toMatchObject({ total: 0, current: -1 });
});

test('a reveal against a layout of an older version waits for the current one', async () => {
  const { session, hook, reveals, stamp } = await mount();
  stamp(0, 'older');
  await act(async () => {
    await hook.result.current.api.search('dog');
  });
  expect(reveals).toHaveLength(1);
  stamp(1, session.version());
  hook.rerender({ version: 1 });
  expect(reveals).toHaveLength(2);
  hook.rerender({ version: 1 });
  stamp(2, session.version());
  hook.rerender({ version: 2 });
  expect(reveals).toHaveLength(2);
});

test('navigating right after an edit steps from the carried match', async () => {
  const { session, first, hook, reveals } = await mount();
  const api = () => hook.result.current.api;
  await act(async () => {
    await api().search('the');
  });
  act(() => {
    api().searchGoTo(1);
  });
  session.insertText({ story: 'body', paraId: first, offset: 0 }, 'the ');
  let state = null as DocxSearchState | null;
  act(() => {
    state = api().searchNext();
  });
  // "The" moved to the third match; the next one is the paragraph's "the"
  expect(state).toMatchObject({ total: 6, current: 3 });
  const revealed = reveals.length;
  // the reveal waits for the edited document's layout
  hook.rerender({ version: 1 });
  expect(reveals).toHaveLength(revealed + 1);
  expect(reveals.at(-1)).toBe(hook.result.current.highlight!.matches[3].displayFrom);
});

test('a listener that clears the search ends the round of events', async () => {
  const { hook, events } = await mount();
  const api = () => hook.result.current.api;
  const later: Array<DocxSearchState | null> = [];
  const stop = api().onSearchChange((state) => {
    if (state?.current === 1) api().clearSearch();
  });
  api().onSearchChange((state) => later.push(state));
  await act(async () => {
    await api().search('the');
  });
  act(() => {
    api().searchNext();
  });
  stop();
  expect(later.at(-1)).toBeNull();
  expect(later.filter((state) => state?.current === 1)).toHaveLength(0);
  expect(events.at(-1)).toBeNull();
});

test('clearing or moving on stops following the last reveal onto an unbuilt page', async () => {
  const { hook, signals } = await mount();
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  act(() => {
    hook.result.current.api.searchNext();
  });
  expect(signals.map((signal) => signal.aborted)).toEqual([true, false]);
  act(() => hook.result.current.api.clearSearch());
  expect(signals.every((signal) => signal.aborted)).toBe(true);
});

test('an empty query clears', async () => {
  const { hook } = await mount();
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await hook.result.current.api.search('');
  });
  expect(state).toEqual({ query: '', options: { caseSensitive: false }, total: 0, current: -1 });
  expect(hook.result.current.api.getSearchState()).toBeNull();
});
