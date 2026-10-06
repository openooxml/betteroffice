import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { LayoutSelectionGate } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  proposalRevisionPreview,
  proposalSetIdentity,
  type DocxProposalSnapshot,
  type ProposalGeometryMirror,
  type ResidentProposalReply,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createProposalRegistry, type DocxProposalSession } from '@betteroffice/docx/yrs/proposals';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import {
  useLayoutPipeline,
  type UseLayoutPipelineReturn,
} from '../components/DocxEditor/hooks/useLayoutPipeline';
import {
  markPresented,
  revisionPreviewKeyOf,
  sourceVersionOf,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import * as workerProposals from '../components/DocxEditor/internals/workerProposalAuthority';
import { deferWorkerOpenReplica } from '../components/DocxEditor/internals/workerOpenReplica';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import { defineDocxPlugin } from './defineDocxPlugin';
import type { DocxAnchorGeometryResult } from './types';
import { useDocxPluginHost, type UseDocxPluginHostOptions } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const restores: Array<() => void> = [];

afterEach(() => {
  cleanup();
  for (const restore of restores.splice(0)) restore();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function stubSession() {
  const snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion: 0, proposals: [] };
  return {
    version: () => snapshot.version,
    getProposals: () => snapshot,
    onUpdate: () => () => {},
    onProposalChange: () => () => {},
    selection: () => null,
  } as unknown as YrsSession;
}

function fakeAuthority(session: YrsSession) {
  const listeners = new Set<() => void>();
  let mirror: ProposalGeometryMirror | null = null;
  const authority = {
    initialized: true,
    geometry: () => mirror,
    subscribe(listener: () => void) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  } as workerProposals.WorkerProposalAuthority;
  return {
    authority,
    listeners,
    publish(paragraph: number) {
      mirror = {
        version: 'v1',
        previewVersion: 0,
        proposals: proposalSetIdentity(session.getProposals()),
        targets: { proposal: { ok: true, ranges: [], paragraph } },
        hidden: [],
      };
      for (const listener of listeners) listener();
    },
  };
}

test('worker geometry arriving after a frame re-notifies plugins and subscriptions follow the session', async () => {
  const first = stubSession();
  const second = stubSession();
  const firstWorker = fakeAuthority(first);
  const secondWorker = fakeAuthority(second);
  const lookup = workerProposals.workerProposalAuthority;
  const routing = spyOn(workerProposals, 'workerProposalAuthority').mockImplementation((session) =>
    session === first
      ? firstWorker.authority
      : session === second
        ? secondWorker.authority
        : lookup(session)
  );
  restores.push(() => routing.mockRestore());
  const pages = document.createElement('div');
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  pages.append(canvas);
  const layer = document.createElement('div');
  document.body.append(pages, layer);
  restores.push(() => {
    pages.remove();
    layer.remove();
  });
  const bounds = () => new DOMRect(0, 0, 100, 200);
  pages.getBoundingClientRect = bounds;
  canvas.getBoundingClientRect = bounds;
  layer.getBoundingClientRect = bounds;
  const displayList = { pages: [{ pageIndex: 0, width: 100, height: 200 }] } as DisplayList;
  const queries = {
    displayList,
    sourceState: () => ({ status: 'ready' }),
    pageCount: () => 1,
    pageSize: () => ({ width: 100, height: 200 }),
    pageBounds: () => ({ pageIndex: 0, x: 0, y: 0, width: 100, height: 200 }),
    rangeRects: () => [],
    anchorRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 12 }),
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, 'v1');
  stampRevisionPreviewKey(queries, '');
  const layout = { pages: [{}], partial: false };
  stampSourceVersion(layout, 'v1');
  const editor = (session: YrsSession) =>
    ({
      getYrsSession: () => session,
      getLayout: () => layout,
      getSelectionRange: () => null,
      hasPendingInput: () => false,
      yrsLocToDisplayPosition: () => {
        throw new Error('worker geometry must not read replica positions');
      },
    }) as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor(first) };
  const events: (DocxAnchorGeometryResult | null)[] = [];
  const plugin = defineDocxPlugin({
    id: 'test.worker-proposal',
    createState: () => null,
    onEvent(context, event) {
      if (
        event.type === 'layout-change' &&
        workerProposals.workerProposalAuthority(pagedEditorRef.current.getYrsSession()!)?.geometry()
      ) {
        events.push(
          context.geometry?.getAnchorGeometry({ kind: 'proposal', id: 'proposal' }) ?? null
        );
      }
    },
  });
  const options: UseDocxPluginHostOptions = {
    plugins: [plugin],
    pagedEditorRef,
    writeModeRef: { current: 'viewing' },
    mode: 'viewing',
    readOnly: true,
    commands: createDocxCommandController(),
    session: first,
    loadGeneration: 0,
    queries,
    layoutError: null,
    zoom: 1,
    canvasHostRef: { current: pages },
    overlayTarget: layer,
    selectionChangeSubscribersRef: { current: new Set<(state: SelectionState | null) => void>() },
    i18n: undefined,
    onRenderedDomContextReady: undefined,
  };
  markPresented(pages, displayList);
  const { result, rerender, unmount } = renderHook(
    (props: UseDocxPluginHostOptions) => useDocxPluginHost(props),
    { initialProps: options }
  );
  await act(async () => {
    result.current.overlayLayerRef(layer);
    result.current.onRenderedDomContext(createRenderedDomContext(pages, 1), queries);
  });
  await waitFor(() => expect(result.current.activations[0]?.context.geometry).toBeTruthy());
  expect(firstWorker.listeners.size).toBe(1);
  const geometry = result.current.activations[0]!.context.geometry!;
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: false,
    failure: { code: 'unknown-proposal' },
  });
  await act(async () => firstWorker.publish(3));
  await waitFor(() => expect(events.at(-1)).toMatchObject({ ok: true, anchor: { x: 3 } }));
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: true,
    anchor: { x: 3 },
  });
  expect(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'revision' })).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  const before = events.length;
  await act(async () => firstWorker.publish(7));
  await waitFor(() => expect(events.length).toBeGreaterThan(before));
  expect(events.at(-1)).toMatchObject({ ok: true, anchor: { x: 7 } });
  await act(async () => {
    pagedEditorRef.current = editor(second);
    rerender({ ...options, session: second, loadGeneration: 1 });
  });
  expect(firstWorker.listeners.size).toBe(0);
  expect(secondWorker.listeners.size).toBe(1);
  await act(async () => result.current.beginLoad());
  expect(secondWorker.listeners.size).toBe(0);
  unmount();
  expect(secondWorker.listeners.size).toBe(0);
});

test('worker proposal toggles keep the shown layout and mirrored geometry at the latest preview', async () => {
  const frames = new Map<number, FrameRequestCallback>();
  let nextFrame = 0;
  const requestFrame = spyOn(globalThis, 'requestAnimationFrame').mockImplementation((callback) => {
    const id = ++nextFrame;
    frames.set(id, callback);
    return id;
  });
  const cancelFrame = spyOn(globalThis, 'cancelAnimationFrame').mockImplementation((id) => {
    frames.delete(id);
  });
  restores.push(() => { requestFrame.mockRestore(); cancelFrame.mockRestore(); });
  let pipeline: UseLayoutPipelineReturn | null = null;
  const workerPreviews: unknown[] = [];
  const revealed: number[] = [];
  const documentRead = mock(async () => { throw new Error('unexpected document read'); });
  const paragraph = {
    kind: 'session' as const, sessionId: 'session', story: 'body', paraId: 'first',
  };
  const workerRegistry = createProposalRegistry({ version: () => 'v1' } as DocxProposalSession);
  const mainRegistry = createProposalRegistry({ version: () => 'v1' } as DocxProposalSession);
  restores.push(() => { workerRegistry.destroy(); mainRegistry.destroy(); });
  const session = {
    version: () => mainRegistry.snapshot().version,
    getProposals: mainRegistry.snapshot,
    mirrorWorkerDocument: mainRegistry.mirror,
    onUpdate: () => () => {},
    onProposalChange: mainRegistry.subscribe,
    selection: () => null,
    layoutFontRequirementsJson: () => '[]',
  } as unknown as YrsSession;
  const replica = deferWorkerOpenReplica(session, async () => {
    throw new Error('proposal toggles must not hydrate the replica');
  }, () => { throw new Error('unexpected replica fallback'); }, () => {});
  restores.push(() => replica.cancel());
  const authority = workerProposals.registerWorkerProposalAuthority(session, {
    async proposal(op) {
      let result: ResidentProposalReply['result'];
      if (op.kind === 'propose') {
        workerRegistry.mirror({
          version: 'v1',
          proposals: {
            previewVersion: 0,
            entries: op.request.proposals.map((input, index) => ({
              key: String(index),
              suggest: input.suggest,
              record: {
                id: input.id,
                state: 'proposed' as const,
                paragraph,
                revisionIds: [`revision-${index}`],
                changed: true,
              },
            })),
          },
        });
        workerRegistry.mirror(null);
        result = { ok: true, snapshot: workerRegistry.snapshot() };
      } else if (op.kind === 'setStates') {
        result = workerRegistry.setStates(op.request);
      }
      const snapshot = workerRegistry.snapshot();
      return {
        result,
        mirror: { version: snapshot.version, proposals: workerRegistry.exportState() },
        geometry: {
          version: snapshot.version,
          previewVersion: snapshot.previewVersion,
          proposals: proposalSetIdentity(snapshot),
          targets: Object.fromEntries(snapshot.proposals.map(({ id }) => [
            id, { ok: true as const, ranges: [], paragraph: 3 },
          ])),
          navigationTargets: Object.fromEntries(snapshot.proposals.map(({ id, paragraph }) => [
            id, { loc: { story: paragraph.story, paraId: paragraph.paraId, offset: 0 }, position: 3 },
          ])),
          hidden: [],
        },
        changedStories: [],
        updates: [],
        stateVector: new Uint8Array(),
      };
    },
    documentRead,
    handOver: async () => { throw new Error('unexpected hand-over'); },
  }, {
    current: () => true,
    laidOut: async () => {},
    adopted: () => {},
    handedOver: () => {},
    relayout: () => pipeline?.scheduleLayout('remote', true),
    contentChanged: () => {},
  });
  await authority.initialize();
  await authority.propose({
    expectVersion: 'v1',
    proposals: Array.from({ length: 10 }, (_, index) => ({
      id: `proposal-${index}`,
      paragraph,
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'insertText' as const,
      at: 'start' as const,
      text: 'Proposed ',
    })),
  }, async () => {
    throw new Error('proposals must stay in the worker');
  });
  const pages = document.createElement('div');
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  pages.append(canvas);
  const layer = document.createElement('div');
  document.body.append(pages, layer);
  restores.push(() => { pages.remove(); layer.remove(); });
  const bounds = () => new DOMRect(0, 0, 100, 200);
  pages.getBoundingClientRect = bounds;
  canvas.getBoundingClientRect = bounds;
  layer.getBoundingClientRect = bounds;
  const displayList = { pages: [{ pageIndex: 0, width: 100, height: 200 }] } as DisplayList;
  const queries = {
    displayList,
    sourceState: () => ({ status: 'ready' }),
    pageCount: () => 1,
    pageSize: () => ({ width: 100, height: 200 }),
    pageBounds: () => ({ pageIndex: 0, x: 0, y: 0, width: 100, height: 200 }),
    rangeRects: () => [],
    anchorRect: () => ({ pageIndex: 0, x: 3, y: 10, width: 1, height: 12 }),
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, 'v1');
  stampRevisionPreviewKey(queries, '');
  const syncCoordinator = new LayoutSelectionGate();
  const layoutHook = renderHook(() => useLayoutPipeline({
    document: null,
    session,
    renderEnv: {},
    experimentalWorkerOpen: true,
    pageGap: 24,
    zoom: 1,
    deferLayoutPass: () => false,
    residentMeasurementConfig: () => ({
      fontChains: {},
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      compat: { noLeading: false, doNotExpandShiftReturn: false },
      authoritativeShaping: true,
    }),
    pagesContainerRef: { current: null },
    viewportLayoutRef: { current: null },
    syncCoordinator,
    getScrollContainer: () => null,
    layoutInWorker: async (_session, request) => {
      workerPreviews.push(JSON.parse(request).renderEnv.revisionPreview);
      return { layout: { pages: [{}] } as Layout, notesConverged: true };
    },
  }));
  pipeline = layoutHook.result.current;
  await act(async () => pipeline!.runLayoutPipeline());
  const options: UseDocxPluginHostOptions = {
    plugins: [defineDocxPlugin({ id: 'test.worker-toggles', createState: () => null })],
    pagedEditorRef: { current: {
      getYrsSession: () => session,
      getLayout: () => layoutHook.result.current.layout,
      getSelectionRange: () => null,
      hasPendingInput: () => false,
      revealDisplayPosition: (position: number) => { revealed.push(position); return 'scrolled'; },
      yrsLocToDisplayPosition: () => { throw new Error('worker geometry must not read replica positions'); },
    } as unknown as PagedEditorRef },
    writeModeRef: { current: 'viewing' },
    mode: 'viewing',
    readOnly: true,
    commands: createDocxCommandController(),
    session,
    loadGeneration: 0,
    queries,
    layoutError: null,
    zoom: 1,
    canvasHostRef: { current: pages },
    overlayTarget: layer,
    selectionChangeSubscribersRef: { current: new Set<(state: SelectionState | null) => void>() },
    i18n: undefined,
    onRenderedDomContextReady: undefined,
  };
  markPresented(pages, displayList);
  const { result, rerender } = renderHook(useDocxPluginHost, { initialProps: options });
  await act(async () => {
    result.current.overlayLayerRef(layer);
    result.current.onRenderedDomContext(createRenderedDomContext(pages, 1), queries);
  });
  expect(result.current.activations[0]!.context.geometry!.layout.previewVersion).toBe(0);
  for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
    const before = session.getProposals();
    await act(async () => {
      expect(await authority.setStates({
        expectVersion: before.version,
        expectPreviewVersion: before.previewVersion,
        changes: [{ id: 'proposal-0', state }],
      }, async () => { throw new Error('decisions must stay in the worker'); })).toMatchObject({
        ok: true, snapshot: { previewVersion: before.previewVersion + 1 },
      });
    });
    await act(async () => {
      const pending = [...frames];
      for (const [id, callback] of pending) {
        if (frames.delete(id)) callback(performance.now());
      }
    });
    const layout = layoutHook.result.current.layout!;
    const shown = { ...queries, displayList: { ...displayList } } as DisplayListQueries;
    stampSourceVersion(shown, sourceVersionOf(layout));
    stampRevisionPreviewKey(shown, revisionPreviewKeyOf(layout)!);
    await act(async () => {
      markPresented(pages, shown.displayList!);
      rerender({ ...options, queries: shown });
      result.current.onRenderedDomContext(createRenderedDomContext(pages, 1), shown);
    });
    const snapshot = session.getProposals();
    expect(snapshot.previewVersion).toBe(before.previewVersion + 1);
    expect(snapshot.proposals[0]!.state).toBe(state);
    expect(snapshot).toEqual(workerRegistry.snapshot());
    expect(workerPreviews.at(-1)).toEqual(proposalRevisionPreview(snapshot));
    expect(result.current.heldGeometry?.layout.previewVersion).toBe(snapshot.previewVersion);
    const geometry = result.current.activations[0]!.context.geometry;
    expect(geometry?.layout.previewVersion).toBe(snapshot.previewVersion);
    for (const { id } of snapshot.proposals) {
      expect(geometry?.getAnchorGeometry({ kind: 'proposal', id })).toMatchObject({ ok: true });
      expect(result.current.heldGeometry?.getAnchorGeometry({ kind: 'proposal', id })).toMatchObject({
        ok: true,
      });
    }
  }
  const snapshot = session.getProposals();
  expect(snapshot.previewVersion).toBe(4);
  expect(snapshot).toEqual(workerRegistry.snapshot());
  const geometry = result.current.activations[0]!.context.geometry!;
  expect(geometry.layout.previewVersion).toBe(snapshot.previewVersion);
  for (const { id } of snapshot.proposals) {
    expect(geometry.getAnchorGeometry({ kind: 'proposal', id })).toMatchObject({ ok: true });
  }
  expect(replica.started).toBe(false);
  expect(await result.current.activations[0]!.context.navigation.scrollToParagraph(paragraph, {
    expectVersion: snapshot.version,
  })).toEqual({ ok: true });
  expect(revealed).toEqual([3]);
  expect(documentRead).not.toHaveBeenCalled();
  expect(replica.started).toBe(false);
});
