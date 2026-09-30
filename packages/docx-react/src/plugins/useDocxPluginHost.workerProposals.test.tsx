import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  proposalSetIdentity,
  type DocxProposalSnapshot,
  type ProposalGeometryMirror,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import {
  markPresented,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import * as workerProposals from '../components/DocxEditor/internals/workerProposalAuthority';
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
