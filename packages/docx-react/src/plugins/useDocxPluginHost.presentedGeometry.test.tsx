import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { DocxProposalSnapshot, YrsSession } from '@betteroffice/docx/yrs';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import {
  markPresented,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import { defineDocxPlugin } from './defineDocxPlugin';
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

function setup() {
  const snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion: 0, proposals: [] };
  const session = {
    version: () => snapshot.version,
    getProposals: () => snapshot,
    onUpdate: () => () => {},
    onProposalChange: () => () => {},
    selection: () => null,
  } as unknown as YrsSession;
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
    anchorRect: () => null,
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, 'v1');
  stampRevisionPreviewKey(queries, '');
  const layout = { pages: [{}], partial: false };
  stampSourceVersion(layout, 'v1');
  const editor = {
    getYrsSession: () => session,
    getLayout: () => layout,
    getSelectionRange: () => null,
    hasPendingInput: () => false,
  } as unknown as PagedEditorRef;
  const events: boolean[] = [];
  const plugin = defineDocxPlugin({
    id: 'test.presented-geometry',
    createState: () => null,
    onEvent(context, event) {
      if (event.type === 'layout-change' && event.layout) events.push(context.geometry !== null);
    },
  });
  const options: UseDocxPluginHostOptions = {
    plugins: [plugin],
    pagedEditorRef: { current: editor },
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
  const { result } = renderHook(() => useDocxPluginHost(options));
  return {
    events,
    result,
    present: () => markPresented(pages, displayList),
    attachLayer: () => result.current.overlayLayerRef(layer),
    emitDom: () =>
      result.current.onRenderedDomContext(createRenderedDomContext(pages, 1), queries),
  };
}

const frames = (count: number) =>
  act(async () => {
    for (let i = 0; i < count; i++) {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    }
  });

const rebuild = async () => {
  await act(async () => {
    window.dispatchEvent(new Event('resize'));
  });
  await frames(3);
};

test('geometry that arrives after its layout was presented reaches plugins in one layout change', async () => {
  const view = setup();
  view.present();
  await act(async () => view.emitDom());
  await waitFor(() => expect(view.events.length).toBeGreaterThan(0));
  await frames(3);
  expect(view.events.every((carried) => !carried)).toBe(true);
  const before = view.events.length;

  await act(async () => view.attachLayer());
  await waitFor(() => expect(view.events.length).toBe(before + 1));
  expect(view.events.at(-1)).toBe(true);
  expect(view.result.current.activations[0]?.context.geometry).toBeTruthy();

  await frames(3);
  await rebuild();
  expect(view.events.length).toBe(before + 1);
});

test('geometry that exists before its layout is presented adds no layout change', async () => {
  const view = setup();
  await act(async () => {
    view.attachLayer();
    view.emitDom();
  });
  await frames(3);
  await act(async () => view.present());
  await frames(3);
  expect(view.events).toEqual([false, true]);

  await rebuild();
  expect(view.events).toEqual([false, true]);
});
