import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxProposalSnapshot, ResidentDocumentRead, ResidentEngineWorkerClient, YrsSession } from '@betteroffice/docx/yrs';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { stampWorkerFrameVersion } from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import { defineDocxPlugin } from './defineDocxPlugin';
import { useDocxPluginHost, type UseDocxPluginHostOptions } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const elements: HTMLElement[] = [];
afterEach(() => { cleanup(); for (const element of elements.splice(0)) element.remove(); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function queries(version: string, y = 20): DisplayListQueries {
  const result = {
    displayList: { pages: [{ pageIndex: 0, width: 100, height: 200 }] },
    pageCount: () => 1,
    pageSize: () => ({ width: 100, height: 200 }),
    anchorRect: () => ({ pageIndex: 0, x: 0, y, width: 1, height: 10 }),
    sourceState: () => ({ status: 'ready' }),
    whenReady: () => Promise.resolve(),
  } as unknown as DisplayListQueries;
  stampWorkerFrameVersion(result, version);
  return result;
}

function setup(read: ResidentEngineWorkerClient['documentRead']) {
  let snapshot: DocxProposalSnapshot = { version: 'A', previewVersion: 0, proposals: [] };
  let update = () => {};
  const session = {
    version: () => snapshot.version,
    getProposals: () => snapshot,
    onUpdate: (listener: () => void) => { update = listener; return () => {}; },
    onProposalChange: () => () => {},
    paragraphs: () => { throw new Error('viewer paragraph read'); },
    selection: () => { throw new Error('viewer selection read'); },
  } as unknown as YrsSession;
  const pages = document.createElement('div');
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  pages.append(canvas);
  const target = document.createElement('div');
  document.body.append(pages, target);
  elements.push(pages, target);
  pages.getBoundingClientRect = canvas.getBoundingClientRect = target.getBoundingClientRect = () => new DOMRect(0, 0, 100, 200);
  const editor = {
    getYrsSession: () => { throw new Error('viewer session access'); },
    getLayout: () => null,
    getSelectionRange: () => null,
    hasPendingInput: () => false,
  } as unknown as PagedEditorRef;
  const plugin = defineDocxPlugin({
    id: 'test.viewer-sidebar',
    createState: () => null,
    getSidebarItems: (context) => [{
      id: 'note', anchor: { version: context.snapshot.version, story: 'body', paraId: 'p1' }, render: () => null,
    }],
  });
  const options: UseDocxPluginHostOptions = {
    plugins: [plugin], pagedEditorRef: { current: editor }, writeModeRef: { current: 'viewing' },
    mode: 'viewing', readOnly: true, commands: createDocxCommandController(), session,
    loadGeneration: 0, queries: queries('A'), viewerDocumentRead: read, layoutError: null,
    zoom: 1, canvasHostRef: { current: pages }, overlayTarget: target,
    selectionChangeSubscribersRef: { current: new Set() }, i18n: undefined, onRenderedDomContextReady: undefined,
  };
  return {
    options,
    changeVersion(version: string) { snapshot = { ...snapshot, version }; update(); },
  };
}

test('managed viewer cards cache worker targets and reproject without accessing a session', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: requests.length === 1 ? 'A' : 'B', value: { loc: { story: 'body', paraId: 'p1', offset: 0 }, position: 7 } };
  }) as ResidentEngineWorkerClient['documentRead'];
  const view = setup(read);
  const { result, rerender } = renderHook(useDocxPluginHost, { initialProps: view.options });
  await waitFor(() => expect(result.current.sidebarItems[0]?.fixedY).toBe(20));
  await act(async () => result.current.publishSelection());
  rerender({ ...view.options, queries: queries('A', 40) });
  await waitFor(() => expect(result.current.sidebarItems[0]?.fixedY).toBe(40));
  expect(requests).toEqual([{ kind: 'navigationTarget', story: 'body', paraId: 'p1' }]);
  await act(async () => view.changeVersion('B'));
  expect(result.current.sidebarItems[0]?.hidden).toBe(true);
  rerender({ ...view.options, queries: queries('B', 60) });
  await waitFor(() => expect(result.current.sidebarItems[0]?.fixedY).toBe(60));
  expect(requests).toHaveLength(2);
});

test('a navigation reply for a superseded viewer version cannot place its card', async () => {
  const pending: Array<(reply: { version: string; value: { loc: { story: string; paraId: string; offset: number }; position: number } }) => void> = [];
  const read = (() => new Promise<unknown>((resolve) => { pending.push(resolve); })) as ResidentEngineWorkerClient['documentRead'];
  const view = setup(read);
  const { result, rerender } = renderHook(useDocxPluginHost, { initialProps: view.options });
  await waitFor(() => expect(pending).toHaveLength(1));
  await act(async () => view.changeVersion('B'));
  rerender({ ...view.options, queries: queries('B') });
  await waitFor(() => expect(pending).toHaveLength(2));
  const target = { loc: { story: 'body', paraId: 'p1', offset: 0 }, position: 7 };
  await act(async () => pending[0]!({ version: 'A', value: target }));
  expect(result.current.sidebarItems[0]?.hidden).toBe(true);
  await act(async () => pending[1]!({ version: 'B', value: target }));
  await waitFor(() => expect(result.current.sidebarItems[0]?.fixedY).toBe(20));
});
