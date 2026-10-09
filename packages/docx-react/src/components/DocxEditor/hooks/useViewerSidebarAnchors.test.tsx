import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries, TrackedChangesResult } from '@betteroffice/docx/layout/render';
import type { DocxSidebarRead, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { stampWorkerFrameVersion } from '../internals/layoutProvenance';
import type { ViewerCommentRanges } from '../internals/viewerSidebarReads';
import { useViewerSidebarAnchors } from './useViewerSidebarAnchors';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const hosts: HTMLElement[] = [];
afterEach(() => { cleanup(); for (const host of hosts.splice(0)) host.remove(); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const value: DocxSidebarRead = {
  comments: [{ id: '7', anchors: [{ start: { story: 'body', position: 5 }, end: { story: 'body', position: 15 } }] }],
  revisions: [{ key: 'revision-11', start: { story: 'body', position: 20 } }],
  trackedChanges: { entries: [{ type: 'insertion', text: 'Added', author: 'Writer', revisionId: 11, from: 20, to: 25 }], commentToRevision: [] },
};

function setup(read: ResidentEngineWorkerClient['documentRead']) {
  const target = document.createElement('div');
  const host = document.createElement('div');
  host.classList.add('canvas-pages');
  target.append(host);
  document.body.append(target);
  hosts.push(target);
  target.getBoundingClientRect = () => new DOMRect(0, 0, 600, 1800);
  for (let page = 0; page < 2; page += 1) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(page);
    canvas.getBoundingClientRect = () => new DOMRect(0, 100 + page * 900, 600, 800);
    host.append(canvas);
  }
  const tracked: TrackedChangesResult[] = [];
  const positions: Map<string, number>[] = [];
  const ranges: ViewerCommentRanges[] = [];
  const options = {
    read, zoom: 1, commentIds: [7],
    canvasHostRef: { current: host }, pagesContainerRef: { current: null }, overlayTarget: target,
    onTracked: (result: TrackedChangesResult) => { tracked.push(result); },
    onPositions: (result: Map<string, number>) => { positions.push(result); },
    onRanges: (result: ViewerCommentRanges) => { ranges.push(result); },
  };
  return { options, tracked, positions, ranges };
}

function queries(version: string, pages = 1): DisplayListQueries {
  const result = {
    displayList: {}, whenReady: () => Promise.resolve(),
    pageCount: () => pages, pageSize: () => ({ width: 600, height: 800 }),
    anchorRect: (position: number) => position === 20 && pages === 1 ? null : {
      pageIndex: position === 20 ? 1 : 0, x: 0, y: position, width: 1, height: 10,
    },
  } as unknown as DisplayListQueries;
  stampWorkerFrameVersion(result, version);
  return result;
}

test('viewer anchors emit ranges and positions and reproject more pages from the cached read', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: 'A', value };
  }) as ResidentEngineWorkerClient['documentRead'];
  const view = setup(read);
  const { rerender } = renderHook((shown: DisplayListQueries) => useViewerSidebarAnchors({ ...view.options, queries: shown }),
    { initialProps: queries('A') });
  await waitFor(() => expect(view.positions.at(-1)).toEqual(new Map([['comment-7', 105]])));
  expect(view.ranges.at(-1)).toEqual(new Map([[7, { from: 5, to: 15 }]]));
  await act(async () => rerender(queries('A', 2)));
  await waitFor(() => expect(view.positions.at(-1)).toEqual(new Map([['comment-7', 105], ['revision-11', 1020]])));
  expect(requests).toHaveLength(1);
  expect(view.tracked).toHaveLength(1);
  expect(view.ranges.at(-1)).toBe(view.ranges[0]);
});

test('a cancelled run cannot publish a delayed sidebar reply', async () => {
  const pending = new Map<string, (reply: { version: string; value: DocxSidebarRead }) => void>();
  const read = ((request: ResidentDocumentRead) => new Promise<unknown>((resolve) => {
    if ('expectVersion' in request) pending.set(request.expectVersion, resolve);
  })) as ResidentEngineWorkerClient['documentRead'];
  const view = setup(read);
  const { rerender } = renderHook((shown: DisplayListQueries) => useViewerSidebarAnchors({ ...view.options, queries: shown }),
    { initialProps: queries('A') });
  await waitFor(() => expect(pending.has('A')).toBe(true));
  await act(async () => rerender(queries('B')));
  await waitFor(() => expect(pending.has('B')).toBe(true));
  await act(async () => pending.get('A')!({ version: 'A', value }));
  expect(view.positions).toEqual([]);
  expect(view.tracked).toEqual([]);
  expect(view.ranges).toEqual([]);
  await act(async () => pending.get('B')!({ version: 'B', value }));
  await waitFor(() => expect(view.positions).toHaveLength(1));
  expect(view.tracked).toHaveLength(1);
});

test('comment ranges from an older version are withdrawn when a newer frame is shown', async () => {
  const pending = new Map<string, (reply: { version: string; value: DocxSidebarRead }) => void>();
  const read = ((request: ResidentDocumentRead) => request.kind === 'sidebar' && request.expectVersion === 'A'
    ? Promise.resolve({ version: 'A', value })
    : new Promise<unknown>((resolve) => {
      if ('expectVersion' in request) pending.set(request.expectVersion, resolve);
    })) as ResidentEngineWorkerClient['documentRead'];
  const view = setup(read);
  const { rerender } = renderHook((shown: DisplayListQueries) => useViewerSidebarAnchors({ ...view.options, queries: shown }),
    { initialProps: queries('A') });
  await waitFor(() => expect(view.ranges.at(-1)).toEqual(new Map([[7, { from: 5, to: 15 }]])));
  await act(async () => rerender(queries('B')));
  expect(view.ranges.at(-1)).toEqual(new Map());
  await waitFor(() => expect(pending.has('B')).toBe(true));
  await act(async () => pending.get('B')!({ version: 'B', value }));
  await waitFor(() => expect(view.ranges.at(-1)).toEqual(new Map([[7, { from: 5, to: 15 }]])));
});
