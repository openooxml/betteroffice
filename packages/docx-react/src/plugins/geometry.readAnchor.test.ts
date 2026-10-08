import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { AnchorDisplayTarget, YrsSession } from '@betteroffice/docx/yrs';
import {
  clearPresented,
  markPresented,
  stampRevisionPreviewKey,
  stampWorkerFrameVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import { createPluginGeometry, type ReadAnchorTarget } from './geometry';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const RANGE = {
  kind: 'range',
  version: 'v1',
  range: {
    story: 'body',
    start: { paraId: 'p', offset: 0 },
    end: { paraId: 'p', offset: 4 },
    view: 'accepted',
  },
} as const;

function viewerGeometry(options: {
  read?: ReadAnchorTarget;
  pages?: Array<{ unbuilt?: boolean; positionSpan?: [number, number] }>;
  shown?: () => boolean;
  workerVersion?: string | null;
  anchorPage?: number;
  presented?: boolean;
}) {
  const pages = document.createElement('div');
  const pageList = options.pages ?? [{}];
  pageList.forEach((_, index) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    pages.appendChild(canvas);
  });
  const layer = document.createElement('div');
  const bounds = new DOMRect(0, 0, 100, 200);
  pages.getBoundingClientRect = () => bounds;
  layer.getBoundingClientRect = () => bounds;
  for (const canvas of pages.querySelectorAll('canvas')) canvas.getBoundingClientRect = () => bounds;
  const queries = {
    displayList: { pages: pageList.map((page, pageIndex) => ({ pageIndex, width: 100, height: 200, ...page })) },
    pageSize: () => ({ width: 100, height: 200 }),
    pageCount: () => pageList.length,
    rangeRects: (from: number, to: number) =>
      from < 10 && to > 0 && !pageList[0]?.unbuilt
        ? [{ pageIndex: 0, x: from, y: 10, width: Math.min(to, 10) - from, height: 12 }]
        : [],
    anchorRect: (position: number) => ({ pageIndex: options.anchorPage ?? 0, x: position, y: 10, width: 1, height: 12 }),
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: pageIndex * 200, width: 100, height: 200 }),
    hitTestRegions: () => null,
  } as unknown as DisplayListQueries;
  if (options.presented !== false) markPresented(pages, queries.displayList);
  stampRevisionPreviewKey(queries, '');
  if (options.workerVersion !== null) stampWorkerFrameVersion(queries, options.workerVersion ?? 'w1');
  const session = {
    version: () => 'v1',
    getProposals: () => null,
    resolveParagraphAnchor: () => ({ status: 'missing' }),
  } as unknown as YrsSession;
  return createPluginGeometry(
    { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1, pageCount: pageList.length },
    createRenderedDomContext(pages, 1),
    layer,
    options.shown ?? (() => true),
    () => null,
    queries,
    () =>
      options.read
        ? null
        : {
            session,
            presented: true,
            editor: { hasPendingInput: () => false, yrsLocToDisplayPosition: () => 2 },
          },
    () => false,
    undefined,
    options.read
  );
}

const reply = (value: AnchorDisplayTarget | null): ReadAnchorTarget => async () => value;

test('serves range rects from a worker read in a viewer without a session', async () => {
  const geometry = viewerGeometry({
    read: reply({ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }),
  });
  expect(geometry.getAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: true,
    version: 'v1',
    rects: [{ pageIndex: 0, x: 2, width: 4 }],
    unbuiltPages: [],
  });
});

test('targets reach the worker at the worker frame version', async () => {
  const calls: Parameters<ReadAnchorTarget>[] = [];
  const geometry = viewerGeometry({
    workerVersion: 'w7',
    read: async (...args) => {
      calls.push(args);
      return { ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] };
    },
  });
  const revision = { kind: 'revision', revisionId: 'r1' } as const;
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({ ok: true, version: 'v1' });
  expect(await geometry.readAnchorGeometry(revision)).toMatchObject({ ok: true });
  expect(calls).toEqual([
    [{ ...RANGE, version: 'w7' }, 'w7', 0, ''],
    [revision, 'w7', 0, ''],
  ]);
});

test('a range of another version or a frame without a worker version is not read', async () => {
  let reads = 0;
  const read: ReadAnchorTarget = async () => {
    reads += 1;
    return { ok: true, ranges: [], paragraph: 2, hidden: [] };
  };
  expect(await viewerGeometry({ read }).readAnchorGeometry({ ...RANGE, version: 'v0' })).toMatchObject({
    ok: false,
    failure: { code: 'stale-version' },
  });
  expect(await viewerGeometry({ read, workerVersion: null }).readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(reads).toBe(0);
});

test('a target on unbuilt pages reports them instead of refusing', async () => {
  const geometry = viewerGeometry({
    pages: [{}, { unbuilt: true, positionSpan: [100, 200] }],
    read: reply({ ok: true, ranges: [{ from: 150, to: 160 }], paragraph: 150, hidden: [] }),
  });
  const result = await geometry.readAnchorGeometry(RANGE);
  expect(result).toMatchObject({ ok: true, rects: [], unbuiltPages: [1], anchor: { pageIndex: 1 } });
});

test('a target ending on an unbuilt page keeps its built rects and anchors on that page', async () => {
  const geometry = viewerGeometry({
    pages: [{ positionSpan: [0, 99] }, { unbuilt: true, positionSpan: [100, 200] }],
    read: reply({ ok: true, ranges: [{ from: 2, to: 150 }], paragraph: 2, hidden: [] }),
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: true,
    rects: [{ pageIndex: 0, x: 2, width: 8 }],
    unbuiltPages: [1],
    anchor: { pageIndex: 1 },
  });
});

test('an anchor placed on an unbuilt page the ranges do not reach lists that page', async () => {
  const geometry = viewerGeometry({
    pages: [{}, { unbuilt: true }],
    anchorPage: 1,
    read: reply({ ok: true, ranges: [{ from: 150, to: 150 }], paragraph: 150, hidden: [] }),
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: true,
    rects: [],
    unbuiltPages: [1],
    anchor: { pageIndex: 1 },
  });
});

test('an unbuilt page reached only by hidden text is not listed and does not take the anchor', async () => {
  const geometry = viewerGeometry({
    pages: [{ positionSpan: [0, 99] }, { unbuilt: true, positionSpan: [100, 200] }],
    read: reply({ ok: true, ranges: [{ from: 2, to: 150 }], paragraph: 2, hidden: [{ from: 99, to: 150 }] }),
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: true,
    rects: [{ pageIndex: 0, x: 2, width: 8 }],
    unbuiltPages: [],
    anchor: { pageIndex: 0 },
  });
});

test('hidden ranges are subtracted from the worker ranges', async () => {
  const geometry = viewerGeometry({
    read: reply({ ok: true, ranges: [{ from: 2, to: 8 }], paragraph: 2, hidden: [{ from: 4, to: 8 }] }),
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: true,
    rects: [{ x: 2, width: 2 }],
  });
});

test('a reply for a superseded version cannot place rects', async () => {
  const geometry = viewerGeometry({ read: reply(null) });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'stale-version' },
  });
});

test('a layout that stops showing while the read runs refuses', async () => {
  let shown = true;
  const geometry = viewerGeometry({
    shown: () => shown,
    read: async () => {
      shown = false;
      return { ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] };
    },
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
});

test('a frame not yet presented refuses before and after the read', async () => {
  let reads = 0;
  const unpainted = viewerGeometry({
    presented: false,
    read: async () => {
      reads += 1;
      return { ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] };
    },
  });
  expect(await unpainted.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(reads).toBe(0);
  let host: HTMLElement | null = null;
  const repainting = viewerGeometry({
    read: async () => {
      if (host) clearPresented(host);
      return { ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] };
    },
  });
  host = repainting.dom.pagesContainer;
  expect(await repainting.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
});

test('a worker failure is passed through and a rejected read refuses', async () => {
  const failed = viewerGeometry({
    read: reply({ ok: false, failure: { code: 'missing-target', message: 'gone' } }),
  });
  expect(await failed.readAnchorGeometry(RANGE)).toEqual({
    ok: false,
    failure: { code: 'missing-target', message: 'gone' },
  });
  const rejected = viewerGeometry({ read: async () => { throw new Error('worker gone'); } });
  expect(await rejected.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
});

test('without a worker read it answers like getAnchorGeometry', async () => {
  const geometry = viewerGeometry({});
  const paragraph = { kind: 'paragraph', paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' } } as const;
  for (const target of [paragraph, { ...RANGE, version: 'v0' }]) {
    const expected = geometry.getAnchorGeometry(target);
    expect(expected).toMatchObject({ ok: false });
    expect(await geometry.readAnchorGeometry(target)).toEqual(expected);
  }
});
