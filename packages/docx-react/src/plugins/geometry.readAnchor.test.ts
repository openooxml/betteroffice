import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  proposalSetIdentity,
  type AnchorDisplayTarget,
  type DocxProposalSnapshot,
  type ProposalGeometryMirror,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  clearPresented,
  markPresented,
  stampRevisionPreviewKey,
  stampWorkerFrameVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import {
  createAnchorReadCache,
  createPluginGeometry,
  type AnchorReadCache,
  type ReadAnchorTargets,
} from './geometry';
import type { DocxGeometryTarget } from './types';

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
  read?: ReadAnchorTargets;
  cache?: AnchorReadCache;
  pages?: Array<{ unbuilt?: boolean; positionSpan?: [number, number]; hfParts?: { header?: string } }>;
  shown?: () => boolean;
  workerVersion?: string | null;
  anchorPage?: number;
  presented?: boolean;
  paragraph?: { from: number; length: number };
  layout?: { id?: string; version?: string; previewVersion?: number; zoom?: number };
  previewKey?: string;
  proposals?: ProposalGeometryMirror['targets'];
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
  stampRevisionPreviewKey(queries, options.previewKey ?? '');
  if (options.workerVersion !== null) stampWorkerFrameVersion(queries, options.workerVersion ?? 'w1');
  const snapshot: DocxProposalSnapshot | null = options.proposals
    ? {
        version: 'v1',
        previewVersion: 0,
        proposals: Object.keys(options.proposals).map((id) => ({
          id,
          state: 'proposed',
          paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' },
          revisionIds: [],
          changed: false,
        })),
      }
    : null;
  const mirror: ProposalGeometryMirror | undefined =
    snapshot && options.proposals
      ? { version: 'v1', previewVersion: 0, proposals: proposalSetIdentity(snapshot), targets: options.proposals, hidden: [] }
      : undefined;
  const session = {
    version: () => 'v1',
    getProposals: () => snapshot,
    resolveParagraphAnchor: (anchor: unknown) =>
      options.paragraph ? { status: 'found', anchor } : { status: 'missing' },
    paragraphSpans: () => [{ paraId: 'p', length: options.paragraph?.length ?? 0 }],
  } as unknown as YrsSession;
  const base = options.paragraph?.from ?? 2;
  return createPluginGeometry(
    { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1, ...options.layout, pageCount: pageList.length },
    createRenderedDomContext(pages, options.layout?.zoom ?? 1),
    layer,
    options.shown ?? (() => true),
    () => null,
    queries,
    (target) =>
      options.read && !options.paragraph && !(mirror && target.kind === 'proposal')
        ? null
        : {
            session,
            ...(mirror ? { proposalGeometry: mirror } : {}),
            presented: true,
            editor: {
              hasPendingInput: () => false,
              yrsLocToDisplayPosition: ({ offset }: { offset: number }) => base + offset,
            },
          },
    () => false,
    undefined,
    options.read && { read: options.read, cache: options.cache ?? createAnchorReadCache() }
  );
}

const reply =
  (value: AnchorDisplayTarget | null): ReadAnchorTargets =>
  async (targets) =>
    value && targets.map(() => value);

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
  const calls: Parameters<ReadAnchorTargets>[] = [];
  const geometry = viewerGeometry({
    workerVersion: 'w7',
    read: async (...args) => {
      calls.push(args);
      return args[0].map(() => ({ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }));
    },
  });
  const revision = { kind: 'revision', revisionId: 'r1' } as const;
  expect(await geometry.readAnchorGeometry(RANGE)).toMatchObject({ ok: true, version: 'v1' });
  expect(await geometry.readAnchorGeometry(revision)).toMatchObject({ ok: true });
  expect(calls).toEqual([
    [[{ ...RANGE, version: 'w7' }], 'w7', 0, ''],
    [[revision], 'w7', 0, ''],
  ]);
});

test('a range of another version or a frame without a worker version is not read', async () => {
  let reads = 0;
  const read: ReadAnchorTargets = async (targets) => {
    reads += 1;
    return targets.map(() => ({ ok: true, ranges: [], paragraph: 2, hidden: [] }));
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

const PARAGRAPH = {
  kind: 'paragraph',
  paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' },
} as const;
const NEXT_UNBUILT = [{}, { unbuilt: true, positionSpan: [10, 20] as [number, number] }];

async function withoutUnbuilt(geometry: ReturnType<typeof viewerGeometry>) {
  const result = await geometry.readAnchorGeometry(PARAGRAPH);
  expect(result).toMatchObject({ ok: true, unbuiltPages: [] });
  const { unbuiltPages: _, ...rest } = result as Extract<typeof result, { ok: true }>;
  return rest;
}

test('a range ending where an unbuilt page starts answers like the sync geometry', async () => {
  const geometry = viewerGeometry({
    pages: NEXT_UNBUILT,
    paragraph: { from: 2, length: 8 },
    read: reply({ ok: true, ranges: [{ from: 2, to: 10 }], paragraph: 2, hidden: [] }),
  });
  const sync = geometry.getAnchorGeometry(PARAGRAPH);
  expect(sync).toMatchObject({ ok: true, rects: [{ pageIndex: 0, x: 2, width: 8 }], anchor: { pageIndex: 0 } });
  expect(await withoutUnbuilt(geometry)).toEqual(sync as Extract<typeof sync, { ok: true }>);
});

test('a hidden suffix starting where an unbuilt page starts answers like its shown prefix', async () => {
  const prefix = viewerGeometry({ pages: NEXT_UNBUILT, paragraph: { from: 2, length: 8 } });
  const geometry = viewerGeometry({
    pages: NEXT_UNBUILT,
    read: reply({ ok: true, ranges: [{ from: 2, to: 15 }], paragraph: 2, hidden: [{ from: 10, to: 15 }] }),
  });
  const sync = prefix.getAnchorGeometry(PARAGRAPH);
  expect(sync).toMatchObject({ ok: true, anchor: { pageIndex: 0 } });
  expect(await withoutUnbuilt(geometry)).toEqual(sync as Extract<typeof sync, { ok: true }>);
});

test('a caret at the start of an unbuilt page stays where the sync geometry puts it', async () => {
  const geometry = viewerGeometry({
    pages: NEXT_UNBUILT,
    paragraph: { from: 10, length: 0 },
    read: reply({ ok: true, ranges: [{ from: 10, to: 10 }], paragraph: 10, hidden: [] }),
  });
  const sync = geometry.getAnchorGeometry(PARAGRAPH);
  expect(sync).toMatchObject({ ok: true, rects: [], anchor: { pageIndex: 0, x: 10 } });
  expect(await withoutUnbuilt(geometry)).toEqual(sync as Extract<typeof sync, { ok: true }>);
});

test('a caret inside an unbuilt page anchors on that page', async () => {
  const geometry = viewerGeometry({
    pages: NEXT_UNBUILT,
    read: reply({ ok: true, ranges: [{ from: 12, to: 12 }], paragraph: 2, hidden: [] }),
  });
  expect(await geometry.readAnchorGeometry(PARAGRAPH)).toMatchObject({
    ok: true,
    rects: [],
    unbuiltPages: [1],
    anchor: { pageIndex: 1 },
  });
});

test('fully built pages answer like the sync geometry', async () => {
  const geometry = viewerGeometry({
    pages: [{}, {}],
    paragraph: { from: 2, length: 8 },
    read: reply({ ok: true, ranges: [{ from: 2, to: 10 }], paragraph: 2, hidden: [] }),
  });
  const sync = geometry.getAnchorGeometry(PARAGRAPH);
  expect(sync).toMatchObject({ ok: true });
  expect(await withoutUnbuilt(geometry)).toEqual(sync as Extract<typeof sync, { ok: true }>);
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
    read: async (targets) => {
      shown = false;
      return targets.map(() => ({ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }));
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
    read: async (targets) => {
      reads += 1;
      return targets.map(() => ({ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }));
    },
  });
  expect(await unpainted.readAnchorGeometry(RANGE)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(reads).toBe(0);
  let host: HTMLElement | null = null;
  const repainting = viewerGeometry({
    read: async (targets) => {
      if (host) clearPresented(host);
      return targets.map(() => ({ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }));
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

const REVISION = { kind: 'revision', revisionId: 'r1' } as const;
const PROPOSAL = { kind: 'proposal', id: 'p1' } as const;

function counted(answer: (target: DocxGeometryTarget) => AnchorDisplayTarget = () => ({
  ok: true,
  ranges: [{ from: 2, to: 6 }],
  paragraph: 2,
  hidden: [],
})) {
  const calls: (readonly DocxGeometryTarget[])[] = [];
  const read: ReadAnchorTargets = async (targets) => {
    calls.push(targets);
    return targets.map(answer);
  };
  return { read, calls };
}

test('a batch answers in input order with one worker read for its distinct targets', async () => {
  const { read, calls } = counted((target) =>
    target.kind === 'revision'
      ? { ok: false, failure: { code: 'missing-target', message: 'gone' } }
      : { ok: true, ranges: [{ from: 3, to: 5 }], paragraph: 3, hidden: [] }
  );
  const geometry = viewerGeometry({ read });
  const answers = await geometry.readAnchorGeometries([RANGE, PROPOSAL, REVISION, { ...RANGE, version: 'v0' }, RANGE]);
  expect(answers).toMatchObject([
    { ok: true, layoutId: 'layout', rects: [{ x: 3, width: 2 }], unbuiltPages: [] },
    { ok: false, failure: { code: 'layout-unavailable' } },
    { ok: false, failure: { code: 'missing-target' } },
    { ok: false, failure: { code: 'stale-version' } },
    { ok: true, rects: [{ x: 3, width: 2 }] },
  ]);
  expect(calls).toEqual([[{ ...RANGE, version: 'w1' }, REVISION]]);
  expect(await geometry.readAnchorGeometries([])).toEqual([]);
  expect(calls).toHaveLength(1);
});

test('worker replies are placed again on later layouts of any zoom or main version', async () => {
  const { read, calls } = counted();
  const cache = createAnchorReadCache();
  const first = await viewerGeometry({ read, cache, layout: { id: 'a' } }).readAnchorGeometries([RANGE, REVISION]);
  const repainted = viewerGeometry({ read, cache, layout: { id: 'b' } });
  const second = await repainted.readAnchorGeometries([REVISION, RANGE]);
  expect(second).toEqual([first[1], first[0]].map((answer) => ({ ...answer!, layoutId: 'b' })));
  expect(await repainted.readAnchorGeometry(RANGE)).toEqual(second[1]!);
  for (const layout of [{ zoom: 2 }, { version: 'v2' }]) {
    expect(await viewerGeometry({ read, cache, layout }).readAnchorGeometry(REVISION)).toMatchObject({ ok: true });
  }
  expect(calls).toHaveLength(1);
});

test('a new worker version, preview version or preview asks the worker again', async () => {
  const { read, calls } = counted();
  const cache = createAnchorReadCache();
  const variants = [{}, { workerVersion: 'w2' }, { layout: { previewVersion: 1 } }, { previewKey: 'accepted' }, {}];
  for (const variant of variants) {
    expect(await viewerGeometry({ read, cache, ...variant }).readAnchorGeometry(REVISION)).toMatchObject({ ok: true });
  }
  expect(calls).toHaveLength(variants.length);
});

test('a reply placed on unbuilt pages completes once they build, without another read', async () => {
  const { read, calls } = counted(() => ({ ok: true, ranges: [{ from: 2, to: 150 }], paragraph: 2, hidden: [] }));
  const cache = createAnchorReadCache();
  const span = (unbuilt: boolean) => [
    { positionSpan: [0, 99] as [number, number] },
    { unbuilt, positionSpan: [100, 200] as [number, number] },
  ];
  expect(await viewerGeometry({ read, cache, pages: span(true) }).readAnchorGeometry(REVISION)).toMatchObject({
    ok: true,
    unbuiltPages: [1],
  });
  expect(await viewerGeometry({ read, cache, pages: span(false) }).readAnchorGeometry(REVISION)).toMatchObject({
    ok: true,
    rects: [{ pageIndex: 0 }],
    unbuiltPages: [],
  });
  expect(calls).toHaveLength(1);
});

test('failed and superseded replies are never reused', async () => {
  const cache = createAnchorReadCache();
  const failing: Array<[ReadAnchorTargets, string]> = [
    [reply({ ok: false, failure: { code: 'missing-target', message: 'gone' } }), 'missing-target'],
    [reply(null), 'stale-version'],
    [
      async () => {
        throw new Error('worker gone');
      },
      'layout-unavailable',
    ],
  ];
  let reads = 0;
  for (const [read, code] of failing) {
    const counting: ReadAnchorTargets = (...args) => {
      reads += 1;
      return read(...args);
    };
    for (let round = 0; round < 2; round += 1) {
      expect(await viewerGeometry({ read: counting, cache }).readAnchorGeometry(RANGE)).toMatchObject({
        ok: false,
        failure: { code },
      });
    }
  }
  expect(reads).toBe(failing.length * 2);
});

test('a reply that arrives after its preview was replaced is not kept for the new one', async () => {
  let release: (value: readonly AnchorDisplayTarget[]) => void = () => {};
  const calls: string[] = [];
  const cache = createAnchorReadCache();
  const held = viewerGeometry({
    cache,
    read: (_, __, ___, previewKey) => {
      calls.push(previewKey);
      return new Promise((resolve) => {
        release = resolve;
      });
    },
  }).readAnchorGeometry(REVISION);
  const failing = viewerGeometry({
    cache,
    previewKey: 'b',
    read: async (targets, _, __, previewKey) => {
      calls.push(previewKey);
      return targets.map(() => ({ ok: false, failure: { code: 'missing-target', message: 'gone' } }));
    },
  });
  expect(await failing.readAnchorGeometry(REVISION)).toMatchObject({ ok: false });
  release([{ ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] }]);
  expect(await held).toMatchObject({ ok: true });
  expect(await failing.readAnchorGeometry(REVISION)).toMatchObject({ ok: false });
  expect(calls).toEqual(['', 'b', 'b']);
});

test('a reply that arrives after the layout stopped showing is not kept', async () => {
  let shown = true;
  const { read, calls } = counted();
  const cache = createAnchorReadCache();
  const hiding: ReadAnchorTargets = async (...args) => {
    shown = false;
    return read(...args);
  };
  expect(await viewerGeometry({ read: hiding, cache, shown: () => shown }).readAnchorGeometry(REVISION)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(await viewerGeometry({ read, cache }).readAnchorGeometry(REVISION)).toMatchObject({ ok: true });
  expect(calls).toHaveLength(2);
});

test('the reply cache is cleared once it holds its limit', async () => {
  const { read, calls } = counted();
  const cache = createAnchorReadCache(2);
  const revision = (revisionId: string) => ({ kind: 'revision', revisionId }) as const;
  const geometry = viewerGeometry({ read, cache });
  await geometry.readAnchorGeometries([revision('a'), revision('b'), revision('c')]);
  expect(cache.replies.size).toBe(1);
  await geometry.readAnchorGeometries([revision('c')]);
  expect(calls).toHaveLength(1);
  await geometry.readAnchorGeometries([revision('a')]);
  expect(calls).toHaveLength(2);
});

test('targets are sent and cached by their known fields only', async () => {
  const { read, calls } = counted();
  const geometry = viewerGeometry({ read });
  const noisy = { revisionId: 'r1', kind: 'revision', note: 'host data' } as unknown as DocxGeometryTarget;
  const range = {
    range: { ...RANGE.range, extra: true, start: { offset: 0, paraId: 'p', extra: 1 } },
    kind: 'range',
    version: 'v1',
  } as unknown as DocxGeometryTarget;
  await geometry.readAnchorGeometries([noisy, REVISION, range, RANGE]);
  expect(calls).toEqual([[REVISION, { ...RANGE, version: 'w1' }]]);
});

test('a superseded worker version makes every read target of the batch stale', async () => {
  const geometry = viewerGeometry({ read: reply(null) });
  expect(await geometry.readAnchorGeometries([RANGE, PROPOSAL, REVISION])).toMatchObject([
    { ok: false, failure: { code: 'stale-version' } },
    { ok: false, failure: { code: 'layout-unavailable' } },
    { ok: false, failure: { code: 'stale-version' } },
  ]);
});

test('concurrent reads of one target share a single worker read', async () => {
  let release: (value: readonly AnchorDisplayTarget[]) => void = () => {};
  const calls: (readonly DocxGeometryTarget[])[] = [];
  const read: ReadAnchorTargets = (targets) => {
    calls.push(targets);
    return new Promise((resolve) => {
      release = resolve;
    });
  };
  const cache = createAnchorReadCache();
  const first = viewerGeometry({ read, cache, layout: { id: 'a' } }).readAnchorGeometries([RANGE, REVISION]);
  const second = viewerGeometry({ read, cache, layout: { id: 'b' } }).readAnchorGeometries([REVISION, RANGE]);
  release([
    { ok: true, ranges: [{ from: 2, to: 6 }], paragraph: 2, hidden: [] },
    { ok: true, ranges: [{ from: 3, to: 4 }], paragraph: 3, hidden: [] },
  ]);
  expect(await first).toMatchObject([{ rects: [{ x: 2 }] }, { rects: [{ x: 3 }] }]);
  expect(await second).toMatchObject([
    { layoutId: 'b', rects: [{ x: 3 }] },
    { layoutId: 'b', rects: [{ x: 2 }] },
  ]);
  expect(calls).toHaveLength(1);
});

test('without a worker read a batch answers like getAnchorGeometry', async () => {
  const geometry = viewerGeometry({});
  const paragraph = { kind: 'paragraph', paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' } } as const;
  const targets = [paragraph, { ...RANGE, version: 'v0' }, PROPOSAL];
  expect(await geometry.readAnchorGeometries(targets)).toEqual(
    targets.map((target) => geometry.getAnchorGeometry(target))
  );
});

test('a batch keeps ranges ending where an unbuilt page starts on the built page', async () => {
  const { read, calls } = counted(() => ({ ok: true, ranges: [{ from: 2, to: 10 }], paragraph: 2, hidden: [] }));
  const geometry = viewerGeometry({ read, pages: NEXT_UNBUILT });
  for (let round = 0; round < 2; round += 1) {
    expect(await geometry.readAnchorGeometries([REVISION, RANGE])).toEqual([
      expect.objectContaining({ ok: true, rects: [expect.objectContaining({ pageIndex: 0, x: 2, width: 8 })], unbuiltPages: [], anchor: expect.objectContaining({ pageIndex: 0, x: 10 }) }),
      expect.objectContaining({ ok: true, unbuiltPages: [], anchor: expect.objectContaining({ pageIndex: 0 }) }),
    ]);
  }
  expect(calls).toHaveLength(1);
});

test('a proposal in a batch lists the unbuilt pages it reaches like the worker targets', async () => {
  const spanning = { ok: true as const, ranges: [{ from: 2, to: 15 }], paragraph: 2 };
  const { read, calls } = counted(() => ({ ...spanning, hidden: [] }));
  const geometry = viewerGeometry({ read, pages: NEXT_UNBUILT, proposals: { p1: spanning } });
  const provisional = expect.objectContaining({
    ok: true,
    rects: [expect.objectContaining({ pageIndex: 0, x: 2, width: 8 })],
    unbuiltPages: [1],
    anchor: expect.objectContaining({ pageIndex: 1 }),
  });
  for (let round = 0; round < 2; round += 1) {
    expect(await geometry.readAnchorGeometries([PROPOSAL, REVISION, RANGE])).toEqual([provisional, provisional, provisional]);
  }
  expect(calls).toEqual([[REVISION, { ...RANGE, version: 'w1' }]]);
  const sync = geometry.getAnchorGeometry(PROPOSAL);
  expect(sync).toMatchObject({ ok: true, anchor: { pageIndex: 0 } });
  expect(sync).not.toHaveProperty('unbuiltPages');
});

test('a header or footer no laid-out page paints is unavailable', async () => {
  const geometry = viewerGeometry({
    pages: [{}, { unbuilt: true, positionSpan: [100, 200], hfParts: { header: 'rIdOther' } }],
    read: reply({ ok: true, root: 'hf:rIdH', ranges: [{ from: 1, to: 4 }], paragraph: 1, hidden: [] }),
  });
  expect(await geometry.readAnchorGeometry(RANGE)).toEqual({
    ok: false,
    failure: { code: 'layout-unavailable', message: 'No laid-out page paints this header or footer' },
  });
});
