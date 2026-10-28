import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import * as display from '@betteroffice/docx/layout/render';
import type { ProposalGeometryTarget } from '@betteroffice/docx/yrs';
import type { DocxGeometryTarget } from '../../../plugins/types';
import { createWorkerPluginGeometry } from './workerPluginGeometry';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE = { width: 300, height: 500 };
const GAP = 20;
const rect = (x: number, y: number, width: number, height: number) =>
  new DOMRect(x, y, width, height);
const restore: { mockRestore(): void }[] = [];
afterEach(() => {
  for (const spy of restore.splice(0)) spy.mockRestore();
});

const DRAWING: DocxGeometryTarget = {
  kind: 'drawing',
  paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' },
  offset: 2,
};
const FRAME = { x: 30, y: 40, w: 50, h: 20, docStart: 10 };

async function shell(painted = true) {
  const container = document.createElement('div');
  const host = document.createElement('div');
  host.setAttribute('data-docx-pages-host', '');
  container.appendChild(host);
  host.getBoundingClientRect = () => rect(0, 0, PAGE.width, 2 * PAGE.height + GAP);
  container.getBoundingClientRect = host.getBoundingClientRect;
  const layer = document.createElement('div');
  layer.getBoundingClientRect = container.getBoundingClientRect;
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  const frameOf = (epoch: number) => ({
    identity,
    coversThroughSeq: 0,
    frameEpoch: epoch,
    layoutRevision: epoch,
    caret: null,
    selectionRects: [],
    pages: [0, 1].map((pageIndex) => ({
      page: { pageId: String(pageIndex), pageIndex },
      transform: [1, 0, 0, 1, 0, 0],
      bounds: { x: 0, y: pageIndex * (PAGE.height + GAP), ...PAGE },
      displayPage: {
        width: PAGE.width,
        height: PAGE.height,
        pageIndex,
        primitives:
          painted && pageIndex === 0
            ? [{ kind: 'image', relId: 'r1', ...FRAME }]
            : [],
        positionSpan: [pageIndex * 100, pageIndex * 100 + 99],
      },
    })),
  });
  const view = {
    identity,
    version: 0,
    throughSeq: 0,
    presentedThroughSeq: 0,
    workerSelection: null,
    anchors: null,
    editState: undefined,
    geometry: frameOf(1),
  };
  const review = { version: 'v1', previewVersion: 0 };
  const reads: { targets: DocxGeometryTarget[]; resolve(): void }[] = [];
  const answerOf = (asked: DocxGeometryTarget): ProposalGeometryTarget =>
    asked.kind === 'drawing'
      ? ({ ok: true, drawing: true, ranges: [{ from: 10, to: 11 }], paragraph: 10 } as ProposalGeometryTarget)
      : ({ ok: true, ranges: [{ from: 10, to: 15 }], paragraph: 10 } as ProposalGeometryTarget);
  const controller = {
    getSnapshot: () => ({ status: 'ready', view }),
    review: async () => review,
    describedReview: () => review,
    readAt: (request: { targets: DocxGeometryTarget[] }) =>
      new Promise((resolve) => {
        reads.push({
          targets: request.targets,
          resolve: () => resolve(request.targets.map(answerOf)),
        });
      }),
    ref: { getZoom: () => 1 },
  };
  restore.push(spyOn(display, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never));
  restore.push(
    spyOn(display, 'createDisplayListQueries').mockImplementation(((list: {
      pages: unknown[];
    }) => ({
      displayList: list,
      pageCount: () => 2,
      pageSize: () => PAGE,
      pageBounds: (pageIndex: number) =>
        pageIndex < 2 ? { pageIndex, x: 0, y: 0, ...PAGE } : null,
      rangeRects: (from: number, to: number) => [
        { pageIndex: 0, x: from, y: 10, width: to - from, height: 12 },
      ],
      anchorRect: () => ({ pageIndex: 0, x: 0, y: 0, width: 0, height: 12 }),
      dispose() {},
    })) as never)
  );
  const plugin = createWorkerPluginGeometry(
    controller as never,
    () => container,
    () => layer,
    () => {}
  );
  await plugin.update();
  return {
    plugin,
    reads,
    async answer() {
      for (const read of reads.splice(0)) read.resolve();
      await new Promise((resolve) => setTimeout(resolve, 0));
    },
  };
}

test('a worker drawing target answers with the drawn frame once its read lands', async () => {
  const s = await shell();
  const geometry = s.plugin.geometry()!;
  expect(geometry.getAnchorGeometry(DRAWING)).toMatchObject({
    ok: false,
    failure: { code: 'pending' },
  });
  await Promise.resolve();
  await s.answer();
  const result = geometry.getAnchorGeometry(DRAWING);
  if (!result.ok) throw new Error(result.failure.message);
  expect(result.rects).toEqual([
    { pageIndex: 0, x: FRAME.x, y: FRAME.y, width: FRAME.w, height: FRAME.h },
  ]);
  expect(result.anchor).toEqual(result.rects[0]);
});

test('worker readAnchorGeometries answers a drawing target with its frame', async () => {
  const s = await shell();
  const read = s.plugin.geometry()!.readAnchorGeometries([DRAWING]);
  await Promise.resolve();
  await s.answer();
  const [result] = await read;
  expect(result).toMatchObject({
    ok: true,
    rects: [{ pageIndex: 0, x: FRAME.x, y: FRAME.y, width: FRAME.w, height: FRAME.h }],
  });
});

test('a worker drawing that is not painted is missing once the layout is complete', async () => {
  const s = await shell(false);
  const read = s.plugin.geometry()!.readAnchorGeometries([DRAWING]);
  await Promise.resolve();
  await s.answer();
  const [result] = await read;
  expect(result).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
});
