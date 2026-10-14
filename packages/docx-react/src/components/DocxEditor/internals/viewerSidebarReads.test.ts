import { expect, test } from 'bun:test';
import type { DocxSidebarRead, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DisplayListQueries, TrackedChangesResult } from '@betteroffice/docx/layout/render';
import { presentedWorkerVersion, stampWorkerFrameVersion } from './layoutProvenance';
import { EMPTY_TRACKED_CHANGES_RESULT, ViewerSidebarReads, viewerCommentRanges, viewerSidebarPositions } from './viewerSidebarReads';

const point = (position: number, hfRid?: string) => ({ story: 'body', position, ...(hfRid ? { hfRid } : {}) });
const value: DocxSidebarRead = {
  comments: [{ id: '7', anchors: [{ start: point(5), end: point(15) }] }],
  revisions: [{ key: 'revision-11', start: point(20) }],
  trackedChanges: {
    entries: [{ type: 'insertion', text: 'Added', author: 'Writer', revisionId: 11, from: 20, to: 25 }],
    commentToRevision: [[7, 11]],
  },
};

function queries(version: string, pages = 1): DisplayListQueries {
  const result = {
    displayList: {},
    pageCount: () => pages,
    pageSize: () => ({ width: 600, height: 800 }),
    anchorRect: (position: number) => position >= 20 && pages === 1 ? null : ({
      pageIndex: position >= 20 ? 1 : 0, x: 0, y: position, width: 1, height: 10,
    }),
  } as unknown as DisplayListQueries;
  stampWorkerFrameVersion(result, version);
  return result;
}

test('tracked changes are delivered once per version with a rebuilt map', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: 'expectVersion' in request ? request.expectVersion : 'A', value };
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerSidebarReads(read);
  const received: TrackedChangesResult[] = [];
  const deliver = (result: TrackedChangesResult) => { received.push(result); };
  let version = 'A';
  const first = await reads.sidebar(version, ['7'], () => version);
  reads.deliver(version, first!, deliver);
  reads.deliver(version, first!, deliver);
  const again = await reads.sidebar(version, ['7'], () => version);
  expect(again).toBe(first);
  expect(received).toHaveLength(1);
  expect(received[0]!.commentToRevision).toEqual(new Map([[7, 11]]));
  version = 'B';
  const second = await reads.sidebar(version, ['7'], () => version);
  reads.deliver(version, second!, deliver);
  expect(received).toHaveLength(2);
  expect(received[1]).not.toBe(received[0]);
  expect(requests).toHaveLength(2);
});

test('more pages at the same version reproject anchors without another read', async () => {
  let count = 0;
  const read = (async () => { count += 1; return { version: 'A', value }; }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerSidebarReads(read);
  let shown = queries('A');
  const current = () => presentedWorkerVersion(shown);
  const first = await reads.sidebar('A', ['7'], current);
  const emitted = [viewerSidebarPositions(first!, shown)];
  shown = queries('A', 2);
  const second = await reads.sidebar('A', ['7'], current);
  emitted.push(viewerSidebarPositions(second!, shown));
  expect(emitted).toEqual([
    new Map([['comment-7', 29]]),
    new Map([['comment-7', 29], ['revision-11', 860]]),
  ]);
  expect(count).toBe(1);
});

test('a superseded reply is dropped and does not deliver tracked changes', async () => {
  let complete!: (reply: { version: string; value: DocxSidebarRead }) => void;
  const read = (() => new Promise<unknown>((resolve) => { complete = resolve; })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerSidebarReads(read);
  let shown = queries('A');
  const pending = reads.sidebar('A', ['7'], () => presentedWorkerVersion(shown));
  shown = queries('B');
  complete({ version: 'A', value });
  expect(await pending).toBeNull();
});

test('comment ids retain request order and a changed list reads again', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: 'A', value };
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerSidebarReads(read);
  await reads.sidebar('A', ['7', '8'], () => 'A');
  await reads.sidebar('A', ['8', '7'], () => 'A');
  expect(requests).toEqual([
    { kind: 'sidebar', commentIds: ['7', '8'], expectVersion: 'A' },
    { kind: 'sidebar', commentIds: ['8', '7'], expectVersion: 'A' },
  ]);
});

test('empty sidebars deliver the shared empty tracked result', () => {
  const reads = new ViewerSidebarReads((async () => ({ version: 'A', value: null })) as ResidentEngineWorkerClient['documentRead']);
  let result = null as TrackedChangesResult | null;
  reads.deliver('A', { comments: [], revisions: [], trackedChanges: { entries: [], commentToRevision: [] } },
    (next) => { result = next; });
  expect(result).toBe(EMPTY_TRACKED_CHANGES_RESULT);
});

test('comment highlights use the outermost points and skip missing ends or header starts', () => {
  expect(viewerCommentRanges({ ...value, comments: [
    { id: '7', anchors: [{ start: point(9), end: point(10) }, { start: point(3), end: point(30) }] },
    { id: '8', anchors: [{ start: point(1), end: null }] },
    { id: '9', anchors: [{ start: null, end: point(10) }] },
    { id: '10', anchors: [{ start: point(1, 'h'), end: point(10, 'h') }] },
    { id: '11', anchors: [] },
  ] })).toEqual(new Map([
    [7, { from: 3, to: 30 }], [8, null], [9, null], [10, null], [11, null],
  ]));
});
