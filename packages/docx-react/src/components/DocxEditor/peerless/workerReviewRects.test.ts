import { expect, test } from 'bun:test';
import { yrsIdToNumericId, type DisplayListRect } from '@betteroffice/docx/layout/render';
import type { DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import { reviewRevisionIds, type ReviewSidebarItem } from '../../sidebar/trackedChanges';
import { workerReviewRects } from './workerReviewRects';

const rect = (pageIndex: number, y: number, x = 10, width = 20): DisplayListRect => ({ pageIndex, x, y, width, height: 12 });
const item = (id: string, revisionId: number, extra: Partial<NonNullable<ReviewSidebarItem['review']>['change']> = {}, proposalId?: string): ReviewSidebarItem => ({
  id, anchorPos: 0, render: () => null,
  review: { kind: proposalId ? 'agent' : 'human', proposalId,
    change: { revisionId, type: 'insertion', author: 'Editor', date: '', text: 'Typed', from: 1, to: 6, ...extra } },
});

test('indexed review rects match the scan algorithm for grouped revisions, ties and proposals', () => {
  const revisionGeometry = { '9': [rect(1, 10)], '7': [rect(0, 40), rect(0, 20)], '8': [rect(0, 20, 10, 30)],
    '007': [rect(0, 20, 10, 40)], 'unrelated': [rect(0, 1)] };
  const proposalHighlights = { proposal: [rect(2, 30), rect(1, 20), rect(1, 20, 5)] };
  const items = [item('replacement', 9, { type: 'replacement', insertionRevisionId: 8, coalescedRevisionIds: [7, 8, 9] }),
    item('coalesced', 7, { coalescedRevisionIds: [9] }), item('proposal', 7, {}, 'proposal'),
    item('missing-proposal', 7, {}, 'missing'), item('missing-revision', 100)];
  const old = new Map(items.map(item => {
    const entry = item.review!;
    const ids = new Set(reviewRevisionIds(entry.change));
    const rects = entry.proposalId ? proposalHighlights[entry.proposalId as keyof typeof proposalHighlights] ?? []
      : Object.entries(revisionGeometry).filter(([id]) => ids.has(yrsIdToNumericId(id))).flatMap(([, rects]) => rects);
    return [item.id, [...rects].sort((first, second) => first.pageIndex - second.pageIndex || first.y - second.y || first.x - second.x)] as const;
  }));
  const before = structuredClone({ revisionGeometry, proposalHighlights });
  let scans = 0;
  const indexedGeometry = new Proxy(revisionGeometry, { ownKeys(target) { scans++; return Reflect.ownKeys(target); } });
  expect(workerReviewRects(items, { revisionGeometry: indexedGeometry, proposalHighlights } satisfies Pick<DocxReviewSnapshot, 'revisionGeometry' | 'proposalHighlights'>)).toEqual(old);
  expect(scans).toBe(1);
  expect({ revisionGeometry, proposalHighlights }).toEqual(before);
  expect(workerReviewRects(items, undefined).get('replacement')).toEqual([]);
});
