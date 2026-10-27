import { yrsIdToNumericId, type DisplayListRect } from '@betteroffice/docx/layout/render';
import type { DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import { reviewRevisionIds, type ReviewSidebarItem } from '../../sidebar/trackedChanges';

export function workerReviewRects(items: readonly ReviewSidebarItem[], review: Pick<DocxReviewSnapshot, 'revisionGeometry' | 'proposalHighlights'> | null | undefined) {
  const revisions = new Map<number, { order: number; rects: readonly DisplayListRect[] }[]>();
  Object.entries(review?.revisionGeometry ?? {}).forEach(([id, rects], order) => {
    const numeric = yrsIdToNumericId(id);
    const groups = revisions.get(numeric) ?? [];
    groups.push({ order, rects });
    revisions.set(numeric, groups);
  });
  return new Map(items.map(item => {
    const entry = item.review!;
    const rects = entry.proposalId ? review?.proposalHighlights[entry.proposalId] ?? []
      : reviewRevisionIds(entry.change).flatMap(id => revisions.get(id) ?? [])
        .sort((first, second) => first.order - second.order).flatMap(group => group.rects);
    return [item.id, [...rects].sort((first, second) => first.pageIndex - second.pageIndex || first.y - second.y || first.x - second.x)] as const;
  }));
}
