/**
 * The one non-body part the editor can have open. A single value rather than a
 * flag per kind, so nothing can claim the caret alongside it.
 */

import type {
  DisplayListImageRegion,
  DisplayListRegionHit,
} from '@betteroffice/docx/layout/render';

/**
 * What the user opened, before the package resolves it: the band names the
 * variant to look up (the first-page one wins on a `titlePg` page 1) and the
 * page it was opened from.
 */
export type PartEditTarget = { kind: 'header' | 'footer'; isFirstPage: boolean; pageIndex: number };

/**
 * The resolved part: a band with the relationship id of its part, null until a
 * newly materialised one registers.
 */
export type PartEdit = { kind: 'header' | 'footer'; rId: string | null };

/** yrs root story the open part types into; the body when none is typeable. */
export function partEditStory(part: PartEdit | null): string {
  if (!part) return 'body';
  return part.rId ? `hf:${part.rId}` : 'body';
}

/** Whether a hit lands in the open part — what separates typing from leaving. */
export function hitBelongsToPart(
  part: PartEdit | null,
  hit: DisplayListRegionHit | null
): boolean {
  return Boolean(part && hit && hit.region === part.kind);
}

/** Display-list image scope of the open part. */
export function partImageRegion(part: PartEdit | null): DisplayListImageRegion {
  return part ? part.kind : 'body';
}
