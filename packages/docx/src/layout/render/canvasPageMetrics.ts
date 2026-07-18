/**
 * Layout constants + vertical page math for the experimental canvas pages
 * hosts (React `CanvasPagesView` / Vue `CanvasPagesView.vue`): a flex column
 * of one <canvas> per page. Shared by the views themselves and by the
 * display-list consumers that need scroll-container Y coordinates (sidebar
 * anchor positions) so the two can never drift.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { DisplayListQueries } from './displayListQueries';

/** vertical gap between canvas pages (flex `gap`) */
export const CANVAS_PAGE_GAP_PX = 16;
/** padding around the canvas pages column */
export const CANVAS_PAGES_PADDING_PX = 24;

/**
 * Top Y of each canvas page inside the `.canvas-pages` host, derived from the
 * display list's own page heights. `tops[i] + pageLocalY` converts a
 * display-list rect into host coordinates.
 */
export function canvasPageTops(queries: DisplayListQueries): number[] {
  const tops: number[] = [];
  let y = CANVAS_PAGES_PADDING_PX;
  for (let i = 0; i < queries.pageCount(); i++) {
    tops.push(y);
    y += (queries.pageSize(i)?.height ?? 0) + CANVAS_PAGE_GAP_PX;
  }
  return tops;
}
