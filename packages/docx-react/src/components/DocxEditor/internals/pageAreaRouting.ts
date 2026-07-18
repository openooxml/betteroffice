/**
 * Shared page-area routing predicate (G4).
 *
 * Click/context-menu routing that needs to know "did this event land inside the
 * document page area?" must accept BOTH renderers' page hosts. On the DOM-painter
 * path the pages live under `.paged-editor__pages`; on the experimental canvas
 * renderer that painter is parked in a 0×0 stage and the visible pages are a
 * SIBLING `.canvas-pages` host, so a bare `closest('.paged-editor__pages')`
 * returns null for every canvas body click and mis-routes it. Using this single
 * selector list keeps both renderers routing identically.
 */

/** Selector list matching either renderer's visible page-area host. */
export const PAGE_AREA_SELECTOR = '.paged-editor__pages, .canvas-pages';

/**
 * True when `target` is inside the document page area of either renderer
 * (DOM-painter `.paged-editor__pages` or canvas `.canvas-pages`).
 */
export function isWithinPageArea(target: Element | null | undefined): boolean {
  return target?.closest(PAGE_AREA_SELECTOR) != null;
}
