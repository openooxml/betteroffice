/** Page-area routing across DOM and canvas renderers. */

/** Selector list matching either renderer's visible page-area host. */
export const PAGE_AREA_SELECTOR = '.paged-editor__pages, .canvas-pages';

/**
 * True when `target` is inside the document page area of either renderer
 * (DOM-painter `.paged-editor__pages` or canvas `.canvas-pages`).
 */
export function isWithinPageArea(target: Element | null | undefined): boolean {
  return target?.closest(PAGE_AREA_SELECTOR) != null;
}
