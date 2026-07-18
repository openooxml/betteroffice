/**
 * Project a page-local (px) point / rect on a canvas page into the overlay
 * portal target's coordinate space, via the live per-page `<canvas>` rect.
 *
 * This is the same client→canvas conversion `CanvasSelectionOverlay` and
 * `useFloatingCommentBtn` use, extracted so the canvas interaction-handle
 * overlays (image resize, table resize) share one implementation. The live
 * canvas rect already reflects the page column's centering, the sidebar-open
 * `translateX`, and any zoom, so a projected point lands on the painted glyph /
 * image / cell edge in every case; `scaleX`/`scaleY` are 1 while the canvas
 * paints at logical px (CSS size == display-list page size) and stay correct if
 * a future zoom scales the canvas.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import {
  resolveDisplayPageClientRect,
  type DisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';

export interface CanvasProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
  scaleX: number;
  scaleY: number;
}

/** page-local px per one overlay-target px on `pageIndex`, or null if unresolved. */
export function canvasPageScale(
  host: HTMLElement,
  queries: DisplayListQueries,
  pageIndex: number
): { canvasRect: DisplayPageClientRect; scaleX: number; scaleY: number } | null {
  const size = queries.pageSize(pageIndex);
  const canvasRect = resolveDisplayPageClientRect(host, queries, pageIndex);
  if (!canvasRect || !size) return null;
  const scaleX = size.width > 0 ? canvasRect.width / size.width : 1;
  const scaleY = size.height > 0 ? canvasRect.height / size.height : 1;
  return { canvasRect, scaleX, scaleY };
}

/**
 * Project a page-local rect (`x`/`y`/`w`/`h`, px on `pageIndex`) into
 * `overlayTarget` coordinates. Returns null when the page's canvas isn't
 * mounted yet (mid-rebuild) so callers can skip a frame.
 */
export function projectPageLocalRect(
  host: HTMLElement,
  overlayTarget: HTMLElement,
  queries: DisplayListQueries,
  pageIndex: number,
  x: number,
  y: number,
  w: number,
  h: number
): CanvasProjectedRect | null {
  const scale = canvasPageScale(host, queries, pageIndex);
  if (!scale) return null;
  const targetRect = overlayTarget.getBoundingClientRect();
  const { canvasRect, scaleX, scaleY } = scale;
  return {
    left: canvasRect.left - targetRect.left + x * scaleX,
    top: canvasRect.top - targetRect.top + y * scaleY,
    width: w * scaleX,
    height: h * scaleY,
    scaleX,
    scaleY,
  };
}
