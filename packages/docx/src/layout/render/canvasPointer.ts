/**
 * Client-coordinate → display-list resolution for the canvas renderer's
 * pointer path. Framework-free: both the React and the Vue adapter route
 * their pointer events through this resolver.
 *
 * The pointer hooks map the client point onto one of the
 * `<canvas data-page-index>` pages, convert it to page-local px (the
 * display list's unit space), and ask the Rust `hit_test_regions` query for
 * the region + doc position.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { DisplayListQueries, DisplayListRegionHit } from './displayListQueries';

export interface CanvasPointHit {
  pageIndex: number;
  /** page-local px, the display list's unit space */
  x: number;
  y: number;
  /** region-aware hit from the Rust query (null when the engine isn't ready) */
  hit: DisplayListRegionHit | null;
}

export interface DisplayPageHostOptions {
  /** Vertical space between page shells when the host has no live canvases. */
  pageGap?: number;
  /** Space before the first page when the host has no live canvases. */
  paddingTop?: number;
}

export interface DisplayPageClientRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

/** the canvas page elements inside a `.canvas-pages` host, in page order */
function pageCanvases(host: HTMLElement): HTMLCanvasElement[] {
  return Array.from(host.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]'));
}

/**
 * Resolve the client-space shell for one display-list page. A mounted canvas
 * is the preferred projection host. During the Phase-2 transition the DOM
 * painter may still be the visible renderer, so its *container* is used as a
 * page-stack projection host without inspecting any painted page/text/table
 * node. Page sizes remain display-list-owned.
 */
export function resolveDisplayPageClientRect(
  host: HTMLElement,
  queries: DisplayListQueries,
  pageIndex: number,
  options?: DisplayPageHostOptions
): DisplayPageClientRect | null {
  const canvas = host.querySelector<HTMLCanvasElement>(`canvas[data-page-index="${pageIndex}"]`);
  if (canvas) {
    const rect = canvas.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) return rect;
  }

  const size = queries.pageSize(pageIndex);
  if (!size || size.width <= 0 || size.height <= 0) return null;
  const hostRect = host.getBoundingClientRect();
  if (hostRect.width <= 0) return null;

  // The painter pages container is transformed as one centered column. Its
  // live host rect is therefore a coordinate projection only; all document
  // geometry and page heights come from the immutable display list.
  const unscaledWidth = host.offsetWidth || host.clientWidth || hostRect.width;
  const scale = unscaledWidth > 0 ? hostRect.width / unscaledWidth : 1;
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const pageGap = options?.pageGap ?? 24;
  // The pages host applies `padding: pageGap` to its page-stack container.
  // The canvas-free fallback projects through that container, so its first
  // page starts after the same leading inset. Callers with an unpadded custom
  // host can opt out explicitly with `paddingTop: 0`.
  const paddingTop = options?.paddingTop ?? pageGap;
  let logicalTop = paddingTop;
  for (let i = 0; i < pageIndex; i++) {
    logicalTop += (queries.pageSize(i)?.height ?? 0) + pageGap;
  }
  const width = size.width * safeScale;
  const height = size.height * safeScale;
  const left = hostRect.left + (hostRect.width - width) / 2;
  const top = hostRect.top + logicalTop * safeScale;
  return { left, top, right: left + width, bottom: top + height, width, height };
}

/**
 * Resolve a client point against the canvas pages.
 *
 * Containment picks the page; with `clampToNearestPage` (drag-selection) a
 * point outside every page snaps to the vertically nearest one and clamps
 * into its bounds, so drags keep extending past page edges — the analogue of
 * the DOM path's nearest-span snapping. Coordinates convert through the
 * page's CSS rect vs its display-list size, so any ancestor scale transform
 * is factored out.
 */
export function resolveCanvasPoint(
  host: HTMLElement,
  queries: DisplayListQueries,
  clientX: number,
  clientY: number,
  options?: { clampToNearestPage?: boolean; pageGap?: number; paddingTop?: number }
): CanvasPointHit | null {
  const canvases = pageCanvases(host);
  const canvasByPage = new Map<number, HTMLCanvasElement>();
  for (const canvas of canvases) {
    const pageIndex = Number(canvas.dataset.pageIndex);
    if (Number.isFinite(pageIndex)) canvasByPage.set(pageIndex, canvas);
  }

  let chosen: { pageIndex: number; rect: DisplayPageClientRect } | null = null;
  let nearest: { pageIndex: number; rect: DisplayPageClientRect; dist: number } | null = null;

  for (let pageIndex = 0; pageIndex < queries.pageCount(); pageIndex++) {
    const canvas = canvasByPage.get(pageIndex);
    const rect =
      canvas?.getBoundingClientRect() ??
      resolveDisplayPageClientRect(host, queries, pageIndex, options);
    if (!rect) continue;
    if (rect.width <= 0 || rect.height <= 0) continue;
    if (
      clientX >= rect.left &&
      clientX <= rect.right &&
      clientY >= rect.top &&
      clientY <= rect.bottom
    ) {
      chosen = { pageIndex, rect };
      break;
    }
    const dy = clientY < rect.top ? rect.top - clientY : Math.max(0, clientY - rect.bottom);
    const dx = clientX < rect.left ? rect.left - clientX : Math.max(0, clientX - rect.right);
    const dist = dy * 4 + dx; // favor the vertically closest page mid-drag
    if (!nearest || dist < nearest.dist) nearest = { pageIndex, rect, dist };
  }

  if (!chosen) {
    if (!options?.clampToNearestPage || !nearest) return null;
    chosen = { pageIndex: nearest.pageIndex, rect: nearest.rect };
  }

  const pageIndex = chosen.pageIndex;
  const size = queries.pageSize(pageIndex);
  if (!size) return null;

  const scaleX = size.width / chosen.rect.width;
  const scaleY = size.height / chosen.rect.height;
  const x = Math.min(Math.max((clientX - chosen.rect.left) * scaleX, 0), size.width);
  const y = Math.min(Math.max((clientY - chosen.rect.top) * scaleY, 0), size.height);

  return { pageIndex, x, y, hit: queries.hitTestRegions(pageIndex, x, y) };
}
