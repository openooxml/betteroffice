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

/**
 * The page canvases a renderer mounts under its pages host, so page lookups
 * read them instead of searching a subtree that also holds every page's
 * accessibility mirror. A renderer adds each `<canvas data-page-index>` it
 * mounts, deletes it on unmount, invalidates the registry whenever it may have
 * moved or renumbered them, and binds the registry to the host with
 * {@link bindDisplayPageRegistry}.
 */
export class DisplayPageRegistry {
  private readonly canvases = new Set<HTMLCanvasElement>();
  private ordered: HTMLCanvasElement[] | null = null;
  private byIndex: Map<number, HTMLCanvasElement> | null = null;
  private materializer: ((pageIndices: readonly number[]) => void) | null = null;

  /**
   * Lets a renderer that keeps page DOM only near the viewport build the
   * DOM of other pages when something needs it; see {@link materializeDisplayPages}.
   */
  setMaterializer(materializer: ((pageIndices: readonly number[]) => void) | null): void {
    this.materializer = materializer;
  }

  /** Builds the DOM of `pageIndices` now, where the renderer keeps it only for some pages. */
  materialize(pageIndices: readonly number[]): void {
    this.materializer?.(pageIndices);
  }

  add(canvas: HTMLCanvasElement): void {
    this.canvases.add(canvas);
    this.invalidate();
  }

  delete(canvas: HTMLCanvasElement): void {
    this.canvases.delete(canvas);
    this.invalidate();
  }

  /** Forgets the order and indices read so far, after the renderer moved or renumbered canvases. */
  invalidate(): void {
    this.ordered = null;
    this.byIndex = null;
  }

  /** The registered canvases in document order. */
  all(): HTMLCanvasElement[] {
    this.ordered ??= [...this.canvases].sort((a, b) =>
      a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_PRECEDING ? 1 : -1
    );
    return this.ordered;
  }

  /** The first canvas in document order that shows `pageIndex`. */
  canvas(pageIndex: number): HTMLCanvasElement | null {
    const cached = this.byIndex?.get(pageIndex);
    if (cached?.dataset.pageIndex === String(pageIndex)) return cached;
    this.byIndex = this.indexPages();
    return this.byIndex.get(pageIndex) ?? null;
  }

  private indexPages(): Map<number, HTMLCanvasElement> {
    const byIndex = new Map<number, HTMLCanvasElement>();
    for (const canvas of this.all()) {
      const pageIndex = canvas.dataset.pageIndex;
      if (pageIndex === undefined) continue;
      const key = Number(pageIndex);
      if (String(key) === pageIndex && !byIndex.has(key)) byIndex.set(key, canvas);
    }
    return byIndex;
  }
}

const displayPageRegistries = new WeakMap<HTMLElement, DisplayPageRegistry>();

/** Binds `registry` to the pages `host` it renders into, or unbinds it with null. */
export function bindDisplayPageRegistry(
  host: HTMLElement,
  registry: DisplayPageRegistry | null
): void {
  if (registry) displayPageRegistries.set(host, registry);
  else displayPageRegistries.delete(host);
}

/** Builds the accessibility DOM of `pageIndices` under `host` before a DOM query reads it. */
export function materializeDisplayPages(host: HTMLElement, pageIndices: readonly number[]): void {
  if (pageIndices.length > 0) displayPageRegistries.get(host)?.materialize(pageIndices);
}

/** The first `<canvas data-page-index>` under `host` for `pageIndex`. */
export function displayPageCanvas(host: HTMLElement, pageIndex: number): HTMLCanvasElement | null {
  const registry = displayPageRegistries.get(host);
  if (registry) return registry.canvas(pageIndex);
  return host.querySelector<HTMLCanvasElement>(`canvas[data-page-index="${pageIndex}"]`);
}

/** The canvas page elements inside a pages host, in document order. */
export function displayPageCanvases(host: HTMLElement): HTMLCanvasElement[] {
  const registry = displayPageRegistries.get(host);
  if (registry) return registry.all();
  return Array.from(host.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]'));
}

/** Resolves a display-list page rectangle from canvas or DOM page hosts. */
export function resolveDisplayPageClientRect(
  host: HTMLElement,
  queries: DisplayListQueries,
  pageIndex: number,
  options?: DisplayPageHostOptions
): DisplayPageClientRect | null {
  const canvas = displayPageCanvas(host, pageIndex);
  if (canvas) {
    const rect = canvas.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) return rect;
  }
  const pageHost = host.querySelector<HTMLElement>(`.canvas-page[data-page-index="${pageIndex}"]`);
  if (pageHost) {
    const rect = pageHost.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) return rect;
  }
  if (host.classList.contains('canvas-pages')) return null;

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
  const canvases = displayPageCanvases(host);
  const canvasByPage = new Map<number, HTMLCanvasElement>();
  for (const canvas of canvases) {
    const pageIndex = Number(canvas.dataset.pageIndex);
    if (Number.isFinite(pageIndex)) canvasByPage.set(pageIndex, canvas);
  }

  let chosen: { pageIndex: number; rect: DisplayPageClientRect } | null = null;
  let nearest: { pageIndex: number; rect: DisplayPageClientRect; dist: number } | null = null;

  for (let pageIndex = 0; pageIndex < queries.pageCount(); pageIndex++) {
    const canvas = canvasByPage.get(pageIndex);
    const canvasRect = canvas?.getBoundingClientRect();
    const rect = canvasRect && canvasRect.width > 0 && canvasRect.height > 0
      ? canvasRect
      : resolveDisplayPageClientRect(host, queries, pageIndex, options);
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
