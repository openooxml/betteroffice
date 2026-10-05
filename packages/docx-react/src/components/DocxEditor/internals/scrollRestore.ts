import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
  type DisplayListVisualLine,
} from '@betteroffice/docx/layout/render';
import type { YrsStickyPosition } from '@betteroffice/docx/yrs';
import {
  computeViewportAnchoredScrollTop,
  type ViewportAnchorSnapshot,
} from './viewportAnchoring';
import { scrollViewport } from './viewportBand';

/** Client-pixel bounds of the band a scroller shows. */
type ViewportBounds = { top: number; bottom: number };

export interface DisplayListScrollAnchor {
  pmPos: number;
  clientOffset: number | null;
  /** Page the caret line sat on; a change means it reflowed across a break. */
  pageIndex: number | null;
  scrollTopSnapshot: number;
}

interface PositionViewportTarget {
  kind: 'position';
  position: YrsStickyPosition;
}

interface PageViewportTarget {
  kind: 'page';
  pageIndex: number;
  pageY: number;
}

export interface DisplayListViewportAnchor extends ViewportAnchorSnapshot {
  target: PositionViewportTarget | PageViewportTarget | null;
}

interface PageProjection {
  top: number;
  scaleY: number;
}

export type CaptureViewportPosition = (displayPosition: number) => YrsStickyPosition | null;

export type ResolveViewportPosition = (position: YrsStickyPosition) => number | null;

/** Candidate anchor lines tried before falling back to a page target. */
const ANCHOR_CANDIDATE_LIMIT = 8;

interface LayoutScrollCompensation {
  from: number;
  to: number;
  scrollTopSnapshot: number;
  sequence: number;
}

let layoutScrollCompensationSequence = 0;
const layoutScrollCompensations = new WeakMap<Element, LayoutScrollCompensation>();

export function layoutScrollCompensation(
  el: Element
): LayoutScrollCompensation | undefined {
  return layoutScrollCompensations.get(el);
}

function setLayoutScrollTop(scrollParent: HTMLElement, top: number, scrollTopSnapshot: number): void {
  const from = scrollParent.scrollTop;
  scrollParent.scrollTop = top;
  layoutScrollCompensations.set(scrollParent, {
    from,
    to: scrollParent.scrollTop,
    scrollTopSnapshot,
    sequence: ++layoutScrollCompensationSequence,
  });
}

function pageProjection(
  queries: DisplayListQueries,
  host: HTMLElement,
  pageIndex: number,
  cache?: Map<number, PageProjection | null>
): PageProjection | null {
  if (cache?.has(pageIndex)) return cache.get(pageIndex) ?? null;
  const pageRect = resolveDisplayPageClientRect(host, queries, pageIndex);
  const pageSize = queries.pageSize(pageIndex);
  const projection =
    pageRect && pageSize && pageSize.height > 0
      ? { top: pageRect.top, scaleY: pageRect.height / pageSize.height }
      : null;
  cache?.set(pageIndex, projection);
  return projection;
}

function projectedRectClientY(
  queries: DisplayListQueries,
  host: HTMLElement,
  rect: DisplayListRect,
  cache?: Map<number, PageProjection | null>
): { top: number; bottom: number } | null {
  const projection = pageProjection(queries, host, rect.pageIndex, cache);
  if (!projection) return null;
  const top = projection.top + rect.y * projection.scaleY;
  return { top, bottom: top + rect.height * projection.scaleY };
}

function projectedAnchorRect(
  queries: DisplayListQueries,
  host: HTMLElement,
  pmPos: number
): { clientY: number; pageIndex: number } | null {
  const rect = queries.anchorRect(pmPos);
  if (!rect) return null;
  const projected = projectedRectClientY(queries, host, rect);
  return projected ? { clientY: projected.top, pageIndex: rect.pageIndex } : null;
}

/**
 * Narrowest visual line covering `position`. Keyed on the document position
 * rather than a `paraId`: the resident engine stamps no `paraId` on its
 * primitives, so a paraId-filtered lookup never resolves on the canvas path.
 */
function lineAtPosition(
  lines: readonly DisplayListVisualLine[],
  position: number
): DisplayListVisualLine | null {
  let best: DisplayListVisualLine | null = null;
  for (const line of lines) {
    if (position < line.from || position > line.to) continue;
    if (!best || line.to - line.from < best.to - best.from) best = line;
  }
  return best;
}

function viewportTargetClientY(
  anchor: DisplayListViewportAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  resolvePosition: ResolveViewportPosition
): number | null {
  const target = anchor.target;
  if (!target) return null;
  if (target.kind === 'position') {
    const position = resolvePosition(target.position);
    if (position == null) return null;
    const line = lineAtPosition(queries.visualLines(), position);
    return line ? (projectedRectClientY(queries, host, line)?.top ?? null) : null;
  }
  const pageRect = resolveDisplayPageClientRect(host, queries, target.pageIndex);
  const pageSize = queries.pageSize(target.pageIndex);
  if (!pageRect || !pageSize || pageSize.height <= 0) return null;
  return pageRect.top + target.pageY * (pageRect.height / pageSize.height);
}

/**
 * Anchor the viewport to the content nearest its top edge: the topmost visible
 * line, or the closest line above/below when the top edge sits in a page gap or
 * margin. Line geometry follows the content across page boundaries, which a page
 * index cannot. Candidates are tried in order because a line start need not be
 * a position the sticky projection can round-trip.
 */
function nearestLineAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: ViewportBounds,
  lines: Iterable<DisplayListVisualLine>,
  capturePosition: CaptureViewportPosition,
  visibleOnly = false
): { target: PositionViewportTarget; clientY: number } | null {
  const visible: Array<{ line: DisplayListVisualLine; clientY: number }> = [];
  let nearest: { line: DisplayListVisualLine; clientY: number; distance: number } | null = null;
  const projectionCache = new Map<number, PageProjection | null>();
  for (const line of lines) {
    // Lines arrive in page order: once the viewport has candidates, nothing a
    // page past them can win, so long documents stop scanning early.
    if (visible.length > 0 && line.pageIndex > visible[0].line.pageIndex + 1) break;
    const projected = projectedRectClientY(queries, host, line, projectionCache);
    if (!projected) continue;
    if (projected.bottom >= viewport.top && projected.top <= viewport.bottom) {
      visible.push({ line, clientY: projected.top });
      continue;
    }
    const distance =
      projected.bottom < viewport.top
        ? viewport.top - projected.bottom
        : projected.top - viewport.bottom;
    if (!nearest || distance < nearest.distance) {
      nearest = { line, clientY: projected.top, distance };
    }
  }
  const candidates = visible.sort((left, right) => left.clientY - right.clientY);
  if (candidates.length === 0 && nearest && !visibleOnly) candidates.push(nearest);
  for (const candidate of candidates.slice(0, ANCHOR_CANDIDATE_LIMIT)) {
    const position = capturePosition(candidate.line.from);
    if (position) {
      return { target: { kind: 'position', position }, clientY: candidate.clientY };
    }
  }
  return null;
}

/** The first page, in their top-to-bottom stacking, whose client rect reaches `top`. */
function firstPageReaching(
  queries: DisplayListQueries,
  host: HTMLElement,
  top: number
): number | null {
  let low = 0;
  let high = queries.pageCount() - 1;
  let found: number | null = null;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const rect = resolveDisplayPageClientRect(host, queries, middle);
    if (!rect) return null;
    if (rect.bottom >= top) {
      found = middle;
      high = middle - 1;
    } else {
      low = middle + 1;
    }
  }
  return found;
}

/**
 * A page-level filter only has to keep every page the exact per-line test can
 * accept, so its comparisons allow this much slack against rounding.
 */
const PAGE_FILTER_SLACK = 1;

/**
 * Every page with a line that can reach the viewport, in page order: the pages
 * the viewport spans, and any other page whose lines come near or beyond its
 * own bounds and near the viewport. Null when a page rect cannot be resolved.
 */
function pagesReachingViewport(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: ViewportBounds,
  projectionCache: Map<number, PageProjection | null>
): number[] | null {
  const pageCount = queries.pageCount();
  const first = firstPageReaching(queries, host, viewport.top);
  if (first === null) return null;
  const last = firstPageReaching(queries, host, viewport.bottom) ?? pageCount - 1;
  const pages: number[] = [];
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    if (pageIndex >= first && pageIndex <= last) {
      pages.push(pageIndex);
      continue;
    }
    const extent = queries.visualLineExtent(pageIndex);
    const size = queries.pageSize(pageIndex);
    if (
      !extent ||
      (size && extent.top >= PAGE_FILTER_SLACK && extent.bottom <= size.height - PAGE_FILTER_SLACK)
    ) {
      continue;
    }
    const projection = pageProjection(queries, host, pageIndex, projectionCache);
    if (!projection) return null;
    const top = projection.top + extent.top * projection.scaleY;
    const bottom = projection.top + extent.bottom * projection.scaleY;
    if (bottom + PAGE_FILTER_SLACK >= viewport.top && top - PAGE_FILTER_SLACK <= viewport.bottom) {
      pages.push(pageIndex);
    }
  }
  return pages;
}

function* linesOnPages(
  queries: DisplayListQueries,
  pages: readonly number[]
): Generator<DisplayListVisualLine> {
  for (const pageIndex of pages) yield* queries.visualLinesOnPage(pageIndex);
}

/**
 * Last resort for a viewport with no projectable text line (empty or image-only
 * pages). A page index cannot track content across a page-count change, so this
 * only runs when there is no line to anchor to at all.
 */
function visiblePageAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: ViewportBounds
): { target: PageViewportTarget; clientY: number } | null {
  for (let pageIndex = 0; pageIndex < queries.pageCount(); pageIndex += 1) {
    const pageRect = resolveDisplayPageClientRect(host, queries, pageIndex);
    const pageSize = queries.pageSize(pageIndex);
    if (
      !pageRect ||
      !pageSize ||
      pageSize.height <= 0 ||
      pageRect.bottom < viewport.top ||
      pageRect.top > viewport.bottom
    ) {
      continue;
    }
    const scaleY = pageRect.height / pageSize.height;
    const pageY = Math.min(
      pageSize.height,
      Math.max(0, (Math.max(viewport.top, pageRect.top) - pageRect.top) / scaleY)
    );
    return {
      target: { kind: 'page', pageIndex, pageY },
      clientY: pageRect.top + pageY * scaleY,
    };
  }
  return null;
}

export function captureDisplayListScrollAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  pmPos: number
): DisplayListScrollAnchor {
  if (!scrollParent.style.overflowAnchor) {
    scrollParent.style.setProperty('overflow-anchor', 'none');
  }
  const projected = projectedAnchorRect(queries, host, pmPos);
  const viewport = scrollViewport(scrollParent);
  return {
    pmPos,
    clientOffset: projected ? (projected.clientY - viewport.top) / viewport.zoom : null,
    pageIndex: projected?.pageIndex ?? null,
    scrollTopSnapshot: scrollParent.scrollTop,
  };
}

export function captureDisplayListViewportAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  capturePosition: CaptureViewportPosition
): DisplayListViewportAnchor {
  if (!scrollParent.style.overflowAnchor) {
    scrollParent.style.setProperty('overflow-anchor', 'none');
  }
  const viewport = scrollViewport(scrollParent);
  // Scanning the pages whose lines can reach the viewport finds the visible
  // lines a scan of every line would. The full scan remains for a viewport
  // showing no line, where the nearest line anywhere wins.
  const pages = pagesReachingViewport(queries, host, viewport, new Map());
  const resolved =
    (pages &&
      nearestLineAnchor(
        queries,
        host,
        viewport,
        linesOnPages(queries, pages),
        capturePosition,
        true
      )) ??
    nearestLineAnchor(queries, host, viewport, queries.visualLines(), capturePosition) ??
    visiblePageAnchor(queries, host, viewport);
  return {
    target: resolved?.target ?? null,
    viewportOffset: resolved ? (resolved.clientY - viewport.top) / viewport.zoom : 0,
    scrollTopSnapshot: scrollParent.scrollTop,
  };
}

/**
 * Pin the local caret's line back to the offset it held before the pass.
 *
 * Two cases deliberately do not pin: an anchor that no longer projects, and one
 * whose line reflowed onto another page. Pinning either would drag the viewport
 * by a whole page break even though nothing above the caret moved, so both hold
 * the captured scrollTop and leave the single corrective move to the
 * caret-into-view step.
 */
export function restoreDisplayListScrollAnchor(
  anchor: DisplayListScrollAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement
): void {
  const projected = projectedAnchorRect(queries, host, anchor.pmPos);
  const pinned =
    projected != null &&
    anchor.clientOffset != null &&
    (anchor.pageIndex == null || anchor.pageIndex === projected.pageIndex)
      ? projected
      : null;
  const viewport = scrollViewport(scrollParent);
  const nextTargetTop = pinned
    ? scrollParent.scrollTop + pinned.clientY / viewport.zoom - viewport.top / viewport.zoom
    : null;
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  setLayoutScrollTop(
    scrollParent,
    computeViewportAnchoredScrollTop(
      { viewportOffset: anchor.clientOffset ?? 0, scrollTopSnapshot: anchor.scrollTopSnapshot },
      nextTargetTop,
      maxScroll
    ),
    anchor.scrollTopSnapshot
  );
}

export function restoreDisplayListViewportAnchor(
  anchor: DisplayListViewportAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  resolvePosition: ResolveViewportPosition
): void {
  const clientY = viewportTargetClientY(anchor, queries, host, resolvePosition);
  const viewport = scrollViewport(scrollParent);
  const nextTargetTop =
    clientY == null
      ? null
      : scrollParent.scrollTop + clientY / viewport.zoom - viewport.top / viewport.zoom;
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  setLayoutScrollTop(
    scrollParent,
    computeViewportAnchoredScrollTop(anchor, nextTargetTop, maxScroll),
    anchor.scrollTopSnapshot
  );
}

export function restoreScrollSnapshot(
  anchor: Pick<DisplayListScrollAnchor, 'scrollTopSnapshot'>,
  scrollParent: HTMLElement
): void {
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  setLayoutScrollTop(
    scrollParent,
    Math.min(Math.max(0, anchor.scrollTopSnapshot), maxScroll),
    anchor.scrollTopSnapshot
  );
}
