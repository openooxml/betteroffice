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
  viewport: DOMRect,
  lines: readonly DisplayListVisualLine[],
  capturePosition: CaptureViewportPosition
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
  if (candidates.length === 0 && nearest) candidates.push(nearest);
  for (const candidate of candidates.slice(0, ANCHOR_CANDIDATE_LIMIT)) {
    const position = capturePosition(candidate.line.from);
    if (position) {
      return { target: { kind: 'position', position }, clientY: candidate.clientY };
    }
  }
  return null;
}

/**
 * Last resort for a viewport with no projectable text line (empty or image-only
 * pages). A page index cannot track content across a page-count change, so this
 * only runs when there is no line to anchor to at all.
 */
function visiblePageAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: DOMRect
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
  const scrollerTop = scrollParent.getBoundingClientRect().top;
  return {
    pmPos,
    clientOffset: projected ? projected.clientY - scrollerTop : null,
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
  const viewport = scrollParent.getBoundingClientRect();
  const lines = queries.visualLines();
  const resolved =
    nearestLineAnchor(queries, host, viewport, lines, capturePosition) ??
    visiblePageAnchor(queries, host, viewport);
  return {
    target: resolved?.target ?? null,
    viewportOffset: resolved ? resolved.clientY - viewport.top : 0,
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
  const scrollerTop = scrollParent.getBoundingClientRect().top;
  const nextTargetTop = pinned ? scrollParent.scrollTop + pinned.clientY - scrollerTop : null;
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = computeViewportAnchoredScrollTop(
    { viewportOffset: anchor.clientOffset ?? 0, scrollTopSnapshot: anchor.scrollTopSnapshot },
    nextTargetTop,
    maxScroll
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
  const scrollerTop = scrollParent.getBoundingClientRect().top;
  const nextTargetTop = clientY == null ? null : scrollParent.scrollTop + clientY - scrollerTop;
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = computeViewportAnchoredScrollTop(anchor, nextTargetTop, maxScroll);
}

export function restoreScrollSnapshot(
  anchor: Pick<DisplayListScrollAnchor, 'scrollTopSnapshot'>,
  scrollParent: HTMLElement
): void {
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = Math.min(Math.max(0, anchor.scrollTopSnapshot), maxScroll);
}
