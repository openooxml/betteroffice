/**
 * Canvas-mode find match highlighting.
 *
 * Draws find highlights over the canvas pages in view and one page either
 * side: only the matches in those pages' position range are resolved to
 * page-local rects through the display-list `range_rects` query, so the cost
 * follows what is on screen, not the match count. Each rect is projected into
 * `overlayTarget` coordinates via the live per-page `<canvas>` rect, the
 * projection `CanvasSelectionOverlay` uses, so highlights land on the glyphs
 * regardless of the page column's centering, the sidebar-open shift, or zoom.
 * The set is recomputed when the pages in view change or move.
 *
 * The current match gets the `.docx-find-highlight-current` token; the rest get
 * `.docx-find-highlight`, the classes the DOM find path defines in
 * `packages/docx/src/styles/editor.css`. Non-interactive (pointer-events: none)
 * so it never steals the caret.
 */

import { useLayoutEffect, useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  displayPageCanvas,
  effectiveZoom,
  type DisplayList,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import { observePageWindow } from './pageWindowObserver';
import {
  buildRemotePresencePageMetrics,
  type RemotePresencePageWindow,
} from './remotePresenceGeometry';

/** One find match, addressed by its live display range. */
export interface CanvasFindMatch {
  displayFrom: number;
  displayTo: number;
}

export interface CanvasFindHighlightOverlayProps {
  /** All matches, each carrying its live display range. */
  matches: readonly CanvasFindMatch[];
  /** Index of the active match (styled distinctly), or -1 for none. */
  currentIndex: number;
  /** Portal target — `editorContentRef.current`, sharing the canvas host's top-left. */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — `range_rects` per match + page sizes for the scale. */
  displayListQueries: DisplayListQueries;
  /** Sidebar open — recompute after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — recompute when the canvas re-rasters larger. */
  zoom: number;
}

interface ProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
  isCurrent: boolean;
}

/** Match indices in display order, with the furthest end reached up to each. */
export function displayOrder(matches: readonly CanvasFindMatch[]): {
  order: number[];
  reach: number[];
} {
  const order = matches.map((_, index) => index);
  if (matches.some((match, index) => index > 0 && match.displayFrom < matches[index - 1].displayFrom)) {
    order.sort((a, b) => matches[a].displayFrom - matches[b].displayFrom || a - b);
  }
  const reach: number[] = [];
  for (const index of order) {
    reach.push(Math.max(reach.at(-1) ?? -Infinity, matches[index].displayTo));
  }
  return { order, reach };
}

/** Indices of the matches overlapping `[from, to)`, in display order. */
export function matchesInRange(
  matches: readonly CanvasFindMatch[],
  { order, reach }: ReturnType<typeof displayOrder>,
  from: number,
  to: number
): number[] {
  let low = 0;
  let high = order.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (reach[middle] <= from) low = middle + 1;
    else high = middle;
  }
  const found: number[] = [];
  for (let index = low; index < order.length; index += 1) {
    const match = matches[order[index]];
    if (match.displayFrom >= to) break;
    if (match.displayTo > from) found.push(order[index]);
  }
  return found;
}

/**
 * The document positions the pages of `pageWindow` paint, as sorted disjoint
 * intervals. A table page repeating its header rows paints them again, far
 * from the rest of the page's positions.
 */
export function pagePositionIntervals(
  displayList: DisplayList,
  pageWindow: RemotePresencePageWindow
): Array<{ from: number; to: number }> {
  const spans: Array<{ from: number; to: number }> = [];
  for (let pageIndex = pageWindow.start; pageIndex <= pageWindow.end; pageIndex += 1) {
    for (const primitive of displayList.pages[pageIndex]?.primitives ?? []) {
      if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun' && primitive.kind !== 'image') {
        continue;
      }
      const from = primitive.docStart;
      const to = primitive.docEnd;
      if (!Number.isSafeInteger(from) || !Number.isSafeInteger(to)) continue;
      spans.push({ from: from!, to: Math.max(to!, from! + 1) });
    }
  }
  spans.sort((a, b) => a.from - b.from);
  const merged: Array<{ from: number; to: number }> = [];
  for (const span of spans) {
    const last = merged.at(-1);
    if (last && span.from <= last.to + 1) last.to = Math.max(last.to, span.to);
    else merged.push({ ...span });
  }
  return merged;
}

export function CanvasFindHighlightOverlay({
  matches,
  currentIndex,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: CanvasFindHighlightOverlayProps) {
  const [rects, setRects] = useState<ProjectedRect[]>([]);
  const order = useMemo(() => displayOrder(matches), [matches]);

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host || matches.length === 0) {
      setRects([]);
      return;
    }
    const displayList = displayListQueries.displayList;
    const metrics = buildRemotePresencePageMetrics(displayList, zoom);
    let shown: string | null = null;
    return observePageWindow(host, metrics, (pageWindow, moved) => {
      const key = pageWindow ? `${pageWindow.start}:${pageWindow.end}` : '';
      if (!moved && key === shown) return;
      shown = key;
      const intervals = pageWindow ? pagePositionIntervals(displayList, pageWindow) : [];
      if (!pageWindow || intervals.length === 0) {
        setRects([]);
        return;
      }
      const targetRect = overlayTarget.getBoundingClientRect();
      const targetZoom = effectiveZoom(overlayTarget);
      // Project a page-local (px) rect on `pageIndex` into `overlayTarget`
      // coordinates via the live `<canvas>` rect, as CanvasSelectionOverlay
      // does. The rect already folds in centering, the sidebar shift, and zoom.
      const project = (r: DisplayListRect, isCurrent: boolean): ProjectedRect | null => {
        if (r.pageIndex < pageWindow.start || r.pageIndex > pageWindow.end) return null;
        const canvasEl = displayPageCanvas(host, r.pageIndex);
        const size = displayListQueries.pageSize(r.pageIndex);
        if (!canvasEl || !size) return null;
        const canvasRect = canvasEl.getBoundingClientRect();
        const scaleX = (size.width > 0 ? canvasRect.width / size.width : 1) / targetZoom;
        const scaleY = (size.height > 0 ? canvasRect.height / size.height : 1) / targetZoom;
        return {
          left: (canvasRect.left - targetRect.left) / targetZoom + r.x * scaleX,
          top: (canvasRect.top - targetRect.top) / targetZoom + r.y * scaleY,
          width: r.width * scaleX,
          height: r.height * scaleY,
          isCurrent,
        };
      };
      const next: ProjectedRect[] = [];
      const visible = new Set<number>();
      for (const { from, to } of intervals) {
        for (const index of matchesInRange(matches, order, from, to)) visible.add(index);
      }
      for (const index of visible) {
        const m = matches[index];
        const source =
          displayListQueries.rangeRectsOnPages?.(
            m.displayFrom,
            m.displayTo,
            pageWindow.start,
            pageWindow.end
          ) ?? displayListQueries.rangeRects(m.displayFrom, m.displayTo);
        for (const r of source) {
          const p = project(r, index === currentIndex);
          if (p) next.push(p);
        }
      }
      setRects(next);
    });
  }, [
    matches,
    order,
    currentIndex,
    overlayTarget,
    canvasHostRef,
    displayListQueries,
    sidebarOpen,
    zoom,
  ]);

  if (rects.length === 0) return null;

  return createPortal(
    <div
      data-testid="canvas-find-highlights"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        pointerEvents: 'none',
        overflow: 'hidden',
        zIndex: 9,
      }}
    >
      {rects.map((r, i) => (
        <div
          key={`find-${i}-${r.left}-${r.top}`}
          className={r.isCurrent ? 'docx-find-highlight-current' : 'docx-find-highlight'}
          style={{
            position: 'absolute',
            left: r.left,
            top: r.top,
            width: r.width,
            height: r.height,
            pointerEvents: 'none',
          }}
        />
      ))}
    </div>,
    overlayTarget
  );
}

export default CanvasFindHighlightOverlay;
