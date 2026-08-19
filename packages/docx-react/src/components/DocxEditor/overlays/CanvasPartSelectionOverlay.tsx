/** Draws the open part's selection geometry over canvas pages. */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import { projectPageLocalRect } from '../internals/canvasProjection';
import type { PartEdit } from '../partEdit';

export interface CanvasPartSelectionOverlayProps {
  /** Which part is open — a header/footer band, or one note. */
  part: PartEdit;
  /** Current selection, in the open part's own display positions, or null. */
  selection: { from: number; to: number } | null;
  /** Portal target — `editorContentRef.current` (shares the canvas host's top-left). */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — part geometry source + page-local → canvas scale. */
  displayListQueries: DisplayListQueries;
  /** Exact display page whose band was activated. Notes paint on one page. */
  activePageIndex?: number;
  /** Sidebar open — recompute after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — recompute when page geometry scales. */
  zoom: number;
}

interface ProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** The open part's own story answers, never the body's. */
function partSelectionRects(
  queries: DisplayListQueries,
  part: PartEdit,
  from: number,
  to: number
): DisplayListRect[] {
  const start = Math.min(from, to);
  const end = Math.max(from, to);
  switch (part.kind) {
    case 'footnote':
    case 'endnote':
      return start === end
        ? queries.noteCaretRects(part.kind, part.noteId, start)
        : queries.noteRangeRects(part.kind, part.noteId, start, end);
    case 'header':
    case 'footer':
      if (!part.rId) return [];
      return start === end
        ? queries.hfCaretRects(part.kind, part.rId, start)
        : queries.hfRangeRects(part.kind, part.rId, start, end);
  }
}

export function CanvasPartSelectionOverlay({
  part,
  selection,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  activePageIndex,
  sidebarOpen,
  zoom,
}: CanvasPartSelectionOverlayProps) {
  const [state, setState] = useState<{ caret: ProjectedRect | null; rects: ProjectedRect[] }>({
    caret: null,
    rects: [],
  });

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host || !selection) {
      setState({ caret: null, rects: [] });
      return;
    }
    const { from, to } = selection;
    const isCaret = from === to;

    const recompute = () => {
      // Candidates come back one-per-page (an HF part paints on every page it
      // covers, a note on exactly one); each rect carries its own pageIndex.
      const candidates = partSelectionRects(displayListQueries, part, from, to);
      if (candidates.length === 0) {
        setState({ caret: null, rects: [] });
        return;
      }

      // Render on the page the user is editing = the one whose band sits
      // nearest the viewport center (the display-list analogue of the DOM
      // path's nearest visible HF-page behavior.
      const pageIndices = [...new Set(candidates.map((c) => c.pageIndex))];
      let bestPage =
        activePageIndex != null && pageIndices.includes(activePageIndex)
          ? activePageIndex
          : pageIndices[0];
      let bestDist = Infinity;
      for (const pi of activePageIndex == null ? pageIndices : []) {
        const canvasEl = host.querySelector<HTMLCanvasElement>(`canvas[data-page-index="${pi}"]`);
        if (!canvasEl) continue;
        const r = canvasEl.getBoundingClientRect();
        const vpCenter = window.innerHeight / 2;
        const dist = Math.abs((r.top + r.bottom) / 2 - vpCenter);
        if (dist < bestDist) {
          bestDist = dist;
          bestPage = pi;
        }
      }

      const projected: ProjectedRect[] = [];
      for (const c of candidates) {
        if (c.pageIndex !== bestPage) continue;
        const p = projectPageLocalRect(
          host,
          overlayTarget,
          displayListQueries,
          c.pageIndex,
          c.x,
          c.y,
          c.width,
          c.height
        );
        if (!p) continue;
        projected.push({ left: p.left, top: p.top, width: p.width, height: p.height });
      }

      if (isCaret) {
        setState({ caret: projected[0] ?? null, rects: [] });
      } else {
        setState({ caret: null, rects: projected });
      }
    };

    recompute();

    // Same geometry-invalidation signals the body overlay watches: host/target
    // resize (window resize, scrollbar toggle) and the sidebar/zoom transitions
    // that animate on the inner column and end with a bubbling `transitionend`.
    const ro = new ResizeObserver(recompute);
    ro.observe(host);
    ro.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [
    part,
    selection,
    overlayTarget,
    canvasHostRef,
    displayListQueries,
    activePageIndex,
    sidebarOpen,
    zoom,
  ]);

  if (!state.caret && state.rects.length === 0) return null;

  return createPortal(
    <>
      {state.caret && (
        <div
          aria-hidden="true"
          data-testid="part-caret"
          style={{
            position: 'absolute',
            top: state.caret.top,
            left: state.caret.left,
            width: 2,
            height: state.caret.height,
            background: '#4285f4',
            pointerEvents: 'none',
            zIndex: 11,
            animation: 'part-caret-blink 1.06s steps(1) infinite',
          }}
        />
      )}
      {state.rects.map((r, i) => (
        <div
          key={`part-sel-${i}-${r.top}-${r.left}`}
          aria-hidden="true"
          data-testid="part-selection-rect"
          style={{
            position: 'absolute',
            top: r.top,
            left: r.left,
            width: r.width,
            height: r.height,
            background: 'rgba(66, 133, 244, 0.25)',
            pointerEvents: 'none',
            zIndex: 10,
          }}
        />
      ))}
    </>,
    overlayTarget
  );
}

export default CanvasPartSelectionOverlay;
