/** Draws header/footer selection geometry over canvas pages. */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import {
  computeHfCaretRectsFromDisplayList,
  computeHfSelectionRectsFromDisplayList,
} from '@betteroffice/docx/layout';
import { projectPageLocalRect } from '../internals/canvasProjection';

export interface CanvasHfSelectionOverlayProps {
  /** Which HF band is being edited. */
  region: 'header' | 'footer';
  /** Relationship id of the active HF part (the display-list variant's `rId`). */
  rId: string;
  /** Current HF selection (positions in the HF doc), or null. */
  selection: { from: number; to: number } | null;
  /** Portal target — `editorContentRef.current` (shares the canvas host's top-left). */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — HF geometry source + page-local → canvas scale. */
  displayListQueries: DisplayListQueries;
  /** Exact display page whose HF band was activated. */
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

export function CanvasHfSelectionOverlay({
  region,
  rId,
  selection,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  activePageIndex,
  sidebarOpen,
  zoom,
}: CanvasHfSelectionOverlayProps) {
  const [state, setState] = useState<{ caret: ProjectedRect | null; rects: ProjectedRect[] }>({
    caret: null,
    rects: [],
  });

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host || !rId || !selection) {
      setState({ caret: null, rects: [] });
      return;
    }
    const { from, to } = selection;
    const isCaret = from === to;

    const recompute = () => {
      // Candidates come back one-per-page (the HF part paints on every page it
      // covers); each rect carries its own pageIndex.
      const candidates = isCaret
        ? computeHfCaretRectsFromDisplayList(displayListQueries, region, rId, from)
        : computeHfSelectionRectsFromDisplayList(displayListQueries, region, rId, from, to);
      if (candidates.length === 0) {
        setState({ caret: null, rects: [] });
        return;
      }

      // Render on the page the user is editing = the one whose HF band sits
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
    region,
    rId,
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
          data-testid="hf-caret"
          style={{
            position: 'absolute',
            top: state.caret.top,
            left: state.caret.left,
            width: 2,
            height: state.caret.height,
            background: '#4285f4',
            pointerEvents: 'none',
            zIndex: 11,
            animation: 'hf-caret-blink 1.06s steps(1) infinite',
          }}
        />
      )}
      {state.rects.map((r, i) => (
        <div
          key={`hf-sel-${i}-${r.top}-${r.left}`}
          aria-hidden="true"
          data-testid="hf-selection-rect"
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

export default CanvasHfSelectionOverlay;
