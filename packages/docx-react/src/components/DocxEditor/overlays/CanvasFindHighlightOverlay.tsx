/**
 * Canvas-mode find/replace match highlighting.
 *
 * This component draws find highlights directly over the visible canvas pages:
 * every match's display range is resolved to
 * page-local rects through the display-list `range_rects` query and each rect is
 * projected into `overlayTarget` coordinates via the live per-page `<canvas>`
 * rect — the exact projection `CanvasSelectionOverlay` uses, so highlights land
 * on the glyphs regardless of the page column's centering, the sidebar-open
 * shift, or zoom.
 *
 * The current match gets the `.docx-find-highlight-current` token; the rest get
 * `.docx-find-highlight` — the same single-source-of-truth classes the DOM
 * find path defines in `packages/docx/src/styles/editor.css`. Non-interactive
 * (pointer-events: none) so it never steals the caret.
 *
 * Renders nothing on the default DOM-painter path: it is only mounted once the
 * canvas overlay target resolves (canvas active).
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';

/** One find match, addressed by its live display range. */
export interface CanvasFindMatch {
  displayFrom: number;
  displayTo: number;
}

export interface CanvasFindHighlightOverlayProps {
  /** All matches in document order, each carrying its live display range. */
  matches: CanvasFindMatch[];
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

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host || matches.length === 0) {
      setRects([]);
      return;
    }

    const recompute = () => {
      const targetRect = overlayTarget.getBoundingClientRect();
      // Project a page-local (px) rect on `pageIndex` into `overlayTarget`
      // coordinates via the live `<canvas>` rect — identical to
      // CanvasSelectionOverlay. The rect already folds in centering, the
      // sidebar shift, and zoom (canvas CSS width = page * zoom ⇒ scale = zoom).
      const project = (r: DisplayListRect): ProjectedRect | null => {
        const canvasEl = host.querySelector<HTMLCanvasElement>(
          `canvas[data-page-index="${r.pageIndex}"]`
        );
        const size = displayListQueries.pageSize(r.pageIndex);
        if (!canvasEl || !size) return null;
        const canvasRect = canvasEl.getBoundingClientRect();
        const scaleX = size.width > 0 ? canvasRect.width / size.width : 1;
        const scaleY = size.height > 0 ? canvasRect.height / size.height : 1;
        return {
          left: canvasRect.left - targetRect.left + r.x * scaleX,
          top: canvasRect.top - targetRect.top + r.y * scaleY,
          width: r.width * scaleX,
          height: r.height * scaleY,
          isCurrent: false,
        };
      };

      const next: ProjectedRect[] = [];
      for (let i = 0; i < matches.length; i++) {
        const m = matches[i];
        const isCurrent = i === currentIndex;
        for (const r of displayListQueries.rangeRects(m.displayFrom, m.displayTo)) {
          const p = project(r);
          if (p) next.push({ ...p, isCurrent });
        }
      }
      setRects(next);
    };

    recompute();

    // Same recompute triggers as CanvasSelectionOverlay: host/target resize,
    // and the sidebar/zoom transition that ends with a bubbling `transitionend`.
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
  }, [matches, currentIndex, overlayTarget, canvasHostRef, displayListQueries, sidebarOpen, zoom]);

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
