/**
 * Canvas-mode transient paragraph flash (G2).
 *
 * `DocxEditorRef.scrollToParaId({ highlight })` flashes the target paragraph so
 * the user's eye lands on it after the jump. This overlay draws the flash
 * directly over the visible canvas pages:
 * the paragraph's PM range is resolved to page-local rects through the
 * display-list `range_rects` query (falling back to `anchorRect` for an empty
 * paragraph) and each rect is projected into `overlayTarget` coordinates via the
 * live per-page `<canvas>` rect — the exact projection `CanvasFindHighlightOverlay`
 * / `CanvasSelectionOverlay` use, so the flash lands on the glyphs regardless of
 * the page column's centering, the sidebar-open shift, or zoom. The projection
 * is scroll-invariant (canvas + overlay share `editorContentRef`), so the flash
 * is correct even though it is requested right as the scroll settles.
 *
 * Driven imperatively: the scroll API bumps `request.nonce` to (re)start a
 * flash; the overlay reuses the same `docx-paragraph-flash-fade` keyframe as the
 * painter path (via `--docx-paragraph-flash-color` / `-duration`) and clears
 * itself after `durationMs`, calling `onDone(nonce)` so the parent can drop the
 * request. Non-interactive (`pointer-events: none`) so it never steals the caret.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import {
  DEFAULT_PARAGRAPH_FLASH_COLOR,
  DEFAULT_PARAGRAPH_FLASH_DURATION_MS,
} from '@betteroffice/docx/utils';

/** One flash request, addressed by the paragraph's live body PM range. */
export interface CanvasParagraphFlashRequest {
  /** Paragraph node start position (inclusive). */
  from: number;
  /** Paragraph node end position (exclusive). */
  to: number;
  /** Bumped on every request so re-flashing the same paragraph restarts. */
  nonce: number;
  /** Flash color; defaults to the painter path's color. */
  color?: string;
  /** Flash duration in ms; defaults to the painter path's duration. */
  durationMs?: number;
}

export interface CanvasParagraphFlashOverlayProps {
  /** The active flash request, or null when nothing is flashing. */
  request: CanvasParagraphFlashRequest | null;
  /** Portal target — `editorContentRef.current`, sharing the canvas host's top-left. */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — `range_rects` / `anchorRect` + page sizes for the scale. */
  displayListQueries: DisplayListQueries;
  /** Sidebar open — recompute after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — recompute when the canvas re-rasters larger. */
  zoom: number;
  /** Called with the request's nonce once its flash duration elapses. */
  onDone: (nonce: number) => void;
}

interface ProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

function normalizedDurationMs(durationMs: number | undefined): number {
  if (durationMs == null || !Number.isFinite(durationMs) || durationMs < 0) {
    return DEFAULT_PARAGRAPH_FLASH_DURATION_MS;
  }
  return durationMs;
}

export function CanvasParagraphFlashOverlay({
  request,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
  onDone,
}: CanvasParagraphFlashOverlayProps) {
  const [rects, setRects] = useState<ProjectedRect[]>([]);

  const nonce = request?.nonce ?? -1;
  const from = request?.from ?? 0;
  const to = request?.to ?? 0;

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!request || !host) {
      setRects([]);
      return;
    }

    const recompute = () => {
      const targetRect = overlayTarget.getBoundingClientRect();
      // Project a page-local (px) rect into `overlayTarget` coordinates via the
      // live `<canvas>` rect — identical to CanvasFindHighlightOverlay. The rect
      // already folds in centering, the sidebar shift, and zoom.
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
        };
      };

      // Whole-paragraph line rects; an empty paragraph has none, so fall back to
      // its first-line anchor rect (a thin band the flash still reads on).
      let source = displayListQueries.rangeRects(from, to);
      if (source.length === 0) {
        const anchor = displayListQueries.anchorRect(from);
        if (anchor) source = [anchor];
      }

      const next: ProjectedRect[] = [];
      for (const r of source) {
        const p = project(r);
        if (p) next.push(p);
      }
      setRects(next);
    };

    recompute();

    // Same recompute triggers as the sibling canvas overlays.
    const ro = new ResizeObserver(recompute);
    ro.observe(host);
    ro.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);

    const duration = normalizedDurationMs(request.durationMs);
    const timer = setTimeout(() => {
      setRects([]);
      onDone(nonce);
    }, duration);

    return () => {
      ro.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
      clearTimeout(timer);
    };
    // `nonce` restarts the flash for the same/next paragraph; `from`/`to` carry
    // the resolved range. `request` object identity intentionally not a dep.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nonce, from, to, overlayTarget, canvasHostRef, displayListQueries, sidebarOpen, zoom]);

  if (!request || rects.length === 0) return null;

  const color = request.color?.trim() || DEFAULT_PARAGRAPH_FLASH_COLOR;
  const durationMs = normalizedDurationMs(request.durationMs);

  return createPortal(
    <div
      data-testid="canvas-paragraph-flash"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        pointerEvents: 'none',
        overflow: 'hidden',
        zIndex: 8,
      }}
    >
      {rects.map((r, i) => (
        <div
          key={`flash-${nonce}-${i}-${r.left}-${r.top}`}
          className="docx-canvas-paragraph-flash"
          style={
            {
              position: 'absolute',
              left: r.left,
              top: r.top,
              width: r.width,
              height: r.height,
              pointerEvents: 'none',
              '--docx-paragraph-flash-color': color,
              '--docx-paragraph-flash-duration': `${durationMs}ms`,
            } as React.CSSProperties
          }
        />
      ))}
    </div>,
    overlayTarget
  );
}

export default CanvasParagraphFlashOverlay;
