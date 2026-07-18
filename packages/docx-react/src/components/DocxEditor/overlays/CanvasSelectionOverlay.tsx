/**
 * Canvas-mode selection + caret overlay.
 *
 * The caret / selection rects come from the display list. This component
 * portals a SelectionOverlay onto the visible canvas pages
 * (portal target = `editorContentRef`, the positioned ancestor that shares the
 * `.canvas-pages` host's top-left, exactly like the comment overlays fixed in
 * #92). Reusing SelectionOverlay keeps the caret-blink animation, the
 * `--doc-caret` token, and the selection-highlight color byte-identical to the
 * existing overlay implementation.
 *
 * The display-list rects come back page-local (px, per page) + a pageIndex. We
 * convert each to `editorContentRef`-relative coordinates through the live
 * per-page `<canvas>` rect — the same conversion `useFloatingCommentBtn` uses
 * for the "Add comment" button — so the caret lands on the glyph regardless of
 * the page column's centering, the sidebar-open `translateX` shift, or zoom.
 * The result is scroll-invariant (both the canvas and this overlay live inside
 * `editorContentRef` and scroll together), so it is recomputed only on
 * selection change, host/target resize, and the sidebar/zoom transition end —
 * never per scroll tick.
 *
 * Renders nothing on the default DOM-painter path: PagedEditor only mounts this
 * (in place of its inline overlay) once `overlayTarget` resolves, which
 * `useCanvasOverlayTarget` returns non-null only while the canvas paints.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { SelectionRect, CaretPosition } from '@betteroffice/docx/layout';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { SelectionOverlay } from './SelectionOverlay';

export interface CanvasSelectionOverlayProps {
  /** Selection rects in page-local px + pageIndex (from the display list). */
  selectionRects: SelectionRect[];
  /** Caret position in page-local px + pageIndex, or null for a range. */
  caretPosition: CaretPosition | null;
  /** Whether the (hidden body) ProseMirror is focused — drives the blink. */
  isFocused: boolean;
  /** Hide the caret / selection in read-only mode. */
  readOnly?: boolean;
  /**
   * Portal target — `editorContentRef.current`, the positioned ancestor that
   * shares the `.canvas-pages` host's top-left. Converted coordinates are
   * expressed relative to it.
   */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — page sizes drive the page-local → canvas scale. */
  displayListQueries: DisplayListQueries;
  /** Sidebar open — recompute after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — recompute when page geometry scales. */
  zoom: number;
}

export function CanvasSelectionOverlay({
  selectionRects,
  caretPosition,
  isFocused,
  readOnly = false,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: CanvasSelectionOverlayProps) {
  const [converted, setConverted] = useState<{
    rects: SelectionRect[];
    caret: CaretPosition | null;
  }>({ rects: [], caret: null });

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host) {
      setConverted({ rects: [], caret: null });
      return;
    }

    // Project a page-local (px) point on `pageIndex` into `overlayTarget`
    // coordinates via the live `<canvas>` rect. The rect already reflects the
    // page column's centering, the sidebar-open shift, and any zoom, so the
    // caret sits on the glyph in every case. `scaleX/Y` are 1 while the canvas
    // paints at logical px (its CSS size equals the display-list page size) but
    // stay correct if a future zoom scales the canvas.
    const recompute = () => {
      const targetRect = overlayTarget.getBoundingClientRect();
      const project = (pageIndex: number, x: number, y: number) => {
        const canvasEl = host.querySelector<HTMLCanvasElement>(
          `canvas[data-page-index="${pageIndex}"]`
        );
        const size = displayListQueries.pageSize(pageIndex);
        if (!canvasEl || !size) return null;
        const canvasRect = canvasEl.getBoundingClientRect();
        const scaleX = size.width > 0 ? canvasRect.width / size.width : 1;
        const scaleY = size.height > 0 ? canvasRect.height / size.height : 1;
        return {
          left: canvasRect.left - targetRect.left + x * scaleX,
          top: canvasRect.top - targetRect.top + y * scaleY,
          scaleX,
          scaleY,
        };
      };

      const nextRects: SelectionRect[] = [];
      for (const r of selectionRects) {
        const p = project(r.pageIndex, r.x, r.y);
        if (!p) continue;
        nextRects.push({
          x: p.left,
          y: p.top,
          width: r.width * p.scaleX,
          height: r.height * p.scaleY,
          pageIndex: r.pageIndex,
        });
      }

      let nextCaret: CaretPosition | null = null;
      if (caretPosition) {
        const p = project(caretPosition.pageIndex, caretPosition.x, caretPosition.y);
        if (p) {
          nextCaret = {
            x: p.left,
            y: p.top,
            height: caretPosition.height * p.scaleY,
            pageIndex: caretPosition.pageIndex,
          };
        }
      }

      setConverted({ rects: nextRects, caret: nextCaret });
    };

    recompute();

    // Host / target resize (window resize, scrollbar toggle) shifts the page
    // geometry; the sidebar-open `translateX` and zoom animate on the inner
    // column, ending with a `transitionend` that bubbles to the host.
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
    selectionRects,
    caretPosition,
    overlayTarget,
    canvasHostRef,
    displayListQueries,
    sidebarOpen,
    zoom,
  ]);

  return createPortal(
    <SelectionOverlay
      selectionRects={converted.rects}
      caretPosition={converted.caret}
      isFocused={isFocused}
      readOnly={readOnly}
    />,
    overlayTarget
  );
}

export default CanvasSelectionOverlay;
