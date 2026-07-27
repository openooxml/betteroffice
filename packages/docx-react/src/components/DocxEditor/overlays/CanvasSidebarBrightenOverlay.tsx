/** Brightens the expanded sidebar item's range over canvas pages. */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';

/** Tint variant — mirrors the painter `<style>` colors. */
export type CanvasBrightenVariant = 'comment' | 'insertion' | 'deletion';

/** The expanded sidebar item's resolved body range + tint. */
export interface CanvasBrightenRange {
  from: number;
  to: number;
  variant: CanvasBrightenVariant;
}

export interface CanvasSidebarBrightenOverlayProps {
  /** Resolved range of the expanded card, or null when nothing is expanded. */
  range: CanvasBrightenRange | null;
  /** Portal target — `editorContentRef.current`, sharing the canvas host's top-left. */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — `range_rects` per range + page sizes for the scale. */
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
}

const VARIANT_CLASS: Record<CanvasBrightenVariant, string> = {
  comment: 'docx-canvas-brighten-comment',
  insertion: 'docx-canvas-brighten-insertion',
  deletion: 'docx-canvas-brighten-deletion',
};

export function CanvasSidebarBrightenOverlay({
  range,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: CanvasSidebarBrightenOverlayProps) {
  const [rects, setRects] = useState<ProjectedRect[]>([]);

  const from = range?.from ?? 0;
  const to = range?.to ?? 0;

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!range || !host) {
      setRects([]);
      return;
    }

    const recompute = () => {
      const targetRect = overlayTarget.getBoundingClientRect();
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

      const next: ProjectedRect[] = [];
      for (const r of displayListQueries.rangeRects(from, to)) {
        const p = project(r);
        if (p) next.push(p);
      }
      setRects(next);
    };

    recompute();

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
  }, [range, from, to, overlayTarget, canvasHostRef, displayListQueries, sidebarOpen, zoom]);

  if (!range || rects.length === 0) return null;

  const className = VARIANT_CLASS[range.variant];

  return createPortal(
    <div
      data-testid="canvas-sidebar-brighten"
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
          key={`brighten-${i}-${r.left}-${r.top}`}
          className={className}
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

export default CanvasSidebarBrightenOverlay;
