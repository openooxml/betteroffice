import { useCallback, useEffect, useRef } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import type { ChartResizeEdges } from '@betteroffice/xlsx';
import { chartResizeEdges, type ChartDragPreview, type ChartResizeHandle } from './chartResize';

export type ChartResizeCommit = (handle: ChartResizeHandle, edges: ChartResizeEdges) => void;

interface Drag {
  pointerId: number;
  handle: ChartResizeHandle;
  startX: number;
  startY: number;
  element: HTMLElement;
  dx: number;
  dy: number;
}

export interface ChartResizeGestureOptions {
  /** CSS pixels per content pixel at the moment the pointer moves. */
  zoom: () => number;
  /** Draw the live frame; `null` clears it. */
  preview: (preview: ChartDragPreview | null) => void;
  /** Called exactly once, on a primary-button release that moved the pointer. */
  commit: ChartResizeCommit;
}

/**
 * The one chart-resize gesture both editable surfaces use: pointerdown on a
 * handle captures the pointer, pointermove previews the frame in content
 * pixels, pointerup commits one edge movement, and Escape/pointercancel/blur
 * drop the gesture without touching the workbook. All bookkeeping lives in a
 * ref so an edited preview rerender cannot restart the drag.
 */
export function useChartResizeGesture(options: ChartResizeGestureOptions) {
  const drag = useRef<Drag | null>(null);
  const optionsRef = useRef(options);
  optionsRef.current = options;

  const finish = useCallback((commit: boolean) => {
    const current = drag.current;
    if (!current) return;
    drag.current = null;
    try {
      if (current.element.hasPointerCapture?.(current.pointerId)) {
        current.element.releasePointerCapture(current.pointerId);
      }
    } catch {
      // capture is best-effort: the pointer can already be gone.
    }
    optionsRef.current.preview(null);
    if (commit && (current.dx !== 0 || current.dy !== 0)) {
      optionsRef.current.commit(current.handle, chartResizeEdges(current.handle, current.dx, current.dy));
    }
  }, []);

  useEffect(() => {
    const move = (event: PointerEvent) => {
      const current = drag.current;
      if (!current || event.pointerId !== current.pointerId) return;
      const zoom = optionsRef.current.zoom() || 1;
      current.dx = (event.clientX - current.startX) / zoom;
      current.dy = (event.clientY - current.startY) / zoom;
      optionsRef.current.preview({ kind: 'resize', handle: current.handle, dx: current.dx, dy: current.dy });
      event.preventDefault();
    };
    const up = (event: PointerEvent) => {
      const current = drag.current;
      if (!current || event.pointerId !== current.pointerId) return;
      finish(event.type === 'pointerup' && event.button === 0);
    };
    const cancel = (event: PointerEvent) => {
      if (!drag.current || event.pointerId !== drag.current.pointerId) return;
      finish(false);
    };
    const blur = () => {
      if (drag.current) finish(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || !drag.current) return;
      // the grid also treats Escape as "drop the chart selection"; cancel the
      // gesture first and keep the key from clearing the selection underneath.
      event.preventDefault();
      event.stopPropagation();
      finish(false);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', cancel);
    window.addEventListener('blur', blur);
    window.addEventListener('keydown', keydown, true);
    return () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', cancel);
      window.removeEventListener('blur', blur);
      window.removeEventListener('keydown', keydown, true);
      const current = drag.current;
      drag.current = null;
      if (current) {
        try {
          if (current.element.hasPointerCapture?.(current.pointerId)) {
            current.element.releasePointerCapture(current.pointerId);
          }
        } catch {
          // unmounting mid-gesture is the same best-effort release.
        }
      }
    };
  }, [finish]);

  const start = useCallback((handle: ChartResizeHandle, event: ReactPointerEvent<HTMLElement>) => {
    if (event.button !== 0 || drag.current) return;
    event.preventDefault();
    event.stopPropagation();
    const element = event.currentTarget;
    try {
      element.setPointerCapture(event.pointerId);
    } catch {
      // happy-dom and older browsers do not implement pointer capture.
    }
    drag.current = { pointerId: event.pointerId, handle, startX: event.clientX, startY: event.clientY,
      element, dx: 0, dy: 0 };
    optionsRef.current.preview({ kind: 'resize', handle, dx: 0, dy: 0 });
  }, []);

  const cancel = useCallback(() => finish(false), [finish]);

  /** Keyboard resize: one edge movement through the same commit path. */
  const nudge = useCallback((handle: ChartResizeHandle, dx: number, dy: number) => {
    if (dx === 0 && dy === 0) return;
    optionsRef.current.preview(null);
    optionsRef.current.commit(handle, chartResizeEdges(handle, dx, dy));
  }, []);

  const active = useCallback(() => drag.current !== null, []);

  return { start, cancel, nudge, active };
}
