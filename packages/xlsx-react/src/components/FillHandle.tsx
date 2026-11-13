import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react';
import type { CellAddr, CellRange } from '@betteroffice/xlsx';

/** Side length of the visible handle and of its pointer hit target. */
export const FILL_HANDLE_SIZE = 8;
/** How far the handle's corner sits past the selection's bottom-right corner. */
export const FILL_HANDLE_OVERHANG = FILL_HANDLE_SIZE / 2;

export interface FillHandleRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Where the handle's square sits for a selection rect, in overlay coordinates. */
export function fillHandlePosition(rect: FillHandleRect): { left: number; top: number; size: number } {
  return {
    left: rect.x + rect.w - FILL_HANDLE_OVERHANG,
    top: rect.y + rect.h - FILL_HANDLE_OVERHANG,
    size: FILL_HANDLE_SIZE,
  };
}

/**
 * The rect that places the handle: the true bottom-right cell when it has
 * visible pixels, else the visible edge of the clipped selection. Selecting
 * past the viewport therefore keeps the handle pinned to what can be drawn
 * without moving it to a different cell than the selection's real corner.
 */
export function fillHandleAnchor(
  clipped: FillHandleRect | null,
  bottomRight: FillHandleRect | null
): FillHandleRect | null {
  return bottomRight ?? clipped;
}

/**
 * The complete fill rectangle for a drag endpoint: the source block plus
 * everything below and/or right of it. A pointer dragged up or left clamps to
 * the source's own bottom-right corner, so a fill never shrinks the source.
 */
export function fillDestination(source: CellRange, endpoint: CellAddr): CellRange {
  return {
    top: source.top,
    left: source.left,
    bottom: Math.max(source.bottom, endpoint.row),
    right: Math.max(source.right, endpoint.col),
  };
}

export interface FillHandleProps {
  /** Selection rect in the overlay's coordinate space; the handle sits on its corner. */
  rect: FillHandleRect;
  /** The selection to tile, already expanded over merged cells. */
  source: CellRange;
  /** Maps a cell range to its overlay rect; `null` when it has no visible pixels. */
  rectForRange(range: CellRange): FillHandleRect | null;
  /** Maps pointer client coordinates to the painted cell under them. */
  cellAt(clientX: number, clientY: number): CellAddr | null;
  /** Called once on release when the drag extended the source. */
  onFill(source: CellRange, destination: CellRange): void;
  /** Translated accessible name for the handle. */
  ariaLabel: string;
}

interface FillDrag {
  pointerId: number;
  endpoint: CellAddr;
  startX: number;
  startY: number;
  moved: boolean;
}

/**
 * Excel's fill handle: one shared overlay square for every editable grid.
 * Dragging it down and/or right previews the filled rectangle and commits a
 * single fill on release; Escape abandons the gesture without writing.
 */
export function FillHandle({
  rect,
  source,
  rectForRange,
  cellAt,
  onFill,
  ariaLabel,
}: FillHandleProps) {
  const [drag, setDrag] = useState<FillDrag | null>(null);
  const dragRef = useRef<FillDrag | null>(null);
  dragRef.current = drag;

  const cancel = useCallback(() => {
    dragRef.current = null;
    setDrag(null);
  }, []);

  const dragging = drag !== null;

  useEffect(() => {
    if (!dragging) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      event.stopPropagation();
      cancel();
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [dragging, cancel]);

  const destination = drag ? fillDestination(source, drag.endpoint) : null;
  const preview = destination && (destination.bottom > source.bottom || destination.right > source.right)
    ? rectForRange(destination) : null;

  const onPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    try {
      event.currentTarget.setPointerCapture?.(event.pointerId);
    } catch {
      // pointer capture is best effort; bubbling pointer events still work.
    }
    const next = {
      pointerId: event.pointerId,
      endpoint: { row: source.bottom, col: source.right },
      startX: event.clientX,
      startY: event.clientY,
      moved: false,
    };
    dragRef.current = next;
    setDrag(next);
  };

  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const current = dragRef.current;
    if (!current || event.pointerId !== current.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    const endpoint = cellAt(event.clientX, event.clientY);
    if (!endpoint) return;
    const moved = current.moved || event.clientX !== current.startX || event.clientY !== current.startY;
    const next = { ...current, endpoint, moved };
    dragRef.current = next;
    setDrag(next);
  };

  const finish = (event: ReactPointerEvent<HTMLDivElement>) => {
    const current = dragRef.current;
    if (!current || event.pointerId !== current.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    const endpoint = cellAt(event.clientX, event.clientY) ?? current.endpoint;
    dragRef.current = null;
    setDrag(null);
    try {
      event.currentTarget.releasePointerCapture?.(event.pointerId);
    } catch {
      // the pointer may already be gone.
    }
    const moved = current.moved || event.clientX !== current.startX || event.clientY !== current.startY;
    const completed = fillDestination(source, endpoint);
    if (moved && (completed.bottom > source.bottom || completed.right > source.right)) onFill(source, completed);
  };

  const onPointerCancel = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!dragRef.current || event.pointerId !== dragRef.current.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    cancel();
  };

  const corner = fillHandlePosition(rect);
  const handleStyle: CSSProperties = {
    position: 'absolute',
    left: corner.left,
    top: corner.top,
    width: corner.size,
    height: corner.size,
    boxSizing: 'border-box',
    background: '#217346',
    border: '1px solid #ffffff',
    cursor: 'crosshair',
    pointerEvents: 'auto',
    touchAction: 'none',
  };
  const previewStyle: CSSProperties | null = preview
    ? {
        position: 'absolute',
        left: preview.x,
        top: preview.y,
        width: preview.w,
        height: preview.h,
        boxSizing: 'border-box',
        border: '1px dashed #217346',
        background: 'rgba(33, 115, 70, 0.06)',
        pointerEvents: 'none',
      }
    : null;

  return (
    <>
      {previewStyle && <div data-testid="xlsx-fill-preview" aria-hidden style={previewStyle} />}
      <div
        data-testid="xlsx-fill-handle"
        role="button"
        tabIndex={-1}
        aria-label={ariaLabel}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={finish}
        onPointerCancel={onPointerCancel}
        onLostPointerCapture={onPointerCancel}
        // a pointer press also produces compatibility mouse events; stop the
        // grid's own mousedown/move/click handling from starting a selection.
        onMouseDown={(event) => {
          event.preventDefault();
          event.stopPropagation();
        }}
        onMouseMove={(event) => {
          if (dragRef.current) event.stopPropagation();
        }}
        onClick={(event) => event.stopPropagation()}
        onDoubleClick={(event) => event.stopPropagation()}
        style={handleStyle}
      />
    </>
  );
}
