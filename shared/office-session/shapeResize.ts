export type ShapeResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export interface ShapeResizeRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ShapeResizeDelta {
  x: number;
  y: number;
}

/** The rectangle produced by dragging a resize handle. Opposite edges stay
 * fixed, and the rectangle never collapses past `minimum`. */
export function resizedShapeRect(
  bounds: ShapeResizeRect,
  handle: ShapeResizeHandle,
  delta: ShapeResizeDelta,
  minimum: number,
  aspectRatioLocked = false
): ShapeResizeRect {
  if (aspectRatioLocked && handle.length === 2 && bounds.width > 0 && bounds.height > 0) {
    const dx = handle.includes('w') ? -delta.x : delta.x;
    const dy = handle.startsWith('n') ? -delta.y : delta.y;
    const scale = Math.max(1 + (dx * bounds.width + dy * bounds.height) /
      (bounds.width ** 2 + bounds.height ** 2),
    minimum / bounds.width, minimum / bounds.height);
    const width = Math.round(bounds.width * scale);
    const height = Math.round(bounds.height * scale);
    return {
      x: handle.includes('w') ? bounds.x + bounds.width - width : bounds.x,
      y: handle.startsWith('n') ? bounds.y + bounds.height - height : bounds.y,
      width, height,
    };
  }
  let { x, y, width, height } = bounds;
  if (handle.includes('w')) {
    const dx = Math.min(delta.x, width - minimum);
    x += dx;
    width -= dx;
  } else if (handle.includes('e')) {
    width = Math.max(minimum, width + delta.x);
  }
  if (handle.startsWith('n')) {
    const dy = Math.min(delta.y, height - minimum);
    y += dy;
    height -= dy;
  } else if (handle.startsWith('s')) {
    height = Math.max(minimum, height + delta.y);
  }
  return { x, y, width, height };
}
