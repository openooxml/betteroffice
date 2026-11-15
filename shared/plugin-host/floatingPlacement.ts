export interface FloatingRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FloatingSize {
  width: number;
  height: number;
}

export type FloatingSide = 'above' | 'below';

export interface FloatingPlacementOptions {
  /** The side tried first; `above` by default. */
  side?: FloatingSide;
  /** Pixels between the anchor and the box; 8 by default. */
  gap?: number;
  /** Pixels the box keeps from the viewport's edges; 8 by default. */
  margin?: number;
}

/** Rectangles share one coordinate space, such as overlay-layer pixels. */
export interface FloatingPlacementInput extends FloatingPlacementOptions {
  anchor: FloatingRect;
  size: FloatingSize;
  viewport: FloatingRect;
  /** The anchor's page; the part of the anchor outside it does not count. */
  page?: FloatingRect | null;
}

export interface FloatingPlacement extends FloatingRect {
  side: FloatingSide;
  /** Placed on the other side because the preferred one had no room. */
  flipped: boolean;
  /** Moved to stay inside the viewport, or placed from the visible part of a cut anchor. */
  clamped: boolean;
  /** No part of the anchor is visible on its page inside the viewport; the box is still placed inside. */
  hidden: boolean;
}

function intersect(a: FloatingRect, b: FloatingRect): FloatingRect | null {
  const left = Math.max(a.x, b.x);
  const top = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  if (right < left || bottom < top) return null;
  if (left === a.x && top === a.y && right === a.x + a.width && bottom === a.y + a.height) return a;
  return { x: left, y: top, width: right - left, height: bottom - top };
}

const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(value, max));

/** Places a box above or below an anchor, flipping when the preferred side has no room, inside the viewport. */
export function placeFloating({
  anchor, size, viewport, page, side = 'above', gap = 8, margin = 8,
}: FloatingPlacementInput): FloatingPlacement {
  const onPage = page ? intersect(anchor, page) : anchor;
  const visible = onPage && intersect(onPage, viewport);
  const target = visible ?? anchor;
  const top = viewport.y + margin;
  const bottom = viewport.y + viewport.height - margin;
  const above = target.y - gap - size.height;
  const below = target.y + target.height + gap;
  const fits = (candidate: FloatingSide) => candidate === 'above' ? above >= top : below + size.height <= bottom;
  const other: FloatingSide = side === 'above' ? 'below' : 'above';
  const flipped = !fits(side) && fits(other);
  const chosen = flipped ? other : side;
  const idealX = target.x + target.width / 2 - size.width / 2;
  const idealY = chosen === 'above' ? above : below;
  const x = clamp(idealX, viewport.x + margin, viewport.x + viewport.width - margin - size.width);
  const y = clamp(idealY, top, bottom - size.height);
  return {
    x, y, width: size.width, height: size.height, side: chosen, flipped,
    clamped: target !== anchor || x !== idealX || y !== idealY,
    hidden: !visible,
  };
}
