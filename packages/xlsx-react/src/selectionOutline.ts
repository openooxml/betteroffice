/**
 * Geometry for the cell-selection marquee.
 *
 * The canvas snaps its unstyled grid strokes to the CSS pixel grid: a one-pixel
 * line on a boundary at `b` covers `[b - 1, b]`. The selection border has to sit
 * on that same grid, or it reads as inside the cell on one edge and outside on
 * the other whenever the zoom makes the boundary fractional.
 *
 * `marqueeBox` returns the element box for a ring drawn with
 * `box-shadow: inset 0 0 0 Npx …`. An even N grows the box by N/2 per side, so
 * the ring straddles each snapped boundary.
 */
export interface PixelRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface MarqueeBox {
  left: number;
  top: number;
  width: number;
  height: number;
  /** The snapped interior, for fills and hit tests. */
  inset: PixelRect;
}

/** The light wash under a selected range, with the active cell left clear. */
export const SELECTION_FILL = 'rgba(33, 115, 70, 0.12)';

export const MARQUEE_COLOR = '#217346';
/** The range border: one pixel, sitting on the selection's own snapped edges. */
export const MARQUEE_WIDTH = 1;
/** The active cell border: two pixels, straddling the cell's snapped edges. */
export const ACTIVE_CELL_WIDTH = 2;

export function marqueeBox(rect: PixelRect, width = MARQUEE_WIDTH): MarqueeBox {
  const x0 = Math.round(rect.x);
  const y0 = Math.round(rect.y);
  const x1 = Math.round(rect.x + rect.w);
  const y1 = Math.round(rect.y + rect.h);
  const w = Math.max(0, x1 - x0);
  const h = Math.max(0, y1 - y0);
  if (width % 2 === 1) {
    return {
      left: x0,
      top: y0,
      width: w,
      height: h,
      inset: { x: x0, y: y0, w, h },
    };
  }
  const outside = width / 2;
  return {
    left: x0 - outside,
    top: y0 - outside,
    width: w + 2 * outside,
    height: h + 2 * outside,
    inset: { x: x0, y: y0, w, h },
  };
}

/**
 * The wash for a selection, split around the active cell so the cell the user
 * is on stays unwashed — the first box of the selection keeps no background,
 * the rest does. Returns one rect when the hole is absent, outside, or when the
 * range is a single cell.
 */
export function marqueeFills(range: PixelRect, hole: PixelRect | null): PixelRect[] {
  const outer = marqueeBox(range).inset;
  if (!hole || outer.w === 0 || outer.h === 0) return outer.w && outer.h ? [outer] : [];
  const inner = marqueeBox(hole).inset;
  const right = outer.x + outer.w;
  const bottom = outer.y + outer.h;
  const innerRight = inner.x + inner.w;
  const innerBottom = inner.y + inner.h;
  const clipped = {
    x: Math.max(outer.x, inner.x),
    y: Math.max(outer.y, inner.y),
    w: Math.max(0, Math.min(right, innerRight) - Math.max(outer.x, inner.x)),
    h: Math.max(0, Math.min(bottom, innerBottom) - Math.max(outer.y, inner.y)),
  };
  if (clipped.w === 0 || clipped.h === 0) return [outer];
  const strips: PixelRect[] = [
    { x: outer.x, y: outer.y, w: outer.w, h: clipped.y - outer.y },
    { x: outer.x, y: clipped.y + clipped.h, w: outer.w, h: bottom - (clipped.y + clipped.h) },
    { x: outer.x, y: clipped.y, w: clipped.x - outer.x, h: clipped.h },
    { x: clipped.x + clipped.w, y: clipped.y, w: right - (clipped.x + clipped.w), h: clipped.h },
  ];
  return strips.filter((strip) => strip.w > 0 && strip.h > 0);
}
