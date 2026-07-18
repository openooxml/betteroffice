/**
 * Renderer-neutral geometry helpers for display-list primitives.
 *
 * These helpers are shared by canvas replay, interaction queries, the a11y
 * mirror, and the temporary painter/display-list differential harness. They
 * must not depend on DOM layout: every returned rectangle is page-local px.
 */

import type {
  DisplayPrimitive,
  GlyphRunPrimitive,
  LinePrimitive,
  TextRunPrimitive,
} from './displayList';

export interface GeoRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

// v0 text primitives carry pen origin + baseline + advance but no vertical
// extent; the canonical text box is derived from the CSS font size with these
// ratios. Both the mirror and display-list queries use this helper.
const TEXT_ASCENT_RATIO = 0.8;
const TEXT_LINE_RATIO = 1.2;
// GlyphRun glyphs carry their real pen advance, so the run's right edge is the
// trailing glyph's `x + advance`. Pre-shaper display lists omitted advances;
// retain their deterministic approximation for compatibility.
const GLYPH_TRAILING_ADVANCE_EM = 0.6;

/** Parse the px size from a CSS font shorthand, with Word's 11pt default. */
export function fontSizePxFromShorthand(font: string): number {
  const match = /(\d+(?:\.\d+)?)px/.exec(font);
  const size = match ? Number(match[1]) : NaN;
  return Number.isFinite(size) && size > 0 ? size : (11 * 96) / 72;
}

/** Canonical page-local box for a text-run primitive. */
export function textRunRect(
  p: Pick<TextRunPrimitive, 'x' | 'baselineY' | 'width' | 'font'>
): GeoRect {
  const size = fontSizePxFromShorthand(p.font);
  return {
    x: p.x,
    y: p.baselineY - size * TEXT_ASCENT_RATIO,
    w: p.width,
    h: size * TEXT_LINE_RATIO,
  };
}

/** Canonical page-local box for a shaped glyph-run primitive. */
export function glyphRunRect(p: Pick<GlyphRunPrimitive, 'glyphs' | 'size'>): GeoRect {
  const { glyphs, size } = p;
  if (glyphs.length === 0) {
    return { x: 0, y: 0, w: 0, h: size * TEXT_LINE_RATIO };
  }
  let minX = Infinity;
  let maxX = -Infinity;
  let realRight = -Infinity;
  let allHaveAdvance = true;
  for (const glyph of glyphs) {
    if (glyph.x < minX) minX = glyph.x;
    if (glyph.x > maxX) maxX = glyph.x;
    if (glyph.advance === undefined) {
      allHaveAdvance = false;
    } else {
      realRight = Math.max(realRight, glyph.x + glyph.advance);
    }
  }
  const width = allHaveAdvance
    ? Math.max(realRight - minX, 0)
    : maxX -
      minX +
      (glyphs.length > 1 ? (maxX - minX) / (glyphs.length - 1) : size * GLYPH_TRAILING_ADVANCE_EM);
  return {
    x: minX,
    y: glyphs[0].y - size * TEXT_ASCENT_RATIO,
    w: width,
    h: size * TEXT_LINE_RATIO,
  };
}

/** Canonical stroke-expanded page-local box for a line primitive. */
export function lineRect(
  p: Pick<LinePrimitive, 'x1' | 'y1' | 'x2' | 'y2' | 'strokeWidth'>
): GeoRect {
  const minX = Math.min(p.x1, p.x2);
  const minY = Math.min(p.y1, p.y2);
  const dx = Math.abs(p.x2 - p.x1);
  const dy = Math.abs(p.y2 - p.y1);
  const sw = p.strokeWidth;
  if (dy === 0) return { x: minX, y: minY - sw / 2, w: dx, h: sw };
  if (dx === 0) return { x: minX - sw / 2, y: minY, w: sw, h: dy };
  return { x: minX - sw / 2, y: minY - sw / 2, w: dx + sw, h: dy + sw };
}

/** Page-local box for any display primitive. */
export function displayPrimitiveRect(primitive: DisplayPrimitive): GeoRect {
  switch (primitive.kind) {
    case 'text':
      return textRunRect(primitive);
    case 'glyphRun':
      return glyphRunRect(primitive);
    case 'rect':
    case 'image':
    case 'shape':
    case 'decoration':
      return { x: primitive.x, y: primitive.y, w: primitive.w, h: primitive.h };
    case 'line':
      return lineRect(primitive);
  }
}
