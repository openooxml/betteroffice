// @betteroffice/docx declaration layout/render/displayListGeometry.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Renderer-neutral page-local geometry for display-list primitives. */
import type { DisplayPrimitive, GlyphRunPrimitive, LinePrimitive, TextRunPrimitive } from './displayList';
export interface GeoRect {
    x: number;
    y: number;
    w: number;
    h: number;
}
/** Parse the px size from a CSS font shorthand, with Word's 11pt default. */
export declare function fontSizePxFromShorthand(font: string): number;
/** Canonical page-local box for a text-run primitive. */
export declare function textRunRect(p: Pick<TextRunPrimitive, 'x' | 'baselineY' | 'width' | 'font'>): GeoRect;
/** Canonical page-local box for a shaped glyph-run primitive. */
export declare function glyphRunRect(p: Pick<GlyphRunPrimitive, 'glyphs' | 'size'>): GeoRect;
/** Canonical stroke-expanded page-local box for a line primitive. */
export declare function lineRect(p: Pick<LinePrimitive, 'x1' | 'y1' | 'x2' | 'y2' | 'strokeWidth'>): GeoRect;
/** Page-local box for any display primitive. */
export declare function displayPrimitiveRect(primitive: DisplayPrimitive): GeoRect;
/** Whether the primitive's paint clips contain a page-local point. */
export declare function clipPaintsPoint(primitive: DisplayPrimitive, x: number, y: number): boolean;
