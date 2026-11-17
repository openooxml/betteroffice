// @betteroffice/docx declaration layout/render/canvasPageMetrics.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Shared vertical geometry for canvas page hosts. */
import type { DisplayListQueries } from './displayListQueries';
/** vertical gap between canvas pages (flex `gap`) */
export declare const CANVAS_PAGE_GAP_PX = 16;
/** padding around the canvas pages column */
export declare const CANVAS_PAGES_PADDING_PX = 24;
/**
 * How page geometry maps to device pixels. `rasterScale` is the canvas
 * backing scale (`devicePixelRatio * zoom`); `unitScale` is the device pixels
 * per unit of the column the boxes are laid out in (`rasterScale` when a
 * transform applies the zoom, `devicePixelRatio` when the column is sized at
 * the zoom).
 */
export interface CanvasPageScale {
    rasterScale: number;
    unitScale: number;
}
/** A page box in column units. */
export interface CanvasPageBox {
    top: number;
    width: number;
    height: number;
}
/** How a page column is stacked. */
export interface CanvasPageStackOptions {
    /** Snaps every box to whole device pixels at this scale; omitted, boxes keep the exact page size. */
    scale?: CanvasPageScale;
    /** Column units per page unit when no `scale` is given (a column sized at the zoom). */
    zoom?: number;
    padding?: number;
}
/** Page length in column units; with a scale, it covers exactly its canvas backing pixels. */
export declare function canvasPageLength(length: number, scale?: CanvasPageScale, zoom?: number): number;
/** A column length, rounded to whole device pixels when a scale is given. */
export declare function snapCanvasLength(length: number, scale?: CanvasPageScale): number;
/**
 * Page boxes stacked down a column. With a scale, each box spans its canvas
 * backing pixels exactly and every top falls on a whole device pixel, so page
 * bitmaps are shown 1:1 instead of resampled. Every surface that converts
 * between page and column coordinates reads its tops from here.
 */
export declare function canvasPageStack(pages: readonly {
    width: number;
    height: number;
}[], { scale, zoom, padding }?: CanvasPageStackOptions): CanvasPageBox[];
/**
 * Top Y of each canvas page inside the `.canvas-pages` host, derived from the
 * display list's own page heights. `tops[i] + pageLocalY` converts a
 * display-list rect into host coordinates.
 */
export declare function canvasPageTops(queries: DisplayListQueries): number[];
