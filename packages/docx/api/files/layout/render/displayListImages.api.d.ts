// @betteroffice/docx declaration layout/render/displayListImages.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Locate an image primitive in a built DisplayList by its document position.
 *
 * The canvas image-selection overlay needs the selected image's page-local
 * rect to place its resize handles, and on the canvas path the painted DOM
 * image (`.layout-run-image`) is parked/hidden — its rect is unusable. The
 * display list carries every inline image as an `ImagePrimitive` stamped with
 * `docStart` (= the image atom's document position, the same value the DOM path reads
 * from `data-doc-start`), so the primitive found here is the geometry source
 * that replaces `imageInfo.element.getBoundingClientRect()`.
 *
 * Body primitives only: header/footer image primitives carry doc positions in a
 * different doc (the HF editor), so a body image pmPos must never match one.
 *
 */
import type { DisplayList, ImagePrimitive } from './displayList';
import type { DisplayListQueries } from './displayListQueries';
export type DisplayListImageRegion = 'body' | 'header' | 'footer';
export interface LocatedImagePrimitive {
    primitive: ImagePrimitive;
    /** page the primitive paints on (page-local px are relative to this page) */
    pageIndex: number;
    /** Document region containing the image. Undefined reads as body. */
    region?: DisplayListImageRegion;
    /** Header/footer relationship id when `region` is not body. */
    rId?: string;
}
/**
 * First body image primitive whose `docStart` equals `pos`, or null. `pos` is
 * the image node's document position (a node selection's `from`, or the left
 * edge of a one-atom text selection).
 */
export declare function findImagePrimitiveByDocPos(list: DisplayList, pos: number, region?: DisplayListImageRegion, rId?: string): LocatedImagePrimitive | null;
export interface LocatedDrawingRect {
    pageIndex: number;
    x: number;
    y: number;
    width: number;
    height: number;
}
/**
 * The drawn frame box of the drawing atom at `pos`, one per page that paints
 * it: the placed extent rectangle, unrotated, without effect extent. A chart
 * answers its placed box; pictures answer their frames, several pictures of
 * one drawing their union; a group of shapes the union of its shapes.
 */
export declare function findDrawingRectsByDocPos(list: DisplayList, pos: number, region?: DisplayListImageRegion, rId?: string): LocatedDrawingRect[];
/** Topmost positioned image painted under a page-local point, or null. */
export declare function findImagePrimitiveAtPoint(list: DisplayList, pageIndex: number, x: number, y: number, region?: DisplayListImageRegion, rId?: string): LocatedImagePrimitive | null;
/**
 * Capture an inline body's authored anchor offsets from display-list geometry.
 * Horizontal position is relative to the page content/column box; vertical
 * position is relative to the containing paragraph fragment. Header/footer
 * images deliberately return undefined until their paragraph/content-origin
 * metadata is part of the display-list contract.
 */
export declare function captureInlinePositionEmuFromDisplayList(queries: DisplayListQueries, pos: number): {
    horizontalEmu: number;
    verticalEmu: number;
} | undefined;
