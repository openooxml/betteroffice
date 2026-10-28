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

import type { DisplayList, DisplayPage, ImagePrimitive } from './displayList';
import type { DisplayListQueries } from './displayListQueries';
import { clipPaintsPoint } from './displayListGeometry';
import { pixelsToEmu } from '../../utils/units';

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

function imagesForRegion(
  page: DisplayPage,
  region: DisplayListImageRegion,
  rId?: string
): ImagePrimitive[] {
  if (region === 'body') {
    return page.primitives.filter(
      (primitive): primitive is ImagePrimitive => primitive.kind === 'image'
    );
  }
  const band = region === 'header' ? page.header : page.footer;
  if (!band || (rId !== undefined && band.rId !== rId)) return [];
  return band.primitives.filter(
    (primitive): primitive is ImagePrimitive => primitive.kind === 'image'
  );
}

/**
 * First body image primitive whose `docStart` equals `pos`, or null. `pos` is
 * the image node's document position (a node selection's `from`, or the left
 * edge of a one-atom text selection).
 */
export function findImagePrimitiveByDocPos(
  list: DisplayList,
  pos: number,
  region: DisplayListImageRegion = 'body',
  rId?: string
): LocatedImagePrimitive | null {
  for (const page of list.pages) {
    for (const primitive of imagesForRegion(page, region, rId)) {
      if (primitive.docStart === pos) {
        return {
          primitive,
          pageIndex: page.pageIndex,
          ...(region !== 'body' ? { region, rId } : {}),
        };
      }
    }
  }
  return null;
}

export interface LocatedDrawingRect {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

function drawingPrimitivesForRegion(
  page: DisplayPage,
  region: DisplayListImageRegion,
  rId?: string
): DisplayPrimitive[] {
  if (region === 'body') return page.primitives ?? [];
  const band = region === 'header' ? page.header : page.footer;
  if (!band || (rId !== undefined && band.rId !== rId)) return [];
  return band.primitives ?? [];
}

type DisplayPrimitive = DisplayPage['primitives'][number];

type Box = { x: number; y: number; w: number; h: number };

function unionBox(boxes: readonly Box[]): Box | null {
  if (!boxes.length) return null;
  const x = Math.min(...boxes.map((box) => box.x));
  const y = Math.min(...boxes.map((box) => box.y));
  const right = Math.max(...boxes.map((box) => box.x + box.w));
  const bottom = Math.max(...boxes.map((box) => box.y + box.h));
  return { x, y, w: right - x, h: bottom - y };
}

function primitiveBox(primitive: DisplayPrimitive): Box | null {
  if (!('x' in primitive) || !('y' in primitive)) return null;
  const w = 'w' in primitive && typeof primitive.w === 'number' ? primitive.w : 0;
  const h = 'h' in primitive && typeof primitive.h === 'number' ? primitive.h : 0;
  return { x: primitive.x, y: primitive.y, w, h };
}

function chartFrame(primitives: readonly DisplayPrimitive[]): Box | null {
  for (const primitive of primitives) {
    const frame = primitive.chart?.frame;
    if (frame && [frame.x, frame.y, frame.w, frame.h].every((value) => typeof value === 'number')) {
      return { x: frame.x!, y: frame.y!, w: frame.w!, h: frame.h! };
    }
  }
  return null;
}

/**
 * The drawn frame box of the drawing atom at `pos`, one per page that paints
 * it: the placed extent rectangle, unrotated, without effect extent. A chart
 * answers its placed box; pictures answer their frames, several pictures of
 * one drawing their union; a group of shapes the union of its shapes.
 */
export function findDrawingRectsByDocPos(
  list: DisplayList,
  pos: number,
  region: DisplayListImageRegion = 'body',
  rId?: string
): LocatedDrawingRect[] {
  const rects: LocatedDrawingRect[] = [];
  for (const page of list.pages) {
    if (page.unbuilt) continue;
    const primitives = drawingPrimitivesForRegion(page, region, rId).filter(
      (primitive) => primitive.docStart === pos
    );
    const images = primitives.filter(
      (primitive): primitive is ImagePrimitive => primitive.kind === 'image'
    );
    const box =
      chartFrame(primitives) ??
      unionBox(
        images.map((image) => ({
          x: image.contentFrame?.x ?? image.x,
          y: image.contentFrame?.y ?? image.y,
          w: image.contentFrame?.w ?? image.w,
          h: image.contentFrame?.h ?? image.h,
        }))
      ) ??
      unionBox(
        primitives
          .filter(
            (primitive) =>
              primitive.kind === 'shape' || (primitive.chart !== undefined && 'w' in primitive)
          )
          .map(primitiveBox)
          .filter((box): box is Box => box !== null)
      ) ??
      unionBox(
        primitives
          .filter((primitive) => primitive.kind === 'rect' && primitive.docEnd === pos + 1)
          .map(primitiveBox)
          .filter((box): box is Box => box !== null)
      );
    if (box) rects.push({ pageIndex: page.pageIndex, x: box.x, y: box.y, width: box.w, height: box.h });
  }
  return rects;
}

/** Topmost positioned image painted under a page-local point, or null. */
export function findImagePrimitiveAtPoint(
  list: DisplayList,
  pageIndex: number,
  x: number,
  y: number,
  region: DisplayListImageRegion = 'body',
  rId?: string
): LocatedImagePrimitive | null {
  const page = list.pages[pageIndex];
  if (!page) return null;
  const images = imagesForRegion(page, region, rId);
  // Reverse paint order: the last image is visually on top.
  for (let index = images.length - 1; index >= 0; index--) {
    const primitive = images[index];
    if (x < primitive.x || x > primitive.x + primitive.w) continue;
    if (y < primitive.y || y > primitive.y + primitive.h) continue;
    if (!clipPaintsPoint(primitive, x, y)) continue;
    return {
      primitive,
      pageIndex: page.pageIndex,
      ...(region !== 'body' ? { region, rId } : {}),
    };
  }
  return null;
}

/**
 * Capture an inline body's authored anchor offsets from display-list geometry.
 * Horizontal position is relative to the page content/column box; vertical
 * position is relative to the containing paragraph fragment. Header/footer
 * images deliberately return undefined until their paragraph/content-origin
 * metadata is part of the display-list contract.
 */
export function captureInlinePositionEmuFromDisplayList(
  queries: DisplayListQueries,
  pos: number
): { horizontalEmu: number; verticalEmu: number } | undefined {
  const image = queries.imageByPos(pos);
  if (!image) return undefined;
  const content = queries.contentBounds(image.pageIndex);
  const paragraph = queries.paragraphRects(pos).find((rect) => rect.pageIndex === image.pageIndex);
  if (!content || !paragraph) return undefined;
  return {
    horizontalEmu: Math.round(pixelsToEmu(image.primitive.x - content.x)),
    verticalEmu: Math.round(pixelsToEmu(image.primitive.y - paragraph.y)),
  };
}
