/**
 * Locate an image primitive in a built DisplayList by its document position.
 *
 * The canvas image-selection overlay needs the selected image's page-local
 * rect to place its resize handles, and on the canvas path the painted DOM
 * image (`.layout-run-image`) is parked/hidden — its rect is unusable. The
 * display list carries every inline image as an `ImagePrimitive` stamped with
 * `docStart` (= the image atom's PM position, the same value the DOM path reads
 * from `data-doc-start`), so the primitive found here is the geometry source
 * that replaces `imageInfo.element.getBoundingClientRect()`.
 *
 * Body primitives only: header/footer image primitives carry doc positions in a
 * different PM doc (the HF editor), so a body image pmPos must never match one.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { DisplayList, DisplayPage, ImagePrimitive } from './displayList';
import type { DisplayListQueries } from './displayListQueries';
import { pixelsToEmu } from '../../utils/units';

export type DisplayListImageRegion = 'body' | 'header' | 'footer';

export interface LocatedImagePrimitive {
  primitive: ImagePrimitive;
  /** page the primitive paints on (page-local px are relative to this page) */
  pageIndex: number;
  /** PM-document region containing the image. Undefined reads as body. */
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
 * the image node's PM position (a NodeSelection's `from`, or the left edge of a
 * one-atom TextSelection).
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

/** Topmost positioned image under a page-local point, or null. */
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
