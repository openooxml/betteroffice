/**
 * Which pointer cursor the canvas pages paint, from one engine hit.
 *
 * `target` comes from `hit_test_regions` (`crates/docx-layout/src/hit.rs`) and
 * names what a click would act on; nothing here re-derives geometry. This maps
 * it onto the editor's modes, as the DOM painter's stylesheet also did for a
 * read-only document, for the body behind an open header/footer editor, and
 * for a picture as a select target.
 *
 * An idle header/footer band deliberately DIFFERS from that stylesheet, which
 * gave the whole band a hand: here it reads as text over its own runs, since a
 * single click on an idle band is inert and only a double-click opens it for
 * editing at that word. Over a note area the same reading is literal — a
 * single click there opens the note and puts the caret under the pointer.
 *
 * A wasm build predating `target` omits it, which reads as "not text".
 */

import type { DisplayListRegionHit } from '@betteroffice/docx/layout/render';

import { hitBelongsToPart, isNoteAreaHit, type PartEdit } from './partEdit';

export type CanvasHoverCursor = 'default' | 'text';

export interface CanvasHoverMode {
  readOnly: boolean;
  /** the non-body part open for editing, if any */
  partEdit?: PartEdit | null;
}

export function canvasHoverCursor(
  mode: CanvasHoverMode,
  hit: DisplayListRegionHit | null
): CanvasHoverCursor {
  if (mode.readOnly || !hit) return 'default';
  if (mode.partEdit) {
    // only the open part accepts typing; the body and every sibling are inert
    if (!hitBelongsToPart(mode.partEdit, hit)) return 'default';
    // the whole part is typeable once open, so an absent target still types
    return hit.target === 'image' ? 'default' : 'text';
  }
  if (isNoteAreaHit(hit)) return hit.target === 'image' ? 'default' : 'text';
  return hit.target === 'text' ? 'text' : 'default';
}

/**
 * Stamps the cursor on the pages host, skipping repeats. `cursor` is
 * inherited, so one write covers every page, and affordances carrying their
 * own (SDT widgets, the image and table overlays) still win. The host is part
 * of the guard, so a replaced one is re-stamped. Returns whether it wrote.
 */
export function createCursorPainter(): (
  host: HTMLElement | null,
  cursor: CanvasHoverCursor
) => boolean {
  let painted: { host: HTMLElement; cursor: CanvasHoverCursor } | null = null;
  return (host, cursor) => {
    if (!host) return false;
    if (painted?.host === host && painted.cursor === cursor) return false;
    painted = { host, cursor };
    host.style.cursor = cursor;
    return true;
  };
}
