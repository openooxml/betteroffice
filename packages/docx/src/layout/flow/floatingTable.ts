/**
 * Floating-table compatibility helpers.
 *
 * A DOCX table with `w:tblpPr` is a positioned ("floating") table — text wraps
 * around it. But contract templates routinely make a full-width form table
 * floating (often a paste artifact), leaving NO horizontal room for text on
 * either side. Word and Google Docs treat such a table as a normal in-flow
 * block: it paginates across pages and following content flows beneath it.
 *
 * The paginator now preserves authored floating status: block-like floats
 * advance flow beneath their exclusion, and unavoidable tall floats fragment
 * at safe row/line boundaries.
 *
 * The predicate mirrors the geometry in `layoutFloatingTable` (the existing
 * block-like cursor-advance) and `extractFloatingTableZone`; keep the three in
 * sync. Narrow floating tables — where text genuinely wraps beside them — fall
 * below the threshold and stay floating.
 */

import type { LayoutBlock, TableBlock } from '../pagination/types';
import { MIN_WRAP_SEGMENT_WIDTH } from '../measure/floatingZones';
import { resolveTableTotalWidthPx } from '../pagination/tableWidthUtils';
export { resolveFloatingTablePlacement } from '../pagination/floatingTablePlacement';
export type {
  FloatingTablePlacement,
  FloatingTablePlacementInput,
} from '../pagination/floatingTablePlacement';

/**
 * True when a floating table is effectively full-width — it leaves less than
 * {@link MIN_WRAP_SEGMENT_WIDTH} of usable text room on BOTH sides, so no text
 * can wrap beside it. Returns false for non-floating tables and for narrow
 * floats that leave real wrap room.
 *
 * @internal
 */
export function isBlockLikeFloatingTable(block: TableBlock, contentWidth: number): boolean {
  const floating = block.floating;
  if (!floating) return false;

  const tableWidth = resolveTableTotalWidthPx(block, contentWidth);

  // Content-relative X of the table's left edge — same resolution order as
  // `layoutFloatingTable` / `extractFloatingTableZone`.
  let x = 0;
  if (floating.tblpX !== undefined) {
    x = floating.tblpX;
  } else if (floating.tblpXSpec) {
    if (floating.tblpXSpec === 'left' || floating.tblpXSpec === 'inside') {
      x = 0;
    } else if (floating.tblpXSpec === 'right' || floating.tblpXSpec === 'outside') {
      x = contentWidth - tableWidth;
    } else if (floating.tblpXSpec === 'center') {
      x = (contentWidth - tableWidth) / 2;
    }
  } else if (block.justification === 'center') {
    x = (contentWidth - tableWidth) / 2;
  } else if (block.justification === 'right') {
    x = contentWidth - tableWidth;
  }

  const leftFromText = floating.leftFromText ?? 0;
  const rightFromText = floating.rightFromText ?? 0;
  const leftSpace = x - leftFromText;
  const rightSpace = contentWidth - (x + tableWidth) - rightFromText;

  return leftSpace < MIN_WRAP_SEGMENT_WIDTH && rightSpace < MIN_WRAP_SEGMENT_WIDTH;
}

/**
 * Legacy no-op kept so external callers compiled against the old helper do not
 * break. Authored `floating` is never cleared.
 *
 * @internal
 */
export function demoteBlockLikeFloatingTables(
  _blocks: LayoutBlock[],
  _blockWidths: number[],
  _fallbackWidth: number
): void {
  // Retained as a compatibility API for callers compiled against the old
  // pipeline. Authored tblpPr status is now authoritative: wide floats keep
  // their anchor and the paginator advances the pen when no side segment is
  // usable; unavoidable tall floats fragment safely instead of overflowing.
}
