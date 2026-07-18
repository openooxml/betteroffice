/**
 * Vertical layout of a table cell's stacked blocks.
 *
 * Single source of truth for *where the lines of a cell's content sit*, used by
 * the row-break paginator (`tableRowBreak`), the selection-rect mapper
 * (`selectionRects`), and — by construction — matching what the painter
 * (`renderCellContent`) renders. Adjacent paragraphs' before/after spacing
 * collapses to the larger of the two (CSS margin-collapse / the body
 * paginator's rule); a paragraph's lines stack from its content top with no
 * before/after between them. Keeping these consumers on one model is what
 * prevents page breaks and selection highlights from drifting off the rendered
 * lines.
 */

import type { LayoutBlock, BlockExtent } from './types';

export interface CellContentLayout {
  /** Per block, the top y of each line (relative to `startY`). Atomic/non-paragraph blocks → []. */
  lineTops: number[][];
  /**
   * All line bottoms in document order, plus one entry per atomic block (its
   * bottom) — the clean break points for the paginator.
   */
  flatBottoms: number[];
  /** Total stacked height incl. the last block's trailing space-after. */
  contentHeight: number;
  /** Line/atomic interiors that a page boundary must never cut through. */
  unbreakableRanges: Array<{ top: number; bottom: number }>;
}

/**
 * Compute the collapsed vertical layout of a cell's blocks starting at `startY`.
 */
export function layoutCellContent(
  blocks: readonly LayoutBlock[] | undefined,
  blockMeasures: readonly BlockExtent[] | undefined,
  startY: number
): CellContentLayout {
  const lineTops: number[][] = [];
  const flatBottoms: number[] = [];
  const unbreakableRanges: Array<{ top: number; bottom: number }> = [];
  let y = startY;
  let prevAfter = 0;
  const n = blockMeasures?.length ?? 0;

  for (let i = 0; i < n; i++) {
    const measure = blockMeasures![i];
    const block = blocks?.[i];
    if (block?.kind === 'paragraph' && measure?.kind === 'paragraph') {
      const spacing = block.attrs?.spacing;
      y += Math.max(prevAfter, spacing?.before ?? 0);
      const tops: number[] = [];
      for (const line of measure.lines) {
        y += line.floatSkipBefore ?? 0;
        tops.push(y);
        const top = y;
        y += line.lineHeight;
        flatBottoms.push(y);
        unbreakableRanges.push({ top, bottom: y });
      }
      lineTops.push(tops);
      prevAfter = spacing?.after ?? 0;
    } else if (
      measure?.kind === 'image' ||
      measure?.kind === 'shape' ||
      measure?.kind === 'chart' ||
      measure?.kind === 'textBox'
    ) {
      // Atomic bbox block (break only at its bottom).
      y += prevAfter;
      const top = y;
      y += measure.height;
      lineTops.push([]);
      flatBottoms.push(y);
      unbreakableRanges.push({ top, bottom: y });
      prevAfter = 0;
    } else if (measure && 'totalHeight' in measure && typeof measure.totalHeight === 'number') {
      // Nested table / non-paragraph: one atomic block (break only at its bottom).
      y += prevAfter;
      const top = y;
      y += measure.totalHeight;
      lineTops.push([]);
      flatBottoms.push(y);
      unbreakableRanges.push({ top, bottom: y });
      prevAfter = 0;
    } else {
      lineTops.push([]);
    }
  }

  // The painter renders the final block's trailing space-after as paddingBottom.
  return { lineTops, flatBottoms, contentHeight: y - startY + prevAfter, unbreakableRanges };
}
