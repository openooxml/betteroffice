/**
 * Floating-table placement geometry (`w:tblpPr`).
 *
 * Resolves one page-space rectangle — position plus wrap exclusion — for a
 * floating table from its anchor/offset properties. Consumed by the flow
 * pre-pass (`layout/flow/floatingTable.ts` re-exports this surface) and
 * mirrored natively by the Rust engine; kept in `layout/pagination` so it
 * stays on the pure seam (plain data in, plain data out).
 */

import type { TableBlock, PageMargins } from './types';

export interface FloatingTablePlacementInput {
  block: TableBlock;
  tableWidth: number;
  tableHeight: number;
  pageSize: { w: number; h: number };
  margins: PageMargins;
  pageNumber: number;
  columnX: number;
  columnWidth: number;
  penY: number;
}

export interface FloatingTablePlacement {
  x: number;
  y: number;
  exclusion: { left: number; top: number; right: number; bottom: number };
  hasExplicitY: boolean;
}

function finite(value: number | undefined, fallback = 0): number {
  return value != null && Number.isFinite(value) ? value : fallback;
}

/** Resolve one page-space `tblpPr` rectangle for placement and exclusion. */
export function resolveFloatingTablePlacement({
  block,
  tableWidth,
  tableHeight,
  pageSize,
  margins,
  pageNumber,
  columnX,
  columnWidth,
  penY,
}: FloatingTablePlacementInput): FloatingTablePlacement {
  const floating = block.floating;
  const horizontal = floating?.horzAnchor ?? 'margin';
  const vertical = floating?.vertAnchor ?? 'text';
  const hStart = horizontal === 'page' ? 0 : horizontal === 'text' ? columnX : margins.left;
  const hEnd =
    horizontal === 'page'
      ? pageSize.w
      : horizontal === 'text'
        ? columnX + columnWidth
        : pageSize.w - margins.right;
  const vStart = vertical === 'page' ? 0 : vertical === 'text' ? penY : margins.top;
  const vEnd = vertical === 'page' ? pageSize.h : pageSize.h - margins.bottom;
  const insideIsStart = pageNumber % 2 === 1;

  let x = hStart;
  if (floating?.tblpX != null && Number.isFinite(floating.tblpX)) {
    x = hStart + floating.tblpX;
  } else {
    const spec = floating?.tblpXSpec;
    const startAligned =
      spec === 'left' ||
      (spec === 'inside' && insideIsStart) ||
      (spec === 'outside' && !insideIsStart);
    const endAligned =
      spec === 'right' ||
      (spec === 'inside' && !insideIsStart) ||
      (spec === 'outside' && insideIsStart);
    if (spec === 'center') x = hStart + (hEnd - hStart - tableWidth) / 2;
    else if (endAligned) x = hEnd - tableWidth;
    else if (startAligned) x = hStart;
    else if (block.justification === 'center') x = hStart + (hEnd - hStart - tableWidth) / 2;
    else if (block.justification === 'right') x = hEnd - tableWidth;
  }

  let y = penY;
  let hasExplicitY = false;
  if (floating?.tblpY != null && Number.isFinite(floating.tblpY)) {
    y = vStart + floating.tblpY;
    hasExplicitY = true;
  } else if (floating?.tblpYSpec && floating.tblpYSpec !== 'inline') {
    hasExplicitY = true;
    const spec = floating.tblpYSpec;
    if (spec === 'center') y = vStart + (vEnd - vStart - tableHeight) / 2;
    else if (spec === 'bottom' || spec === 'outside') y = vEnd - tableHeight;
    else y = vStart;
  }

  return {
    x,
    y,
    hasExplicitY,
    exclusion: {
      left: x - finite(floating?.leftFromText),
      top: y - finite(floating?.topFromText),
      right: x + tableWidth + finite(floating?.rightFromText),
      bottom: y + tableHeight + finite(floating?.bottomFromText),
    },
  };
}
