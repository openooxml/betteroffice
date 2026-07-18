/**
 * Header / Footer Layout Utilities
 *
 * The header/footer rendering pipeline lives here so any rendering adapter
 * (React, Vue, etc.) can share the conversion logic and just supply its
 * platform-specific {@link MeasureBlocksFn}. Mirrors the footnote pipeline
 * in `footnoteLayout.ts`.
 *
 * Pipeline:
 *   HF.content → headerFooterToProseDoc → toLayoutBlocks
 *     → measureBlocks (caller-supplied, Canvas-aware)
 *     → HeaderFooterContent (blocks, measures, height, visualTop/Bottom)
 *
 * The render side uses the normalized block list so paint and measurement stay
 * in lockstep. Visual-bounds calculation still inspects the original block
 * list because floating images can paint above/below the nominal flow box even
 * when they do not contribute to flow height.
 */

import type {
  LayoutBlock,
  ImageRun,
  BlockExtent,
  PageMargins,
  TableBlock,
} from '../pagination/types';
import type { HeaderFooter, StyleDefinitions, Theme } from '../../types/document';
import type { HeaderFooterContent } from '../pagination/types';
import { emuToPixels } from '../../utils/units';
import type { MeasureBlocksFn } from './footnoteLayout';
import type { DisplayListQueries, DisplayListRect } from '../render/displayListQueries';

// ============================================================================
// 1. Page-level metrics passed in by the caller
// ============================================================================

export type HeaderFooterMetrics = {
  section: 'header' | 'footer';
  pageSize: { w: number; h: number };
  margins: PageMargins;
};

// ============================================================================
// 2. Measurement-time block normalization
// ============================================================================
//
// Two transforms are applied to the LayoutBlock list before measurement/render:
//
// 1. **Strip style-inherited paragraph spacing** (#380) — Word visibly
//    does NOT honor inherited `spaceBefore` / `spaceAfter` (e.g. Normal's
//    default 8pt-after) inside the HF text frame. Inline `<w:spacing>`
//    set explicitly on the HF paragraph IS honored. The parser flags
//    inline spacing via `spacingExplicit.before` / `.after`; anything
//    not flagged was inherited from the style chain and is zeroed for
//    both measurement and painting.
//
// 2. **Zero trailing empty paragraph after a table** (#381) — OOXML
//    requires a trailing block-level element after the last `<w:tbl>`
//    in any block container, including `<w:hdr>` / `<w:ftr>`. Word
//    renders that empty paragraph as a zero-height anchor (just the
//    paragraph mark glyph) when it has no runs AND no authored visual
//    content (no paragraph borders, no explicit spacing). We mark its
//    measure with `suppressEmptyParagraphHeight` so the BLOCK survives
//    (click-to-position into the empty space below the table places
//    the cursor in the trailing paragraph, matching Word) but the
//    measure returns zero height. Empty paragraphs with authored
//    `pBdr` (e.g. a horizontal rule under the header) or
//    `spacingExplicit` are NOT suppressed — they exist for their
//    visual side effect, not just as a structural anchor.

function hasAuthoredVisualContent(block: LayoutBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const attrs = block.attrs;
  if (!attrs) return false;
  if (attrs.borders?.top || attrs.borders?.bottom) return true;
  if (attrs.spacingExplicit?.before || attrs.spacingExplicit?.after) return true;
  return false;
}

export function normalizeHeaderFooterMeasureBlocks(blocks: LayoutBlock[]): LayoutBlock[] {
  return normalizeLayoutBlockArray(blocks);
}

function normalizeLayoutBlockArray(blocks: LayoutBlock[]): LayoutBlock[] {
  const trailingEmptyAfterTable = new Set<number>();
  for (let i = 1; i < blocks.length; i++) {
    const prev = blocks[i - 1];
    const cur = blocks[i];
    if (prev.kind !== 'table') continue;
    if (cur.kind !== 'paragraph') continue;
    if (cur.runs.length > 0) continue;
    if (hasAuthoredVisualContent(cur)) continue;
    trailingEmptyAfterTable.add(i);
  }

  return blocks.map((block, index) => {
    if (block.kind === 'table') {
      return normalizeTableBlock(block);
    }
    if (block.kind !== 'paragraph') return block;

    const isTrailingEmpty = trailingEmptyAfterTable.has(index);

    const explicit = block.attrs?.spacingExplicit;
    const hasResolvedBefore = block.attrs?.spacing?.before != null;
    const hasResolvedAfter = block.attrs?.spacing?.after != null;
    const beforeIsInherited = hasResolvedBefore && !explicit?.before;
    const afterIsInherited = hasResolvedAfter && !explicit?.after;
    const stripsSpacing = beforeIsInherited || afterIsInherited;

    if (!stripsSpacing && !isTrailingEmpty) return block;

    let attrs = block.attrs;
    if (stripsSpacing && attrs?.spacing) {
      attrs = {
        ...attrs,
        spacing: {
          ...attrs.spacing,
          before: explicit?.before ? attrs.spacing.before : undefined,
          after: explicit?.after ? attrs.spacing.after : undefined,
        },
      };
    }

    if (isTrailingEmpty) {
      attrs = { ...(attrs ?? {}), suppressEmptyParagraphHeight: true };
    }

    return { ...block, attrs };
  });
}

function normalizeTableBlock(block: TableBlock): TableBlock {
  let changed = false;
  const rows = block.rows.map((row) => {
    let rowChanged = false;
    const cells = row.cells.map((cell) => {
      const normalizedBlocks = normalizeLayoutBlockArray(cell.blocks);
      const cellChanged = normalizedBlocks.some(
        (normalizedBlock, idx) => normalizedBlock !== cell.blocks[idx]
      );
      if (!cellChanged) return cell;
      rowChanged = true;
      return { ...cell, blocks: normalizedBlocks };
    });
    if (!rowChanged) return row;
    changed = true;
    return { ...row, cells };
  });

  return changed ? { ...block, rows } : block;
}

// ============================================================================
// 3. Visual bounds (account for floating images that paint above/below the
//    nominal flow rectangle so HF clipping & shadow regions size correctly)
// ============================================================================

type PositionedAxis = {
  relativeTo?: string;
  posOffset?: number;
  align?: string;
  alignment?: string;
};

function getPositionAlignment(axis: PositionedAxis | undefined): string | undefined {
  return axis?.align ?? axis?.alignment;
}

export function resolveHeaderFooterVisualTop(
  run: ImageRun,
  paragraphY: number,
  flowHeight: number,
  metrics: HeaderFooterMetrics
): number {
  const flowTop =
    metrics.section === 'header'
      ? (metrics.margins.header ?? 48)
      : metrics.pageSize.h - (metrics.margins.footer ?? 48) - flowHeight;
  const vertical = run.position?.vertical;

  if (!vertical) {
    return paragraphY;
  }

  const align = getPositionAlignment(vertical);
  const offsetPx = vertical.posOffset !== undefined ? emuToPixels(vertical.posOffset) : undefined;

  if (vertical.relativeTo === 'page') {
    if (offsetPx !== undefined) return offsetPx - flowTop;
    if (align === 'top') return -flowTop;
    if (align === 'bottom') return metrics.pageSize.h - run.height - flowTop;
    if (align === 'center') return (metrics.pageSize.h - run.height) / 2 - flowTop;
  }

  if (vertical.relativeTo === 'margin') {
    const marginTop = metrics.margins.top;
    const marginHeight = metrics.pageSize.h - metrics.margins.top - metrics.margins.bottom;
    if (offsetPx !== undefined) return marginTop + offsetPx - flowTop;
    if (align === 'top') return marginTop - flowTop;
    if (align === 'bottom') return marginTop + marginHeight - run.height - flowTop;
    if (align === 'center') return marginTop + (marginHeight - run.height) / 2 - flowTop;
  }

  if (offsetPx !== undefined) {
    return paragraphY + offsetPx;
  }

  return paragraphY;
}

/**
 * Whether a header/footer block participates in the in-flow band height that
 * pushes the body margin.
 *
 * OOXML semantics: Word grows the header/footer band — and shifts body text —
 * based only on the story's in-flow content. A floating/anchored object
 * (`wp:anchor` DrawingML or an absolutely-positioned VML shape, e.g. a
 * full-page letterhead anchored to the page in a header) is removed from the
 * text flow and positioned on the page; it does NOT grow the band or push the
 * body. So only inline-flow blocks count here. Anchored image *runs* inside a
 * paragraph are likewise out of flow, but they don't contribute to the
 * paragraph's measured line height, so paragraphs need no special handling.
 *
 * @public
 */
export function contributesToHeaderFooterFlowHeight(block: LayoutBlock): boolean {
  switch (block.kind) {
    case 'paragraph':
    case 'table':
      return true;
    case 'image':
      // Inline images count; page/paragraph-anchored floats do not.
      return !block.anchor?.isAnchored;
    case 'shape':
    case 'chart':
      return true;
    case 'textBox':
      // Only genuinely inline text boxes count. 'float' (square/tight/through/
      // behind/inFront) and 'block' (topAndBottom) are positioned out of the
      // body's flow and must not push the body margin.
      return block.displayMode === undefined || block.displayMode === 'inline';
    default:
      return false; // sectionBreak / pageBreak / columnBreak
  }
}

function measureFlowHeight(measure: BlockExtent | undefined): number {
  if (!measure) return 0;
  if (measure.kind === 'paragraph') return measure.totalHeight;
  if (measure.kind === 'table') return measure.totalHeight;
  if (measure.kind === 'image') return measure.height;
  if (measure.kind === 'shape') return measure.height;
  if (measure.kind === 'chart') return measure.height;
  if (measure.kind === 'textBox') return measure.height;
  return 0;
}

export function calculateHeaderFooterVisualBounds(
  blocks: LayoutBlock[],
  measures: BlockExtent[],
  flowHeight: number,
  metrics: HeaderFooterMetrics
): { visualTop: number; visualBottom: number } {
  let visualTop = 0;
  // Accumulate the real extent from the blocks below. Do NOT seed with the
  // caller's `flowHeight` arg (it is the float-inclusive `totalHeight`): when a
  // floating box doesn't advance the cursor, seeding with the stacked total
  // would keep `visualBottom` artificially tall and the header container/hover
  // highlight would read taller than the painted content (#705/#729).
  let visualBottom = 0;
  let cursorY = 0;

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const measure = measures[i];
    if (!block || !measure) continue;

    if (block.kind === 'paragraph' && measure.kind === 'paragraph') {
      const paragraphStartY = cursorY;
      const paragraphBottomY = paragraphStartY + measure.totalHeight;
      visualTop = Math.min(visualTop, paragraphStartY);
      visualBottom = Math.max(visualBottom, paragraphBottomY);

      for (const run of block.runs) {
        if (run.kind !== 'image' || !run.position) continue;
        const runTop = resolveHeaderFooterVisualTop(run, paragraphStartY, flowHeight, metrics);
        visualTop = Math.min(visualTop, runTop);
        visualBottom = Math.max(visualBottom, runTop + run.height);
      }

      cursorY = paragraphBottomY;
    } else if (block.kind === 'table' && measure.kind === 'table') {
      const blockBottomY = cursorY + measure.totalHeight;
      visualTop = Math.min(visualTop, cursorY);
      visualBottom = Math.max(visualBottom, blockBottomY);
      cursorY = blockBottomY;
    } else if (block.kind === 'image' && measure.kind === 'image') {
      const blockBottomY = cursorY + measure.height;
      visualTop = Math.min(visualTop, cursorY);
      visualBottom = Math.max(visualBottom, blockBottomY);
      cursorY = blockBottomY;
    } else if (block.kind === 'shape' && measure.kind === 'shape') {
      const blockBottomY = cursorY + measure.height;
      visualTop = Math.min(visualTop, cursorY);
      visualBottom = Math.max(visualBottom, blockBottomY);
      cursorY = blockBottomY;
    } else if (block.kind === 'chart' && measure.kind === 'chart') {
      const blockBottomY = cursorY + measure.height;
      visualTop = Math.min(visualTop, cursorY);
      visualBottom = Math.max(visualBottom, blockBottomY);
      cursorY = blockBottomY;
    } else if (block.kind === 'textBox' && measure.kind === 'textBox') {
      const blockBottomY = cursorY + measure.height;
      visualTop = Math.min(visualTop, cursorY);
      visualBottom = Math.max(visualBottom, blockBottomY);
      // A floating text box is positioned, not in-flow: it extends the visual
      // bounds (so the band/container stays tall enough to show it) but does
      // NOT advance the cursor, mirroring the painter (renderHeaderFooterContent)
      // and floating tables. Otherwise the header container outgrows its actual
      // content and the hover highlight reads taller than the header (#705/#729).
      if (block.displayMode !== 'float') {
        cursorY = blockBottomY;
      }
    }
  }

  return { visualTop, visualBottom };
}

// ============================================================================
// 4. HeaderFooter → HeaderFooterContent (the public entry point)
// ============================================================================

export type ConvertHeaderFooterOptions = {
  styles?: StyleDefinitions | null;
  theme?: Theme | null;
  measureBlocks: MeasureBlocksFn;
  /** Authoritative yrs story source. */
  yrsStoryBlocks?: (storyId: string) => LayoutBlock[] | null;
  /**
   * `w:defaultTabStop` (twips) read from `state.doc.attrs.defaultTabStopTwips`
   * on the body doc — HF content doesn't carry its own doc-level setting,
   * so pass it through so list markers inside headers/footers honor the
   * same tab grid as the body.
   */
  defaultTabStopTwips?: number | null;
};

/**
 * Convert HeaderFooter (document type) to HeaderFooterContent (render type).
 *
 * Routes through the same pipeline as the body: HF.content →
 * headerFooterToProseDoc → toLayoutBlocks → measureBlocks. The inline editor
 * uses the same conversion chain, so block support (paragraph, table, image,
 * textBox, fields) and the inline editor's content stay in lockstep.
 */
export function convertHeaderFooterToContent(
  headerFooter: HeaderFooter | null | undefined,
  contentWidth: number,
  metrics: HeaderFooterMetrics,
  options: ConvertHeaderFooterOptions,
  hfStoryId?: string
): HeaderFooterContent | undefined {
  if (!headerFooter || !headerFooter.content || headerFooter.content.length === 0) {
    return undefined;
  }

  const blocks = hfStoryId ? options.yrsStoryBlocks?.(hfStoryId) : null;
  if (!blocks || blocks.length === 0) return undefined;

  const blocksForMeasure = normalizeHeaderFooterMeasureBlocks(blocks);
  const measures = options.measureBlocks(blocksForMeasure, contentWidth);
  let totalHeight = 0;
  let flowHeight = 0;
  for (let i = 0; i < blocksForMeasure.length; i++) {
    const h = measureFlowHeight(measures[i]);
    totalHeight += h;
    if (contributesToHeaderFooterFlowHeight(blocksForMeasure[i])) flowHeight += h;
  }
  // Use `blocksForMeasure` (the normalized list the `measures` were computed
  // from), NOT the raw `blocks` — otherwise block[i] and measure[i] can desync
  // and per-block flags like `displayMode` are read off the wrong block.
  const { visualTop, visualBottom } = calculateHeaderFooterVisualBounds(
    blocksForMeasure,
    measures,
    totalHeight,
    metrics
  );

  return {
    blocks: blocksForMeasure,
    measures,
    height: totalHeight,
    flowHeight,
    visualTop,
    visualBottom,
  };
}

// ============================================================================
// HF caret rect — used by both React and Vue adapters
// ============================================================================

/** Query-backed HF caret geometry, optionally narrowed to the edited page. */
export function computeHfCaretRectsFromDisplayList(
  queries: DisplayListQueries,
  section: 'header' | 'footer',
  rId: string,
  pos: number,
  pageIndex?: number
): DisplayListRect[] {
  const rects = queries.hfCaretRects(section, rId, pos);
  return pageIndex === undefined ? rects : rects.filter((rect) => rect.pageIndex === pageIndex);
}

/** Query-backed HF selection geometry, optionally narrowed to the edited page. */
export function computeHfSelectionRectsFromDisplayList(
  queries: DisplayListQueries,
  section: 'header' | 'footer',
  rId: string,
  from: number,
  to: number,
  pageIndex?: number
): DisplayListRect[] {
  const rects = queries.hfRangeRects(section, rId, Math.min(from, to), Math.max(from, to));
  return pageIndex === undefined ? rects : rects.filter((rect) => rect.pageIndex === pageIndex);
}
