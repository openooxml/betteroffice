/**
 * Floating-aware block measurement pipeline.
 *
 * Pre-scans a block list to extract exclusion zones from anchored images,
 * floating tables, and floating text boxes; groups co-located floats so
 * their combined exclusion applies starting from the earliest anchor; then
 * walks the blocks calling the caller-supplied `measureBlock` with the
 * active zones and cumulative Y at each step.
 *
 * Adapters (React, Vue) provide their own `measureBlock` so they can
 * decide e.g. whether to cache paragraph measures. The orchestration,
 * extraction, and grouping live here so both adapters stay in lockstep.
 *
 * @packageDocumentation
 * @public
 */
import {
  isFloatingTextBoxBlock,
  isWrapNone,
  type LayoutBlock,
  type ImageRun,
  type ImageRunPosition,
  type BlockExtent,
  type ParagraphBlock,
  type TableBlock,
  type TextBoxBlock,
  type PageGeometry,
  isTextWrappingFloatingImageRun,
  resolveAnchoredObjectVerticalTop,
} from '../pagination';
import { emuToPixels } from '../../utils/units';
import { clampFloatingWrapMargins, type FloatingImageZone } from './floatingZones';
import { measureTableBlock } from './measureTable';

/**
 * A floating exclusion zone tagged with the block index that anchors it.
 */
interface FloatingZoneWithAnchor extends FloatingImageZone {
  anchorBlockIndex: number;
  /** True for floats positioned relative to page/margin (not paragraph). */
  isMarginRelative?: boolean;
}

/**
 * Maximum block-index distance for paragraph-relative floats to be considered
 * co-located. Anchors within this window with overlapping Y ranges get merged
 * so a body paragraph between them sees the combined exclusion zone. Beyond
 * this window we keep zones independent — different sections of the document
 * routinely have float topY values that coincidentally overlap.
 */
const ANCHOR_PROXIMITY = 4;

/**
 * Block-measurement callback shape passed to {@link measureBlocksWithFloats}.
 * Adapters (React, Vue) supply this so they can decide platform-specific
 * concerns (e.g. paragraph-measure caching, per-section width) while
 * sharing the floating-zone orchestration. This is adapter-author API,
 * not end-consumer API.
 *
 * @public
 */
export type MeasureBlockFn = (
  block: LayoutBlock,
  contentWidth: number,
  floatingZones?: FloatingImageZone[],
  cumulativeY?: number
) => BlockExtent;

/**
 * Page geometry (CSS px) used to resolve page/margin-relative anchored objects
 * into content-area coordinates — currently the vertical anchor of a top-pinned
 * `topAndBottom` band. Same shape the painter uses (see `pageGeometryFromPage`),
 * so both paths resolve to identical positions.
 *
 * @public
 */
export type FloatPageGeometry = PageGeometry;

/**
 * Walk `blocks` and produce one `BlockExtent` per block. Before measuring, this
 * extracts floating exclusion zones (images / floating tables / floating
 * textboxes), groups overlapping co-located floats, and threads the active
 * zones plus cumulative Y into each `measureBlock` call.
 *
 * Pass `pageGeometry` whenever the document may contain page/margin-anchored
 * `topAndBottom` text boxes (e.g. a title banner pinned to the page top):
 * without it their reserved band falls back to flow-relative Y and the band
 * won't line up with where the painter places the box. Build it with the
 * shared `pageGeometryFromPage` helper.
 *
 * @public
 */
export function measureBlocksWithFloats(
  blocks: LayoutBlock[],
  contentWidth: number | number[],
  measureBlock: MeasureBlockFn,
  pageGeometry?: FloatPageGeometry
): BlockExtent[] {
  const defaultWidth = Array.isArray(contentWidth) ? (contentWidth[0] ?? 0) : contentWidth;
  const floatingZonesWithAnchors = extractFloatingZones(
    blocks,
    defaultWidth,
    measureBlock,
    pageGeometry
  );

  const marginRelative = floatingZonesWithAnchors.filter((z) => z.isMarginRelative);
  const paragraphRelative = floatingZonesWithAnchors.filter((z) => !z.isMarginRelative);

  // Margin-relative zones at the same Y likely belong to the same page —
  // group by topY and re-anchor to the earliest block index so subsequent
  // paragraphs see the combined zone.
  const marginByTopY = new Map<number, FloatingZoneWithAnchor[]>();
  for (const z of marginRelative) {
    const group = marginByTopY.get(z.topY) ?? [];
    group.push(z);
    marginByTopY.set(z.topY, group);
  }

  // Paragraph-relative zones merge only when (a) Y ranges overlap AND
  // (b) anchors are within ANCHOR_PROXIMITY blocks. The proximity bound
  // keeps unrelated floats in distant sections from being merged just
  // because their paragraph-local topY values happen to overlap.
  const paragraphGroups = groupOverlappingZones(paragraphRelative, ANCHOR_PROXIMITY);

  const adjustedZones: FloatingZoneWithAnchor[] = [];
  collectReanchoredToEarliest(paragraphGroups, adjustedZones);
  collectReanchoredToEarliest(Array.from(marginByTopY.values()), adjustedZones);

  const zonesByAnchor = new Map<number, FloatingImageZone[]>();
  for (const z of adjustedZones) {
    // A page/margin-pinned full-width band (e.g. a title banner at the top of
    // the page) reserves space from the top of content, so it must reach the
    // blocks that precede its own anchor paragraph. Anchor it at block 0 and
    // keep its content-relative topY/bottomY (cumulativeY is then the running
    // content offset for the blocks it covers).
    //
    // Caveat: this pre-pagination pass has no page/section model, so "top of
    // content" means the start of the whole flow. Exact for the common case
    // (single banner near the document start); a band anchored on a later page
    // or in a later section with different geometry can over-reach. Follow-up.
    const anchor = z.fullWidthBlock && z.isMarginRelative ? 0 : z.anchorBlockIndex;
    const existing = zonesByAnchor.get(anchor) ?? [];
    // Strip the anchor-tracking fields; the rest IS a FloatingImageZone. Spread
    // (rather than copying each field) so new zone fields can't be dropped here.
    const { anchorBlockIndex: _anchorBlockIndex, isMarginRelative: _isMarginRelative, ...zone } = z;
    existing.push(zone);
    zonesByAnchor.set(anchor, existing);
  }

  // Derive from the map keys, not the raw zones — full-width bands are
  // re-anchored to block 0 above, and the activation set must match.
  const anchorIndices = new Set(zonesByAnchor.keys());

  let cumulativeY = 0;
  let activeZones: FloatingImageZone[] = [];

  return blocks.map((block, blockIndex) => {
    if (anchorIndices.has(blockIndex)) {
      cumulativeY = 0;
      activeZones = zonesByAnchor.get(blockIndex) ?? [];
    }

    const zones = activeZones.length > 0 ? activeZones : undefined;
    const blockWidth = Array.isArray(contentWidth)
      ? (contentWidth[blockIndex] ?? defaultWidth)
      : contentWidth;

    const measure = measureBlock(block, blockWidth, zones, cumulativeY);

    if ('totalHeight' in measure) {
      // Floating tables don't advance flow Y (their wrap zone already
      // accounts for vertical space). Other blocks do.
      if (!(block.kind === 'table' && (block as TableBlock).floating)) {
        cumulativeY += measure.totalHeight;
      }
    } else if ('height' in measure) {
      cumulativeY += measure.height;
    }

    return measure;
  });
}

/**
 * Extract floating exclusion zones from all blocks that anchor floats —
 * paragraph runs (images), top-level floating tables, and top-level
 * floating textboxes. The returned zones are in content-area coordinates
 * relative to each anchor block; the orchestration loop in
 * {@link measureBlocksWithFloats} re-anchors and threads them through.
 */
function extractFloatingZones(
  blocks: LayoutBlock[],
  contentWidth: number,
  measureBlock: MeasureBlockFn,
  pageGeometry?: FloatPageGeometry
): FloatingZoneWithAnchor[] {
  const zones: FloatingZoneWithAnchor[] = [];

  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex];
    switch (block.kind) {
      case 'paragraph':
        extractImageZonesFromParagraph(block as ParagraphBlock, blockIndex, contentWidth, zones);
        break;
      case 'table':
        extractFloatingTableZone(
          block as TableBlock,
          blockIndex,
          contentWidth,
          measureBlock,
          zones
        );
        break;
      case 'textBox':
        extractFloatingTextBoxZone(
          block as TextBoxBlock,
          blockIndex,
          contentWidth,
          zones,
          pageGeometry
        );
        break;
    }
  }

  return zones;
}

/**
 * Resolve left/right exclusion margins for an OOXML-positioned anchored
 * object (image or text box). Shared between image-in-paragraph and
 * top-level textbox extraction since both use the same
 * `ImageRunPosition` shape and `cssFloat` fallback.
 */
function computeAnchoredMargins(
  position: ImageRunPosition | undefined,
  cssFloat: 'left' | 'right' | 'none' | undefined,
  width: number,
  distLeft: number,
  distRight: number,
  contentWidth: number
): { leftMargin: number; rightMargin: number } {
  let leftMargin = 0;
  let rightMargin = 0;

  const h = position?.horizontal;
  if (h?.align === 'left') {
    leftMargin = width + distRight;
  } else if (h?.align === 'right') {
    rightMargin = width + distLeft;
  } else if (h?.posOffset !== undefined) {
    const x = emuToPixels(h.posOffset);
    if (x < contentWidth / 2) {
      leftMargin = x + width + distRight;
    } else {
      rightMargin = contentWidth - x + distLeft;
    }
  } else if (cssFloat === 'left') {
    leftMargin = width + distRight;
  } else if (cssFloat === 'right') {
    rightMargin = width + distLeft;
  }

  return clampFloatingWrapMargins(leftMargin, rightMargin, contentWidth);
}

/**
 * True when an OOXML position anchors vertically against the page or
 * margin (not the surrounding paragraph). Margin/page-relative zones
 * apply globally across blocks instead of attaching to one anchor
 * paragraph.
 */
function isPositionMarginRelative(position: ImageRunPosition | undefined): boolean {
  const rel = position?.vertical?.relativeTo;
  return rel === 'margin' || rel === 'page';
}

function extractImageZonesFromParagraph(
  paragraphBlock: ParagraphBlock,
  blockIndex: number,
  contentWidth: number,
  out: FloatingZoneWithAnchor[]
): void {
  for (const run of paragraphBlock.runs) {
    if (run.kind !== 'image') continue;
    const imgRun = run as ImageRun;
    if (!isTextWrappingFloatingImageRun(imgRun)) continue;

    const distTop = imgRun.distTop ?? 0;
    const distBottom = imgRun.distBottom ?? 0;
    const distLeft = imgRun.distLeft ?? 12;
    const distRight = imgRun.distRight ?? 12;

    let topY = 0;
    const v = imgRun.position?.vertical;
    if (v?.align === 'top' && v.relativeTo === 'margin') {
      topY = 0;
    } else if (v?.posOffset !== undefined) {
      topY = emuToPixels(v.posOffset);
    }
    const bottomY = topY + imgRun.height;

    const { leftMargin, rightMargin } = computeAnchoredMargins(
      imgRun.position,
      imgRun.cssFloat,
      imgRun.width,
      distLeft,
      distRight,
      contentWidth
    );

    if (leftMargin > 0 || rightMargin > 0) {
      out.push({
        leftMargin,
        rightMargin,
        topY: topY - distTop,
        bottomY: bottomY + distBottom,
        anchorBlockIndex: blockIndex,
        isMarginRelative: isPositionMarginRelative(imgRun.position),
      });
    }
  }
}

function extractFloatingTableZone(
  tableBlock: TableBlock,
  blockIndex: number,
  contentWidth: number,
  measureBlock: MeasureBlockFn,
  out: FloatingZoneWithAnchor[]
): void {
  const floating = tableBlock.floating;
  if (!floating) return;

  const tableMeasure = measureTableBlock(tableBlock, contentWidth, measureBlock);
  const tableWidth = tableMeasure.totalWidth;
  const tableHeight = tableMeasure.totalHeight;

  const distLeft = floating.leftFromText ?? 12;
  const distRight = floating.rightFromText ?? 12;
  const distTop = floating.topFromText ?? 0;
  const distBottom = floating.bottomFromText ?? 0;

  // Tables use OOXML `w:tblpXSpec` / `tblpX` instead of the image-style
  // `align` / `posOffset`, so the common helper above doesn't apply.
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
  } else if (tableBlock.justification === 'center') {
    x = (contentWidth - tableWidth) / 2;
  } else if (tableBlock.justification === 'right') {
    x = contentWidth - tableWidth;
  }

  let leftMargin = 0;
  let rightMargin = 0;
  if (x < contentWidth / 2) {
    leftMargin = x + tableWidth + distRight;
  } else {
    rightMargin = contentWidth - x + distLeft;
  }

  ({ leftMargin, rightMargin } = clampFloatingWrapMargins(leftMargin, rightMargin, contentWidth));

  const topY = floating.tblpY ?? 0;
  const bottomY = topY + tableHeight;

  out.push({
    leftMargin,
    rightMargin,
    topY: topY - distTop,
    bottomY: bottomY + distBottom,
    anchorBlockIndex: blockIndex,
  });
}

function extractFloatingTextBoxZone(
  tbBlock: TextBoxBlock,
  blockIndex: number,
  contentWidth: number,
  out: FloatingZoneWithAnchor[],
  pageGeometry?: FloatPageGeometry
): void {
  if (!isFloatingTextBoxBlock(tbBlock)) return;
  if (isWrapNone(tbBlock.wrapType)) return;

  const tbWidth = tbBlock.width ?? 0;
  const tbHeight = tbBlock.height ?? 0;
  if (tbWidth <= 0 || tbHeight <= 0) return;

  const distTop = tbBlock.distTop ?? 0;
  const distBottom = tbBlock.distBottom ?? 0;
  const distLeft = tbBlock.distLeft ?? 12;
  const distRight = tbBlock.distRight ?? 12;

  // NOTE: the page-pinned topAndBottom band below is currently text-box only.
  // A topAndBottom anchored *image* is still laid out as a block image on its
  // own line at its anchor (see extractImageZonesFromParagraph / renderPage),
  // so a page-anchored image band is not yet honored — follow-up.
  //
  // topAndBottom: reserve a full-width vertical band so body text flows above
  // and below the box. Page/margin-relative boxes (e.g. a banner pinned to the
  // page top) need their offset translated into content-area coordinates.
  if (tbBlock.wrapType === 'topAndBottom') {
    // Resolve the band's vertical top via the SAME resolver the painter uses,
    // so the reserved band lines up with where the box is painted regardless of
    // anchor kind (page / margin / topMargin / bottomMargin, align or posOffset).
    // fragmentY=0: a topAndBottom band is page/margin-pinned; the paragraph-Y
    // fallback only applies to genuinely paragraph-relative boxes, which this
    // pre-pagination pass anchors at their own block (cumulativeY 0 there).
    const rawTopY = resolveAnchoredObjectVerticalTop(
      { width: tbWidth, height: tbHeight, position: tbBlock.position },
      0,
      pageGeometry
    );
    // Signed top may be negative when the box reaches up into the top margin.
    // The band reserves only the part intruding into content (topY clamped at
    // 0), but its bottom is measured from the true top so the reserved height
    // matches how far the box extends below the content edge.
    const bottomY = rawTopY + tbHeight + distBottom;
    if (bottomY <= 0) return;
    out.push({
      leftMargin: 0,
      rightMargin: 0,
      topY: Math.max(0, rawTopY - distTop),
      bottomY,
      anchorBlockIndex: blockIndex,
      isMarginRelative: isPositionMarginRelative(tbBlock.position),
      fullWidthBlock: true,
    });
    return;
  }

  let topY = 0;
  if (tbBlock.position?.vertical?.posOffset !== undefined) {
    topY = emuToPixels(tbBlock.position.vertical.posOffset);
  }
  const bottomY = topY + tbHeight;

  const { leftMargin, rightMargin } = computeAnchoredMargins(
    tbBlock.position,
    tbBlock.cssFloat,
    tbWidth,
    distLeft,
    distRight,
    contentWidth
  );

  if (leftMargin <= 0 && rightMargin <= 0) return;

  out.push({
    leftMargin,
    rightMargin,
    topY: topY - distTop,
    bottomY: bottomY + distBottom,
    anchorBlockIndex: blockIndex,
    isMarginRelative: isPositionMarginRelative(tbBlock.position),
  });
}

/**
 * Group `zones` such that any two whose Y ranges overlap AND whose
 * anchorBlockIndex differs by no more than `maxAnchorGap` land in the same
 * group. Single-pass; groups merge transitively as zones connect them.
 */
function groupOverlappingZones(
  zones: FloatingZoneWithAnchor[],
  maxAnchorGap: number
): FloatingZoneWithAnchor[][] {
  const groups: FloatingZoneWithAnchor[][] = [];
  for (const z of zones) {
    const target = groups.find((g) =>
      g.some(
        (other) =>
          Math.abs(other.anchorBlockIndex - z.anchorBlockIndex) <= maxAnchorGap &&
          z.topY < other.bottomY &&
          z.bottomY > other.topY
      )
    );
    if (target) target.push(z);
    else groups.push([z]);
  }
  return groups;
}

/**
 * Re-anchor every zone in each group to the group's earliest block index and
 * append the result to `out`.
 */
function collectReanchoredToEarliest(
  groups: FloatingZoneWithAnchor[][],
  out: FloatingZoneWithAnchor[]
): void {
  for (const group of groups) {
    const minAnchor = Math.min(...group.map((z) => z.anchorBlockIndex));
    for (const z of group) {
      out.push({ ...z, anchorBlockIndex: minAnchor });
    }
  }
}
