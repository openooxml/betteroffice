/**
 * ProseMirror to LayoutBlock Converter
 *
 * Converts a ProseMirror document into LayoutBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 *
 * The deep import `@openooxml/.../layout/flow` is part of the
 * public surface (Vue adapter + tests), so the per-domain helpers under
 * ./toLayoutBlocks/ are re-exported from here to keep that path stable.
 * @packageDocumentation
 * @public
 */

import type { EditorTreeNode as PMNode } from '../../types/editorTree';
import type {
  LayoutBlock,
  ParagraphBlock,
  TableBlock,
  TableRow,
  TableCell,
  ImageBlock,
  TextBoxBlock,
  PageBreakBlock,
  SectionBreakBlock,
  SdtGroup,
  ColumnLayout,
  ParagraphAttrs,
} from '../pagination/types';
import { DEFAULT_TEXTBOX_MARGINS, DEFAULT_TEXTBOX_WIDTH } from '../pagination/types';
import type { LegacyParagraphAttrs } from './toLayoutBlocks/legacyTypes';
import type { Theme, SectionProperties } from '../../types/document';
import { resolveColorToHex } from '../../utils/colorResolver';

import { twipsToPixels, constrainImageToPage, nextBlockId } from './toLayoutBlocks/shared';
import { AUTO_PARAGRAPH_SPACING_PX } from '../../utils/units';
import type { ToLayoutBlocksOptions } from './toLayoutBlocks/shared';
import {
  chartNodeToChartBlock,
  paragraphToRuns,
  shapeNodeToShapeBlock,
} from './toLayoutBlocks/runs';
import { convertBorderSpecToLayout, extractCellBorders } from './toLayoutBlocks/borders';
import { computeListMarker } from './toLayoutBlocks/listMarkers';

export type { ToLayoutBlocksOptions } from './toLayoutBlocks/shared';
export { resetBlockIdCounter } from './toLayoutBlocks/shared';
export { convertBorderSpecToLayout } from './toLayoutBlocks/borders';
export { resolveListTemplate } from './toLayoutBlocks/listMarkers';

const DEFAULT_FONT = 'Calibri';
const DEFAULT_SIZE = 11; // points (Word 2007+ default)

/**
 * Convert PM paragraph attrs to layout engine paragraph attrs.
 */
function convertParagraphAttrs(
  pmAttrs: LegacyParagraphAttrs,
  theme?: Theme | null,
  listCounters?: Map<number, number[]>,
  listSeenNumIds?: Set<string>,
  defaultTabStopTwips?: number
): ParagraphAttrs {
  const attrs: ParagraphAttrs = {};

  // Alignment - map DOCX values to CSS-compatible values
  // DOCX uses 'both' for justify, 'distribute' for distributed justify
  if (pmAttrs.alignment) {
    const align = pmAttrs.alignment;
    if (align === 'both' || align === 'distribute') {
      attrs.alignment = 'justify';
    } else if (align === 'left') {
      attrs.alignment = 'left';
    } else if (align === 'center') {
      attrs.alignment = 'center';
    } else if (align === 'right') {
      attrs.alignment = 'right';
    }
    // Other DOCX alignments (mediumKashida, highKashida, lowKashida, thaiDistribute, justify)
    // default to no alignment set (inherits from style or defaults to left)
  }

  // Spacing. HTML-origin auto spacing (w:beforeAutospacing/afterAutospacing)
  // rides on _originalFormatting and overrides any explicit before/after with
  // Word's ~14px auto value — measure it here so pagination matches the rendered
  // editor (issue #811).
  const autoBefore = pmAttrs._originalFormatting?.beforeAutospacing;
  const autoAfter = pmAttrs._originalFormatting?.afterAutospacing;
  if (
    autoBefore ||
    autoAfter ||
    pmAttrs.spaceBefore != null ||
    pmAttrs.spaceAfter != null ||
    pmAttrs.lineSpacing != null
  ) {
    attrs.spacing = {};
    if (autoBefore) {
      attrs.spacing.before = AUTO_PARAGRAPH_SPACING_PX;
    } else if (pmAttrs.spaceBefore != null) {
      attrs.spacing.before = twipsToPixels(pmAttrs.spaceBefore);
    }
    if (autoAfter) {
      attrs.spacing.after = AUTO_PARAGRAPH_SPACING_PX;
    } else if (pmAttrs.spaceAfter != null) {
      attrs.spacing.after = twipsToPixels(pmAttrs.spaceAfter);
    }
    if (pmAttrs.lineSpacing != null) {
      // Line spacing in twips - convert to multiplier or exact
      if (pmAttrs.lineSpacingRule === 'exact' || pmAttrs.lineSpacingRule === 'atLeast') {
        attrs.spacing.line = twipsToPixels(pmAttrs.lineSpacing);
        attrs.spacing.lineUnit = 'px';
        attrs.spacing.lineRule = pmAttrs.lineSpacingRule;
      } else {
        // Auto - line spacing is in 240ths of a line
        attrs.spacing.line = pmAttrs.lineSpacing / 240;
        attrs.spacing.lineUnit = 'multiplier';
        attrs.spacing.lineRule = 'auto';
      }
    }
  }
  if (pmAttrs.spacingExplicit) {
    attrs.spacingExplicit = pmAttrs.spacingExplicit;
  }

  // Indentation - handle list item fallback calculation
  // For list items without explicit indentation, calculate based on level
  let indentLeft = pmAttrs.indentLeft;
  let indentFirstLine = pmAttrs.indentFirstLine;
  let hangingIndent = pmAttrs.hangingIndent;
  if (pmAttrs.numPr?.numId && indentLeft == null) {
    // Fallback: calculate indentation based on level
    // Each level indents 0.5 inch (720 twips) more
    const level = pmAttrs.numPr.ilvl ?? 0;
    // Base indentation: 0.5 inch (720 twips) per level
    // Level 0 = 720 twips, Level 1 = 1440 twips, etc.
    indentLeft = (level + 1) * 720;
    // Default hanging indent of 360 twips for the list marker
    if (indentFirstLine == null) {
      indentFirstLine = -360;
      hangingIndent = true;
    }
  }

  if (indentLeft != null || pmAttrs.indentRight != null || indentFirstLine != null) {
    attrs.indent = {};
    if (indentLeft != null) {
      attrs.indent.left = twipsToPixels(indentLeft);
    }
    if (pmAttrs.indentRight != null) {
      attrs.indent.right = twipsToPixels(pmAttrs.indentRight);
    }
    if (indentFirstLine != null) {
      if (hangingIndent) {
        // Hanging indent: indentFirstLine is stored as negative, convert to positive for rendering
        attrs.indent.hanging = Math.abs(twipsToPixels(indentFirstLine));
      } else {
        attrs.indent.firstLine = twipsToPixels(indentFirstLine);
      }
    }
  }

  // Style ID
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // Borders
  if (pmAttrs.borders) {
    const borders = pmAttrs.borders;
    attrs.borders = {};

    const convertBorder = (border: typeof borders.top) =>
      border ? convertBorderSpecToLayout(border, theme) : undefined;

    if (borders.top) attrs.borders.top = convertBorder(borders.top);
    if (borders.bottom) attrs.borders.bottom = convertBorder(borders.bottom);
    if (borders.left) attrs.borders.left = convertBorder(borders.left);
    if (borders.right) attrs.borders.right = convertBorder(borders.right);
    if (borders.between) attrs.borders.between = convertBorder(borders.between);
    if (borders.bar) attrs.borders.bar = convertBorder(borders.bar);

    // Only include if at least one border is set
    if (
      !attrs.borders.top &&
      !attrs.borders.bottom &&
      !attrs.borders.left &&
      !attrs.borders.right &&
      !attrs.borders.between &&
      !attrs.borders.bar
    ) {
      delete attrs.borders;
    }
  }

  const shadingHex = resolveColorToHex(pmAttrs.shading?.fill, theme);
  if (shadingHex) attrs.shading = `#${shadingHex}`;

  // Tab stops
  if (pmAttrs.tabs && pmAttrs.tabs.length > 0) {
    attrs.tabs = pmAttrs.tabs.map((tab) => ({
      val: mapTabAlignment(tab.alignment),
      pos: tab.position,
      leader: tab.leader as
        | 'none'
        | 'dot'
        | 'hyphen'
        | 'underscore'
        | 'heavy'
        | 'middleDot'
        | undefined,
    }));
  }

  // Page break control. `renderedPageBreakBefore` (Word's
  // `<w:lastRenderedPageBreak/>` marker) is informational — it records where
  // Word last broke the page. ECMA-376 §17.4.16 does NOT specify it as a
  // forced break, and Word does not honor it as one on reflow. Preserve the
  // attr through round-trip so the marker is re-emitted on save, but do not
  // act on it during layout.
  if (pmAttrs.pageBreakBefore) {
    attrs.pageBreakBefore = true;
  }
  if (pmAttrs.keepNext) {
    attrs.keepNext = true;
  }
  if (pmAttrs.keepLines) {
    attrs.keepLines = true;
  }
  if (pmAttrs.contextualSpacing) {
    attrs.contextualSpacing = true;
  }
  if (pmAttrs.bidi) {
    attrs.bidi = true;
  }
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // List properties
  if (pmAttrs.numPr) {
    attrs.numPr = {
      numId: pmAttrs.numPr.numId,
      ilvl: pmAttrs.numPr.ilvl,
    };
  }
  // Resolve the OOXML lvlText template (e.g. "%1.") into the rendered marker
  // ("1.", "II.", "1.1.", etc.). Single source of truth — covers body, table,
  // and text-box paragraphs since they all share this attr conversion.
  const resolvedMarker =
    listCounters && listSeenNumIds
      ? computeListMarker(pmAttrs, listCounters, listSeenNumIds)
      : null;
  if (resolvedMarker != null) {
    attrs.listMarker = resolvedMarker;
  } else if (pmAttrs.listMarker) {
    attrs.listMarker = pmAttrs.listMarker;
  }
  if (pmAttrs.listIsBullet != null) {
    attrs.listIsBullet = pmAttrs.listIsBullet;
  }
  if (pmAttrs.listMarkerHidden) {
    attrs.listMarkerHidden = true;
  }
  if (pmAttrs.listMarkerFontFamily) {
    attrs.listMarkerFontFamily = pmAttrs.listMarkerFontFamily;
  }
  if (pmAttrs.listMarkerFontSize) {
    attrs.listMarkerFontSize = pmAttrs.listMarkerFontSize;
  }
  if (pmAttrs.listMarkerSuffix) {
    attrs.listMarkerSuffix = pmAttrs.listMarkerSuffix;
  }
  if (defaultTabStopTwips !== undefined) {
    attrs.defaultTabStopTwips = defaultTabStopTwips;
  }

  // Default font for empty paragraph measurement (from style's rPr / pPr/rPr)
  const dtf = pmAttrs.defaultTextFormatting as
    | { fontSize?: number; fontFamily?: { ascii?: string; hAnsi?: string } }
    | undefined;
  if (dtf) {
    if (dtf.fontSize != null) {
      // fontSize in TextFormatting is in half-points, convert to points
      attrs.defaultFontSize = dtf.fontSize / 2;
    }
    if (dtf.fontFamily) {
      attrs.defaultFontFamily = (dtf.fontFamily.ascii || dtf.fontFamily.hAnsi) as
        | string
        | undefined;
    }
  }

  // Paragraph-mark tracked-change attrs — surface to the painter and to the
  // measurement cache key so revised paragraphs paint their pilcrow / change
  // bar and don't cross-pollute the cache with same-text-no-revision peers.
  if (pmAttrs.pPrIns) attrs.pPrIns = pmAttrs.pPrIns;
  if (pmAttrs.pPrDel) attrs.pPrDel = pmAttrs.pPrDel;

  // List-marker tracked-change state. When the numbering itself is a pending
  // revision the painter colors the marker (Word shows an inserted list item's
  // number in the change color). The numbering counts as inserted when the
  // paragraph mark was inserted (pPrIns) or a pPrChange added it (its prior
  // snapshot had no numId); deleted when the paragraph mark is deleted.
  if (pmAttrs.numPr) {
    const pPrChange = pmAttrs.pPrChange as Array<{
      previousFormatting?: { numPr?: unknown };
    }> | null;
    const numberingAdded =
      Array.isArray(pPrChange) && pPrChange.some((c) => !c?.previousFormatting?.numPr);
    if (pmAttrs.pPrDel) attrs.listMarkerRevision = 'del';
    else if (pmAttrs.pPrIns || numberingAdded) attrs.listMarkerRevision = 'ins';
  }

  return attrs;
}

/**
 * Map document TabStopAlignment to layout engine TabAlignment
 */
function mapTabAlignment(
  align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num'
): 'start' | 'end' | 'center' | 'decimal' | 'bar' | 'clear' {
  switch (align) {
    case 'left':
      return 'start';
    case 'right':
      return 'end';
    case 'center':
      return 'center';
    case 'decimal':
      return 'decimal';
    case 'bar':
      return 'bar';
    case 'clear':
      return 'clear';
    case 'num':
      return 'start'; // Number tab treated as left-aligned
    default:
      return 'start';
  }
}

/**
 * Convert a paragraph node to a ParagraphBlock.
 */
function convertParagraph(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions
): ParagraphBlock {
  const pmAttrs = node.attrs as LegacyParagraphAttrs;
  const runs = paragraphToRuns(node, startPos, options);
  const attrs = convertParagraphAttrs(
    pmAttrs,
    options.theme,
    options.listCounters,
    options.listSeenNumIds,
    options.defaultTabStopTwips
  );

  return {
    kind: 'paragraph',
    id: nextBlockId(),
    paraId: pmAttrs.paraId || undefined,
    runs,
    attrs,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

function convertParagraphWithShapes(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions
): LayoutBlock[] {
  let hasDrawingBlock = false;
  node.forEach((child) => {
    if (child.type.name === 'shape' || child.type.name === 'chart') hasDrawingBlock = true;
  });
  if (!hasDrawingBlock) return [convertParagraph(node, startPos, options)];

  const blocks: LayoutBlock[] = [];
  let pendingChildren: PMNode[] = [];
  let pendingStartOffset = 0;

  const flushPendingParagraph = (): void => {
    if (pendingChildren.length === 0) return;
    const paragraph = node.type.create(node.attrs, pendingChildren, node.marks);
    const block = convertParagraph(paragraph, startPos + pendingStartOffset, options);
    if (block.runs.length > 0) {
      blocks.push(block);
    }
    pendingChildren = [];
  };

  node.forEach((child, childOffset) => {
    if (child.type.name === 'shape' || child.type.name === 'chart') {
      flushPendingParagraph();
      if (child.type.name === 'shape') {
        const shapeBlock = shapeNodeToShapeBlock(child, startPos + 1 + childOffset, options);
        if (shapeBlock) blocks.push(shapeBlock);
      } else {
        const chartBlock = chartNodeToChartBlock(child, startPos + 1 + childOffset, options);
        if (chartBlock) blocks.push(chartBlock);
      }
      return;
    }

    if (pendingChildren.length === 0) {
      pendingStartOffset = childOffset;
    }
    pendingChildren.push(child);
  });

  flushPendingParagraph();
  return blocks;
}

/**
 * Convert a table cell node.
 */
function convertTableCell(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions,
  tableCellMargins?: { top?: number; bottom?: number; left?: number; right?: number }
): TableCell {
  const blocks: LayoutBlock[] = [];
  let offset = startPos + 1; // +1 for opening tag

  node.forEach((child) => {
    if (child.type.name === 'paragraph') {
      blocks.push(...convertParagraphWithShapes(child, offset, options));
    } else if (child.type.name === 'table') {
      blocks.push(convertTable(child, offset, options));
    }
    offset += child.nodeSize;
  });

  const attrs = node.attrs;
  const widthValue = attrs.width as number | undefined;
  const widthType = attrs.widthType as string | undefined;
  const width =
    widthValue && (!widthType || widthType === 'dxa' || widthType === 'auto')
      ? twipsToPixels(widthValue)
      : undefined;

  // Resolve cell padding via the OOXML cascade (§17.4.41 + §17.4.79):
  //   1. cell w:tcMar (per-side, only when value > 0 — Word treats an
  //      explicit zero as "fall through, not literal zero")
  //   2. table-level w:tblCellMar / resolved table style's tblPr.cellMargins
  //
  // Tier 2 is fully resolved upstream by toProseDoc.convertTable, which
  // walks the inline tblCellMar → table-style → basedOn chain → default
  // table style cascade. We just consume the flattened result here. There
  // is no hardcoded "TableNormal default" fallback any more — any document
  // with a styles.xml will have its default table style's cellMargins
  // already in tableCellMargins; a document genuinely missing every tier
  // renders with 0 padding (the spec literal), which is correct for that
  // edge case.
  const margins = attrs.margins as
    | { top?: number; bottom?: number; left?: number; right?: number }
    | undefined;
  const resolveSide = (cellTwips: number | undefined, tableTwips: number | undefined): number => {
    if (cellTwips != null) {
      const px = twipsToPixels(cellTwips);
      if (px > 0) return px;
    }
    if (tableTwips != null) {
      const px = twipsToPixels(tableTwips);
      if (px >= 0) return px;
    }
    return 0;
  };
  const padding = {
    top: resolveSide(margins?.top, tableCellMargins?.top),
    right: resolveSide(margins?.right, tableCellMargins?.right),
    bottom: resolveSide(margins?.bottom, tableCellMargins?.bottom),
    left: resolveSide(margins?.left, tableCellMargins?.left),
  };

  // Surface tracked-cell marker (cellIns / cellDel / cellMerge) so the
  // painter can color the cell border. Same model as ParagraphAttrs.pPrIns.
  const cellMarker = attrs.cellMarker as
    | import('../../types/content/trackedChange').CellMarker
    | null
    | undefined;

  return {
    id: nextBlockId(),
    blocks,
    colSpan: attrs.colspan as number,
    rowSpan: attrs.rowspan as number,
    width,
    widthValue,
    widthType,
    verticalAlign: attrs.verticalAlign as 'top' | 'center' | 'bottom' | undefined,
    background: attrs.backgroundColor ? `#${attrs.backgroundColor}` : undefined,
    borders: extractCellBorders(attrs as Record<string, unknown>, options.theme),
    padding,
    noWrap: (attrs.noWrap as boolean | undefined) || undefined,
    trackedMarker: cellMarker ?? undefined,
  };
}

/**
 * Convert a table row node.
 */
function convertTableRow(
  node: PMNode,
  startPos: number,
  options: ToLayoutBlocksOptions,
  tableCellMargins?: { top?: number; bottom?: number; left?: number; right?: number }
): TableRow {
  const cells: TableCell[] = [];
  let offset = startPos + 1; // +1 for opening tag

  node.forEach((child) => {
    if (child.type.name === 'tableCell' || child.type.name === 'tableHeader') {
      cells.push(convertTableCell(child, offset, options, tableCellMargins));
    }
    offset += child.nodeSize;
  });

  const attrs = node.attrs;
  // `w:cantSplit` (§17.4.6) is preserved only in _originalFormatting; surface it
  // so the layout engine keeps the row whole across page boundaries.
  const rowFormatting = attrs._originalFormatting as { cantSplit?: boolean } | null | undefined;
  return {
    id: nextBlockId(),
    cells,
    height: attrs.height ? twipsToPixels(attrs.height as number) : undefined,
    heightRule: (attrs.heightRule as 'auto' | 'atLeast' | 'exact') ?? undefined,
    isHeader: attrs.isHeader as boolean | undefined,
    cantSplit: rowFormatting?.cantSplit || undefined,
    trackedIns:
      (attrs.trIns as import('../../types/content/trackedChange').RevisionInfo | null) ?? undefined,
    trackedDel:
      (attrs.trDel as import('../../types/content/trackedChange').RevisionInfo | null) ?? undefined,
  };
}

/**
 * Convert a table node to a TableBlock.
 */
function convertTable(node: PMNode, startPos: number, options: ToLayoutBlocksOptions): TableBlock {
  const rows: TableRow[] = [];
  let offset = startPos + 1; // +1 for opening tag

  // Read the table-level <w:tblCellMar> default cell margins (twips). Cells
  // cascade to this when their own w:tcMar is absent or explicit-zero. PM
  // stores it as `cellMargins: { top, bottom, left, right }` in twips.
  const tableCellMargins = node.attrs.cellMargins as
    | { top?: number; bottom?: number; left?: number; right?: number }
    | undefined;

  node.forEach((child) => {
    if (child.type.name === 'tableRow') {
      rows.push(convertTableRow(child, offset, options, tableCellMargins));
    }
    offset += child.nodeSize;
  });

  // Extract columnWidths from node attributes and convert from twips to pixels
  const columnWidthsTwips = node.attrs.columnWidths as number[] | undefined;
  let columnWidths = columnWidthsTwips?.map(twipsToPixels);

  const width = node.attrs.width as number | undefined;
  const widthType = node.attrs.widthType as string | undefined;

  // Fallback: compute column widths from first row cell widths if table attr is missing
  if (!columnWidths && rows.length > 0) {
    const firstRow = rows[0];
    const cellWidths = firstRow.cells.map((cell) => cell.width);
    // Only use if all cells have widths defined
    if (cellWidths.every((w) => w !== undefined && w > 0)) {
      columnWidths = cellWidths as number[];
    }
  }

  // Extract justification
  const justification = node.attrs.justification as 'left' | 'center' | 'right' | undefined;

  // Extract table indent from _originalFormatting (w:tblInd)
  const originalFormatting = node.attrs._originalFormatting as
    | { indent?: { value: number; type: string } }
    | undefined;
  const indentPx =
    originalFormatting?.indent?.value && originalFormatting.indent.type === 'dxa'
      ? twipsToPixels(originalFormatting.indent.value)
      : undefined;

  const floating = node.attrs.floating as
    | {
        horzAnchor?: 'margin' | 'page' | 'text';
        vertAnchor?: 'margin' | 'page' | 'text';
        tblpX?: number;
        tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
        tblpY?: number;
        tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
        topFromText?: number;
        bottomFromText?: number;
        leftFromText?: number;
        rightFromText?: number;
      }
    | undefined;

  const floatingPx = floating
    ? {
        horzAnchor: floating.horzAnchor,
        vertAnchor: floating.vertAnchor,
        tblpX: floating.tblpX !== undefined ? twipsToPixels(floating.tblpX) : undefined,
        tblpXSpec: floating.tblpXSpec,
        tblpY: floating.tblpY !== undefined ? twipsToPixels(floating.tblpY) : undefined,
        tblpYSpec: floating.tblpYSpec,
        topFromText:
          floating.topFromText !== undefined ? twipsToPixels(floating.topFromText) : undefined,
        bottomFromText:
          floating.bottomFromText !== undefined
            ? twipsToPixels(floating.bottomFromText)
            : undefined,
        leftFromText:
          floating.leftFromText !== undefined ? twipsToPixels(floating.leftFromText) : undefined,
        rightFromText:
          floating.rightFromText !== undefined ? twipsToPixels(floating.rightFromText) : undefined,
      }
    : undefined;

  return {
    kind: 'table',
    id: nextBlockId(),
    rows,
    columnWidths,
    width,
    widthType,
    justification,
    bidi: node.attrs.bidi || undefined,
    indent: indentPx,
    floating: floatingPx,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Convert an image node to an ImageBlock.
 */
function convertImage(node: PMNode, startPos: number, pageContentHeight?: number): ImageBlock {
  const attrs = node.attrs;
  const wrapType = attrs.wrapType as string | undefined;

  // Only anchor images with 'behind' or 'inFront' wrap types
  // Other wrap types (square, tight, through, topAndBottom) need text wrapping
  // which we don't support yet, so treat them as block-level images
  const shouldAnchor = wrapType === 'behind' || wrapType === 'inFront';

  const constrained = constrainImageToPage(
    (attrs.width as number) || 100,
    (attrs.height as number) || 100,
    pageContentHeight
  );
  const transform = attrs.transform as string | undefined;
  const rotate = transform?.match(/rotate\(\s*(-?\d+(?:\.\d+)?)deg\s*\)/i);
  const parsedRotation = rotate ? Number(rotate[1]) : 0;
  const rotationDeg = Number.isFinite(parsedRotation) ? ((parsedRotation % 360) + 360) % 360 : 0;
  const radians = (rotationDeg * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const boundsWidth = constrained.width * cos + constrained.height * sin;
  const boundsHeight = constrained.width * sin + constrained.height * cos;

  return {
    kind: 'image',
    id: nextBlockId(),
    src: attrs.src as string,
    width: constrained.width,
    height: constrained.height,
    alt: attrs.alt as string | undefined,
    transform,
    rotationDeg: rotationDeg || undefined,
    flipH: transform?.includes('scaleX(-1)') || undefined,
    flipV: transform?.includes('scaleY(-1)') || undefined,
    rotationBounds:
      rotationDeg !== 0
        ? {
            width: boundsWidth,
            height: boundsHeight,
            offsetX: (boundsWidth - constrained.width) / 2,
            offsetY: (boundsHeight - constrained.height) / 2,
          }
        : undefined,
    anchor: shouldAnchor
      ? {
          isAnchored: true,
          offsetH: attrs.distLeft as number | undefined,
          offsetV: attrs.distTop as number | undefined,
          behindDoc: wrapType === 'behind',
        }
      : undefined,
    hlinkHref: attrs.hlinkHref as string | undefined,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Convert a textBox PM node to a TextBoxBlock.
 */
function convertTextBoxNode(
  node: PMNode,
  startPos: number,
  opts: ToLayoutBlocksOptions
): TextBoxBlock {
  const attrs = node.attrs;
  const contentBlocks: ParagraphBlock[] = [];

  // Convert child paragraphs inside the text box
  node.forEach((child, offset) => {
    if (child.type.name === 'paragraph') {
      const block = convertParagraph(child, startPos + 1 + offset, opts);
      contentBlocks.push(block);
    }
  });

  return {
    kind: 'textBox',
    id: nextBlockId(),
    width: (attrs.width as number) ?? DEFAULT_TEXTBOX_WIDTH,
    height: (attrs.height as number) ?? undefined,
    fillColor: attrs.fillColor as string | undefined,
    outlineWidth: attrs.outlineWidth as number | undefined,
    outlineColor: attrs.outlineColor as string | undefined,
    outlineStyle: attrs.outlineStyle as string | undefined,
    margins: {
      top: (attrs.marginTop as number) ?? DEFAULT_TEXTBOX_MARGINS.top,
      bottom: (attrs.marginBottom as number) ?? DEFAULT_TEXTBOX_MARGINS.bottom,
      left: (attrs.marginLeft as number) ?? DEFAULT_TEXTBOX_MARGINS.left,
      right: (attrs.marginRight as number) ?? DEFAULT_TEXTBOX_MARGINS.right,
    },
    content: contentBlocks,
    displayMode: attrs.displayMode as TextBoxBlock['displayMode'],
    cssFloat: attrs.cssFloat as TextBoxBlock['cssFloat'],
    wrapType: attrs.wrapType as string | undefined,
    wrapText: attrs.wrapText as TextBoxBlock['wrapText'],
    anchorTarget: attrs.anchorTarget as TextBoxBlock['anchorTarget'],
    position: attrs.position as TextBoxBlock['position'],
    distTop: attrs.distTop as number | undefined,
    distBottom: attrs.distBottom as number | undefined,
    distLeft: attrs.distLeft as number | undefined,
    distRight: attrs.distRight as number | undefined,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Convert a ProseMirror document to LayoutBlock array.
 *
 * Walks the document tree, converting each node to the appropriate block type.
 * Tracks pmStart/pmEnd positions for each block for click-to-position mapping.
 */
export function toLayoutBlocks(doc: PMNode, options: ToLayoutBlocksOptions = {}): LayoutBlock[] {
  // Doc-level `defaultTabStopTwips` (from settings.xml) rides on the PM
  // doc node so callers don't have to plumb a separate prop. Explicit
  // options still win for callers that override.
  const docDefaultTabStop = doc.attrs?.defaultTabStopTwips as number | undefined;
  const opts: ToLayoutBlocksOptions = {
    ...options,
    defaultFont: options.defaultFont ?? DEFAULT_FONT,
    defaultSize: options.defaultSize ?? DEFAULT_SIZE,
    defaultTabStopTwips: options.defaultTabStopTwips ?? docDefaultTabStop,
  };

  const blocks: LayoutBlock[] = [];
  const offset = 0; // Start at document beginning
  let lastSectionMarginsTwips: { top: number; bottom: number; left: number; right: number } = {
    top: 1440,
    bottom: 1440,
    left: 1440,
    right: 1440,
  };
  // Shared counter map: paragraphs in tables and text boxes update it too,
  // so list numbering stays continuous across containers.
  if (!opts.listCounters) {
    opts.listCounters = new Map<number, number[]>();
  }
  if (!opts.listSeenNumIds) {
    opts.listSeenNumIds = new Set<string>();
  }

  /**
   * Convert one PM node to flow block(s), appending to `blocks`. Recurses
   * into block-level SDTs: their children become normal, independently
   * paginated flow blocks (a control may span pages), each tagged with the
   * enclosing SDT group(s) via `sdtGroups` so the painter can redraw the
   * control boundary.
   */
  const processNode = (node: PMNode, pos: number, sdtGroups: SdtGroup[]): void => {
    if (node.type.name === 'blockSdt') {
      const a = node.attrs as Record<string, unknown>;
      const group: SdtGroup = {
        id: `sdt@${pos}`,
        sdtType: typeof a.sdtType === 'string' ? a.sdtType : 'richText',
        tag: a.tag != null ? String(a.tag) : undefined,
        alias: a.alias != null ? String(a.alias) : undefined,
        lock: a.lock != null ? String(a.lock) : undefined,
        checked: typeof a.checked === 'boolean' ? a.checked : undefined,
        bound: a.dataBinding != null ? true : undefined,
        repeatingItem: /<w15:repeatingSectionItem[\s/>]/.test(String(a.rawPropertiesXml ?? ''))
          ? true
          : undefined,
      };
      const childGroups = [...sdtGroups, group];
      // Child PM position = SDT node start + 1 (enter the node) + child offset.
      node.forEach((child, childOffset) => {
        processNode(child, pos + 1 + childOffset, childGroups);
      });
      return;
    }

    const startLen = blocks.length;

    switch (node.type.name) {
      case 'paragraph':
        {
          const converted = convertParagraphWithShapes(node, pos, opts);
          const pmAttrs = node.attrs as LegacyParagraphAttrs;

          blocks.push(...converted);

          // Emit section break block if this paragraph ends a section
          const secProps = pmAttrs._sectionProperties as SectionProperties | undefined;
          if (secProps || pmAttrs.sectionBreakType) {
            const sectionBreak: SectionBreakBlock = {
              kind: 'sectionBreak',
              id: nextBlockId(),
              type: (secProps?.sectionStart ??
                pmAttrs.sectionBreakType) as SectionBreakBlock['type'],
            };

            if (secProps) {
              // Populate page size when at least one dimension is overridden.
              if (secProps.pageWidth !== undefined || secProps.pageHeight !== undefined) {
                sectionBreak.pageSize = {
                  w: twipsToPixels(secProps.pageWidth ?? 12240),
                  h: twipsToPixels(secProps.pageHeight ?? 15840),
                };
              }
              // Section overrides any margin → emit a full margins record;
              // unset sides inherit from the prior section (tracked above)
              // instead of resetting to the OOXML 1440 default.
              if (
                secProps.marginTop !== undefined ||
                secProps.marginBottom !== undefined ||
                secProps.marginLeft !== undefined ||
                secProps.marginRight !== undefined
              ) {
                const mergedTwips = {
                  top: secProps.marginTop ?? lastSectionMarginsTwips.top,
                  bottom: secProps.marginBottom ?? lastSectionMarginsTwips.bottom,
                  left: secProps.marginLeft ?? lastSectionMarginsTwips.left,
                  right: secProps.marginRight ?? lastSectionMarginsTwips.right,
                };
                sectionBreak.margins = {
                  top: twipsToPixels(mergedTwips.top),
                  bottom: twipsToPixels(mergedTwips.bottom),
                  left: twipsToPixels(mergedTwips.left),
                  right: twipsToPixels(mergedTwips.right),
                };
                lastSectionMarginsTwips = mergedTwips;
              }
              // Populate columns
              const colCount = secProps.columnCount ?? 1;
              if (colCount > 1) {
                const cols: ColumnLayout = {
                  count: colCount,
                  gap: twipsToPixels(secProps.columnSpace ?? 720),
                  equalWidth: secProps.equalWidth ?? true,
                  separator: secProps.separator,
                };
                sectionBreak.columns = cols;
              }
            }

            blocks.push(sectionBreak);
          }
        }
        break;

      case 'table':
        blocks.push(convertTable(node, pos, opts));
        break;

      case 'image':
        // Standalone image block (if not inline)
        blocks.push(convertImage(node, pos, opts.pageContentHeight));
        break;

      case 'textBox':
        blocks.push(convertTextBoxNode(node, pos, opts));
        break;

      case 'horizontalRule':
      case 'pageBreak': {
        const pb: PageBreakBlock = {
          kind: 'pageBreak',
          id: nextBlockId(),
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };
        blocks.push(pb);
        break;
      }
    }

    if (sdtGroups.length > 0) {
      for (let k = startLen; k < blocks.length; k++) {
        blocks[k].sdtGroups = sdtGroups;
      }
    }
  };

  doc.forEach((node, nodeOffset) => {
    processNode(node, offset + nodeOffset, []);
  });

  return blocks;
}
