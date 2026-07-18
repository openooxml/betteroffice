import type {
  ParagraphFormatting,
  TableCellFormatting,
  TableCellPropertyChange,
  TableFormatting,
  TablePropertyChange,
  TableRowFormatting,
  TableRowPropertyChange,
} from '../types/document';
import type { RevisionInfo } from '../types/content/trackedChange';

export interface ParagraphSaveAttrs extends Record<string, unknown> {
  alignment?: ParagraphFormatting['alignment'];
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  lineSpacingRule?: ParagraphFormatting['lineSpacingRule'];
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  hangingIndent?: boolean;
  numPr?: ParagraphFormatting['numPr'];
  numPrFromStyle?: ParagraphFormatting['numPr'];
  styleId?: string;
  borders?: ParagraphFormatting['borders'];
  shading?: ParagraphFormatting['shading'];
  tabs?: ParagraphFormatting['tabs'];
  outlineLevel?: number;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
  bidi?: boolean;
  _originalFormatting?: ParagraphFormatting;
  _originalRunBoundaries?: unknown[];
}

function isStyleSourcedNumPr(attrs: ParagraphSaveAttrs): boolean {
  return (
    attrs.numPrFromStyle != null &&
    attrs.numPr != null &&
    JSON.stringify(attrs.numPr) === JSON.stringify(attrs.numPrFromStyle)
  );
}

export function paragraphAttrsToFormatting(
  attrs: ParagraphSaveAttrs
): ParagraphFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.alignment !== (orig.alignment || undefined)) {
      result.alignment = attrs.alignment || undefined;
    }
    if (isStyleSourcedNumPr(attrs)) {
      delete result.numPr;
      delete result.numPrFromStyle;
    } else if (
      attrs.numPr !== orig.numPr &&
      JSON.stringify(attrs.numPr) !== JSON.stringify(orig.numPr)
    ) {
      result.numPr = attrs.numPr || undefined;
      delete result.numPrFromStyle;
    }
    if (attrs.styleId !== (orig.styleId || undefined)) {
      result.styleId = attrs.styleId || undefined;
    }
    if (attrs.pageBreakBefore !== (orig.pageBreakBefore || undefined)) {
      result.pageBreakBefore = attrs.pageBreakBefore || undefined;
    }
    if (attrs.bidi !== (orig.bidi || undefined)) {
      result.bidi = attrs.bidi || undefined;
    }
    return result;
  }

  const hasFormatting =
    attrs.alignment ||
    attrs.spaceBefore ||
    attrs.spaceAfter ||
    attrs.lineSpacing ||
    attrs.indentLeft ||
    attrs.indentRight ||
    attrs.indentFirstLine ||
    attrs.numPr ||
    attrs.styleId ||
    attrs.borders ||
    attrs.shading ||
    attrs.tabs ||
    attrs.outlineLevel != null ||
    attrs.contextualSpacing ||
    attrs.pageBreakBefore ||
    attrs.bidi;
  if (!hasFormatting) return undefined;

  return {
    alignment: attrs.alignment || undefined,
    spaceBefore: attrs.spaceBefore || undefined,
    spaceAfter: attrs.spaceAfter || undefined,
    lineSpacing: attrs.lineSpacing || undefined,
    lineSpacingRule: attrs.lineSpacingRule || undefined,
    indentLeft: attrs.indentLeft || undefined,
    indentRight: attrs.indentRight || undefined,
    indentFirstLine: attrs.indentFirstLine || undefined,
    hangingIndent: attrs.hangingIndent || undefined,
    numPr: isStyleSourcedNumPr(attrs) ? undefined : attrs.numPr || undefined,
    styleId: attrs.styleId || undefined,
    borders: attrs.borders || undefined,
    shading: attrs.shading || undefined,
    tabs: attrs.tabs || undefined,
    outlineLevel: attrs.outlineLevel ?? undefined,
    contextualSpacing: attrs.contextualSpacing || undefined,
    pageBreakBefore: attrs.pageBreakBefore || undefined,
    bidi: attrs.bidi || undefined,
  };
}

export interface TableSaveAttrs extends Record<string, unknown> {
  styleId?: string;
  width?: number;
  widthType?: string;
  justification?: TableFormatting['justification'];
  columnWidths?: number[];
  tableLayout?: TableFormatting['layout'];
  floating?: TableFormatting['floating'];
  cellMargins?: { top?: number; bottom?: number; left?: number; right?: number };
  look?: TableFormatting['look'];
  bidi?: boolean;
  _originalFormatting?: TableFormatting;
  tblPrChange?: TablePropertyChange[] | null;
}

export function tableAttrsToFormatting(attrs: TableSaveAttrs): TableFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.styleId !== (orig.styleId || undefined)) result.styleId = attrs.styleId || undefined;
    if (attrs.justification !== (orig.justification || undefined)) {
      result.justification = attrs.justification || undefined;
    }
    if (attrs.floating !== (orig.floating || undefined)) {
      result.floating = attrs.floating || undefined;
    }
    if (attrs.tableLayout !== (orig.layout || undefined)) {
      result.layout = attrs.tableLayout || undefined;
    }
    if (attrs.look !== (orig.look || undefined)) result.look = attrs.look || undefined;
    if (attrs.bidi !== (orig.bidi || undefined)) result.bidi = attrs.bidi || undefined;
    if (attrs.width !== orig.width?.value || attrs.widthType !== orig.width?.type) {
      result.width =
        attrs.width != null || attrs.widthType
          ? {
              value: attrs.width ?? 0,
              type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
            }
          : undefined;
    }
    if (attrs.cellMargins) {
      result.cellMargins = measurementMargins(attrs.cellMargins);
    }
    return result;
  }

  const hasFormatting =
    attrs.styleId ||
    attrs.width != null ||
    attrs.widthType ||
    attrs.justification ||
    attrs.tableLayout ||
    attrs.floating ||
    attrs.cellMargins ||
    attrs.look ||
    attrs.bidi;
  if (!hasFormatting) return undefined;
  return {
    styleId: attrs.styleId || undefined,
    width:
      attrs.width != null || attrs.widthType
        ? {
            value: attrs.width ?? 0,
            type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
          }
        : undefined,
    justification: attrs.justification || undefined,
    layout: attrs.tableLayout || undefined,
    floating: attrs.floating || undefined,
    cellMargins: attrs.cellMargins ? measurementMargins(attrs.cellMargins) : undefined,
    look: attrs.look || undefined,
    bidi: attrs.bidi || undefined,
  };
}

function measurementMargins(margins: {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}): NonNullable<TableFormatting['cellMargins']> {
  return {
    top: margins.top != null ? { value: margins.top, type: 'dxa' } : undefined,
    bottom: margins.bottom != null ? { value: margins.bottom, type: 'dxa' } : undefined,
    left: margins.left != null ? { value: margins.left, type: 'dxa' } : undefined,
    right: margins.right != null ? { value: margins.right, type: 'dxa' } : undefined,
  };
}

export interface TableRowSaveAttrs extends Record<string, unknown> {
  height?: number;
  heightRule?: string;
  isHeader?: boolean;
  _originalFormatting?: TableRowFormatting;
  trIns?: RevisionInfo | null;
  trDel?: RevisionInfo | null;
  trPrChange?: TableRowPropertyChange[] | null;
}

export function tableRowAttrsToFormatting(
  attrs: TableRowSaveAttrs
): TableRowFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.height !== (orig.height?.value || undefined)) {
      result.height = attrs.height ? { value: attrs.height, type: 'dxa' } : undefined;
    }
    if (attrs.heightRule !== (orig.heightRule || undefined)) {
      result.heightRule = (attrs.heightRule as 'auto' | 'atLeast' | 'exact') || undefined;
    }
    if (attrs.isHeader !== (orig.header || undefined)) {
      result.header = attrs.isHeader || undefined;
    }
    return result;
  }
  if (!attrs.height && !attrs.isHeader) return undefined;
  return {
    height: attrs.height ? { value: attrs.height, type: 'dxa' } : undefined,
    heightRule: (attrs.heightRule as 'auto' | 'atLeast' | 'exact') || undefined,
    header: attrs.isHeader || undefined,
  };
}

export interface TableCellSaveAttrs extends Record<string, unknown> {
  colspan: number;
  rowspan: number;
  width?: number;
  widthType?: string;
  verticalAlign?: TableCellFormatting['verticalAlign'];
  backgroundColor?: string;
  textDirection?: TableCellFormatting['textDirection'];
  borders?: TableCellFormatting['borders'];
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  _originalFormatting?: TableCellFormatting;
  _originalResolvedFill?: string;
  cellMarker?:
    | { kind: 'ins'; info: RevisionInfo }
    | { kind: 'del'; info: RevisionInfo }
    | {
        kind: 'merge';
        info: RevisionInfo;
        vMerge: 'rest' | 'cont';
        vMergeOrig?: 'rest' | 'cont';
      }
    | null;
  tcPrChange?: TableCellPropertyChange[] | null;
}

export function tableCellAttrsToFormatting(
  attrs: TableCellSaveAttrs
): TableCellFormatting | undefined {
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };
    if (attrs.colspan > 1) result.gridSpan = attrs.colspan;
    if (attrs.width != null) {
      result.width = {
        value: attrs.width,
        type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
      };
    }
    if (attrs.verticalAlign !== (orig.verticalAlign || undefined)) {
      result.verticalAlign = attrs.verticalAlign || undefined;
    }
    if (attrs.backgroundColor) {
      result.shading =
        attrs._originalResolvedFill === attrs.backgroundColor && orig.shading
          ? orig.shading
          : { fill: { rgb: attrs.backgroundColor } };
    } else if (orig.shading) {
      result.shading = undefined;
    }
    if (attrs.borders) result.borders = attrs.borders;
    if (attrs.margins) result.margins = cellMargins(attrs.margins);
    if (attrs.textDirection !== (orig.textDirection || undefined)) {
      result.textDirection = attrs.textDirection || undefined;
    }
    return result;
  }

  const hasFormatting =
    attrs.colspan > 1 ||
    attrs.rowspan > 1 ||
    attrs.width != null ||
    attrs.verticalAlign ||
    attrs.backgroundColor ||
    attrs.borders ||
    attrs.margins ||
    attrs.textDirection;
  if (!hasFormatting) return undefined;
  return {
    gridSpan: attrs.colspan > 1 ? attrs.colspan : undefined,
    width:
      attrs.width != null
        ? {
            value: attrs.width,
            type: (attrs.widthType as 'auto' | 'dxa' | 'pct' | 'nil') || 'dxa',
          }
        : undefined,
    verticalAlign: attrs.verticalAlign || undefined,
    textDirection: attrs.textDirection || undefined,
    shading: attrs.backgroundColor ? { fill: { rgb: attrs.backgroundColor } } : undefined,
    borders: attrs.borders,
    margins: attrs.margins ? cellMargins(attrs.margins) : undefined,
  };
}

function cellMargins(margins: {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}): NonNullable<TableCellFormatting['margins']> {
  const result: NonNullable<TableCellFormatting['margins']> = {};
  if (margins.top != null) result.top = { value: margins.top, type: 'dxa' };
  if (margins.bottom != null) result.bottom = { value: margins.bottom, type: 'dxa' };
  if (margins.left != null) result.left = { value: margins.left, type: 'dxa' };
  if (margins.right != null) result.right = { value: margins.right, type: 'dxa' };
  return result;
}
