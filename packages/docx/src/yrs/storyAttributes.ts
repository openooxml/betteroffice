/** Shared Document-model ↔ yrs story attribute helpers. */

export type YrsAttrs = Record<string, unknown>;

/** Recursively remove absent object entries while preserving array positions. */
export function dropNulls(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(dropNulls);
  if (value !== null && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
      if (entry === null || entry === undefined) continue;
      out[key] = dropNulls(entry);
    }
    return out;
  }
  return value;
}

const PARA_ATTR_RENAME: Readonly<Record<string, string>> = {
  styleId: 'pStyle',
  _sectionProperties: 'sectPr',
};

const PARA_ATTR_SKIP: ReadonlySet<string> = new Set([
  'paraId',
  'textId',
  'renderedPageBreakBefore',
  'numPrFromStyle',
]);

export function paraAttrsToPpr(attrs: Record<string, unknown>): YrsAttrs {
  const ppr: YrsAttrs = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (PARA_ATTR_SKIP.has(key) || value === null || value === undefined) continue;
    ppr[PARA_ATTR_RENAME[key] ?? key] = dropNulls(value);
  }
  return ppr;
}

const TABLE_ATTR_SKIP: ReadonlySet<string> = new Set(['columnWidths']);

function structuralAttrs(
  attrs: Record<string, unknown>,
  skipped: ReadonlySet<string> = new Set()
): YrsAttrs {
  const lowered: YrsAttrs = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (skipped.has(key) || value === null || value === undefined) continue;
    lowered[key] = dropNulls(value);
  }
  return lowered;
}

export function tableAttrsToTblPr(attrs: Record<string, unknown>): YrsAttrs {
  return structuralAttrs(attrs, TABLE_ATTR_SKIP);
}

export function tableAttrsToGrid(attrs: Record<string, unknown>): unknown[] {
  const widths = attrs.columnWidths;
  return Array.isArray(widths) ? widths.map(dropNulls) : [];
}

export function tableRowAttrsToTrPr(attrs: Record<string, unknown>): YrsAttrs {
  return structuralAttrs(attrs);
}

export function tableCellAttrsToTcPr(
  attrs: Record<string, unknown>,
  isHeader: boolean
): YrsAttrs {
  const tcPr = structuralAttrs(attrs);
  if (isHeader) tcPr.header = true;
  return tcPr;
}

export function tableCellStoryId(
  parentStoryId: string,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number
): string {
  return `${parentStoryId}:t${tableIndex}:r${rowIndex}c${cellIndex}`;
}

export function blockSdtStoryId(parentStoryId: string, sdtIndex: number): string {
  return `${parentStoryId}:sdt${sdtIndex}`;
}

export function blockSdtAttrsToPayload(attrs: Record<string, unknown>): YrsAttrs {
  return structuralAttrs(attrs);
}

export function headerFooterStoryId(rId: string): string {
  return `hf:${rId}`;
}

export function footnoteStoryId(id: string | number): string {
  return `fn:${id}`;
}

export function endnoteStoryId(id: string | number): string {
  return `en:${id}`;
}
