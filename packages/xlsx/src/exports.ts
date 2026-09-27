/**
 * Read-only structured export of workbook content and its Markdown projection. Every shape
 * here is plain JSON; the walker, limits and Markdown renderer live in Rust.
 *
 * Export never recalculates: formula results are the stored values (`calculation.policy:
 * 'asStored'`, `freshness: 'unverified'`). Anchors name positions by sheet id, index and name
 * and A1 in the exported version or snapshot; none follows later row, column or sheet edits.
 */

import type { OperationRefusal } from '../../../shared/host-contracts/edits';
import type {
  ExportCompletion,
  ExportDiagnostic,
  MarkdownAnchor,
} from '../../../shared/host-contracts/exports';
import type { XlsxCellValue } from './edits';

export type {
  ExportCompletion,
  ExportDiagnostic,
  ExportSeverity,
  MarkdownAnchor,
} from '../../../shared/host-contracts/exports';

/** A zero-based sheet index and at most one rectangular A1 range, defaulting to the used range. */
export interface XlsxExportScope {
  sheet: number;
  range?: string;
}

/** Hidden sheets, rows, columns and names are excluded unless asked for. */
export interface XlsxExportOptions {
  /** Defaults to every sheet. Unions, whole rows or columns and sheet-qualified ranges refuse. */
  scope?: readonly XlsxExportScope[];
  includeHiddenSheets?: boolean;
  includeHiddenRows?: boolean;
  includeHiddenColumns?: boolean;
  /** Defaults to `true`. */
  includeDefinedNames?: boolean;
  includeHiddenNames?: boolean;
  /** Stored cells, 1 to 1,000,000; defaults to 100,000. */
  maxCells?: number;
  /** Compact UTF-8 JSON bytes of the content, at most 16 MiB; defaults to 8 MiB. */
  maxBytes?: number;
}

/** Empty grid positions count toward `maxCells`. */
export interface XlsxMarkdownOptions {
  /** Grid rows per sheet; defaults to 200. */
  maxRows?: number;
  /** Grid columns per sheet; defaults to 50. */
  maxColumns?: number;
  /** Grid positions across all sheets; defaults to 10,000. */
  maxCells?: number;
  /** UTF-8 bytes of Markdown; defaults to 8 MiB. */
  maxBytes?: number;
}

/** A sheet by the id edit batches take and, descriptively, its current position and name. */
export interface XlsxSheetIdentity {
  /** The exporting session's catalog id; `sheet:{index}` for bytes. */
  sheetId: string;
  index: number;
  name: string;
}

/**
 * A `cell` or `range` anchor's `{ sheetId: sheet.sheetId, range: { kind: 'a1', a1 } }` is its
 * `XlsxRangeTarget` at the exported version.
 */
export type XlsxAnchor =
  | { kind: 'sheet'; sheet: XlsxSheetIdentity }
  | { kind: 'cell'; sheet: XlsxSheetIdentity; a1: string }
  | { kind: 'range'; sheet: XlsxSheetIdentity; a1: string }
  | {
      kind: 'definedName';
      name: string;
      localSheet: XlsxSheetIdentity | null;
      /** Position in the workbook's defined-name list. */
      ordinal: number;
    }
  | { kind: 'sourcePart'; part: string; partSha256: string; path: number[] };

/** Retained source XML: part, SHA-256 of its bytes, element-child ordinals from its root. */
export interface XlsxSourcePart {
  part: string;
  partSha256: string;
  path: number[];
}

export type XlsxExportDiagnosticCode =
  | 'hidden-content-excluded'
  | 'visibility-unknown'
  | 'unsupported-sheet'
  | 'unsupported-content'
  | 'object-placeholder'
  | 'unreadable-object'
  | 'provenance-unavailable'
  | 'comments-omitted'
  | 'rich-text-omitted'
  | 'formatting-approximate'
  | 'formula-cache-unverified'
  | 'formula-result-missing'
  | 'calculation-failure'
  | 'value-unavailable'
  | 'truncated'
  | 'markdown-lossy';

export type XlsxExportDiagnostic = ExportDiagnostic<XlsxExportDiagnosticCode, XlsxAnchor>;

/** A cell value; `unavailable` stands for one JSON cannot carry, such as a non-finite number. */
export type XlsxExportValue = XlsxCellValue | { kind: 'unavailable'; reason: string };

/**
 * `unverified`: a result is stored, not checked. `missing`: none is stored, and the value is the
 * engine's placeholder. `uncertain`: presence cannot be established, as once anything has
 * calculated since a file that stored none was read. `cycle` and `limited`: what the last
 * calculation reported.
 */
export type XlsxFormulaResult = 'unverified' | 'missing' | 'uncertain' | 'cycle' | 'limited';

export interface XlsxExportCell {
  id: string;
  anchor: XlsxAnchor;
  value: XlsxExportValue;
  /** Without the leading `=`. */
  formula: string | null;
  /** The engine's formatting of the value, without column-width clipping or overflow. */
  displayText: string;
  numberFormat: string;
  /** `null` for cells without a formula. */
  formulaResult: XlsxFormulaResult | null;
  merge: { range: string; origin: boolean } | null;
}

export interface XlsxExportMerge {
  id: string;
  anchor: XlsxAnchor;
  /** Whether the merge extends past the selected range. */
  clipped: boolean;
}

/** A ListObject. */
export interface XlsxExportTable {
  id: string;
  anchor: XlsxAnchor;
  name: string;
  headerRows: number;
  totalsRows: number;
  columns: string[];
  clipped: boolean;
}

export interface XlsxExportHyperlink {
  id: string;
  anchor: XlsxAnchor;
  externalTarget: string | null;
  location: string | null;
  tooltip: string | null;
  display: string | null;
  clipped: boolean;
}

export type XlsxObjectKind =
  | 'chart'
  | 'picture'
  | 'shape'
  | 'group'
  | 'connector'
  | 'diagram'
  | 'graphicFrame'
  | 'contentPart'
  | 'unknown';

/** A drawing object as a placeholder: no chart data, image bytes or shape text. */
export interface XlsxExportObject {
  id: string;
  kind: XlsxObjectKind;
  /** The cells holding its corners; `null` when its grid position is unknown. */
  anchor: XlsxAnchor | null;
  name: string | null;
  altText: string | null;
  title: string | null;
  hidden: boolean;
  /** The chart or image part. */
  part: string | null;
  source: XlsxSourcePart | null;
}

export interface XlsxExportDefinedName {
  id: string;
  anchor: XlsxAnchor;
  name: string;
  formula: string;
  hidden: boolean;
  localSheet: XlsxSheetIdentity | null;
}

/**
 * One sheet. `truncated` means export stopped inside it: its lists are complete in the order
 * hidden spans, merges, tables, hyperlinks, objects, cells, up to where it stopped.
 */
export interface XlsxExportSheet extends ExportCompletion {
  id: string;
  anchor: XlsxAnchor;
  kind: 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet' | 'other';
  visibility: 'visible' | 'hidden' | 'veryHidden' | 'unknown';
  /** Stored cells and hyperlinks; drawings, merges and tables do not extend it. */
  usedRange: string | null;
  selectedRange: string | null;
  /** Hidden row spans such as `"5:7"` inside the selected range, exported or not. */
  hiddenRows: string[];
  /** Hidden column spans such as `"C:D"` inside the selected range, exported or not. */
  hiddenColumns: string[];
  merges: XlsxExportMerge[];
  tables: XlsxExportTable[];
  hyperlinks: XlsxExportHyperlink[];
  objects: XlsxExportObject[];
  /** Stored cells in row-major order, styled empty cells included. */
  cells: XlsxExportCell[];
  source: XlsxSourcePart | null;
}

export interface XlsxStructuredContent extends ExportCompletion {
  schemaVersion: 1;
  /** `session` for a live export, `snapshot` for bytes. */
  anchorScope: 'session' | 'snapshot';
  dateSystem: '1900' | '1904';
  calculation: { policy: 'asStored'; freshness: 'unverified' };
  included: {
    hiddenSheets: boolean;
    hiddenRows: boolean;
    hiddenColumns: boolean;
    definedNames: boolean;
    hiddenNames: boolean;
  };
  sheets: XlsxExportSheet[];
  definedNames: XlsxExportDefinedName[];
  diagnostics: XlsxExportDiagnostic[];
}

export interface XlsxMarkdownContent extends ExportCompletion {
  markdown: string;
  /** One per `<!-- xlsx-export:N -->` marker, in order. */
  anchors: MarkdownAnchor<XlsxAnchor>[];
  /** The structured export's diagnostics, then what Markdown lost. */
  diagnostics: XlsxExportDiagnostic[];
}

export type XlsxExportFailureCode = 'invalid-options' | 'invalid-scope' | 'limit-exceeded';

export interface XlsxExportFailure {
  code: XlsxExportFailureCode;
  /** The sheet a scope failure names, else `null`. */
  target: XlsxAnchor | null;
  message: string;
}

export type XlsxExportResult<T> =
  | { ok: true; version: string; content: T }
  | OperationRefusal<XlsxExportFailure>;
