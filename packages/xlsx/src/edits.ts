/**
 * Versioned reads and version-checked edit batches over one workbook. Every shape here is plain
 * JSON; the semantics live in Rust.
 *
 * Rows and columns are zero-based and ranges inclusive. Sheet ids come from the current catalog:
 * `sheet:{index}` standalone, the replica's sheet keys in collaboration. Versions and sheet ids
 * are session-scoped; neither survives save and reopen.
 */

import type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../shared/host-contracts/edits';
import type { NumberFormatMutation, RangeStylePatch } from './wasm/loader';

export type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../shared/host-contracts/edits';

/** Provenance only: it neither grants permission nor selects history. */
export type XlsxEditSource = 'host' | 'agent';

/**
 * `separate`: exactly one undo step. `none`: no undo step, existing entries kept; standalone
 * undo replays inverse operations, so undoing an older step that wrote the same cell still
 * overwrites it.
 */
export type XlsxEditHistory = 'separate' | 'none';

export interface XlsxCellPosition {
  row: number;
  col: number;
}

/**
 * An A1 cell or `A1:B2` range (`$` allowed), or zero-based corners. Sheet-qualified, union,
 * whole-row/column and named references are refused.
 */
export type XlsxRangeAddress =
  | { kind: 'a1'; a1: string }
  | { kind: 'rowCol'; start: XlsxCellPosition; end: XlsxCellPosition };

export interface XlsxRangeTarget {
  sheetId: string;
  range: XlsxRangeAddress;
}

/** One cell in canonical form. */
export interface XlsxCellAddress {
  sheetId: string;
  row: number;
  col: number;
  a1: string;
}

export type XlsxErrorValue =
  | '#DIV/0!'
  | '#N/A'
  | '#NAME?'
  | '#NULL!'
  | '#NUM!'
  | '#REF!'
  | '#VALUE!'
  | '#SPILL!'
  | '#CALC!';

export type XlsxCellValue =
  | { kind: 'empty' }
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; value: XlsxErrorValue };

/** Conditions on a cell's pre-batch state; an absent field imposes none. */
export interface XlsxCellGuard {
  /** The stored value, a formula's current result included. */
  value?: XlsxCellValue;
  /** Formula source without the leading `=`; `null` asserts there is no formula. */
  formula?: string | null;
  /** The exact text the engine formats the cell as. */
  displayText?: string;
}

/** Matrices are row-major and exactly the target's shape. */
export type XlsxEditOperation =
  | {
      /** What a user would type, parsed against each cell's current number format. */
      op: 'setCellInputs';
      target: XlsxRangeTarget;
      inputs: string[][];
    }
  | {
      /** Formula source without the leading `=`, stored as a formula whatever the format. */
      op: 'setFormulas';
      target: XlsxRangeTarget;
      formulas: string[][];
    }
  | {
      /** Relative mutations resolve against each cell's current format. */
      op: 'setNumberFormat';
      target: XlsxRangeTarget;
      format: NumberFormatMutation;
    }
  | { op: 'patchStyle'; target: XlsxRangeTarget; patch: RangeStylePatch };

export type XlsxEditStep = XlsxEditOperation & { expect?: { cells: XlsxCellGuard[][] } };

/** Requests above 16 MiB and results above 64 MiB are refused with `limit-exceeded`. */
export interface XlsxEditRequest {
  /** The version the targets were read at; a changed workbook refuses with `stale-version`. */
  expectVersion: string;
  /** Defaults to `host`. */
  source?: XlsxEditSource;
  /** Defaults to `separate`. */
  history?: XlsxEditHistory;
  /** Without `nowSerial`, volatile functions such as NOW() have no clock. */
  calculation?: { nowSerial?: number };
  steps: readonly XlsxEditStep[];
}

/**
 * `locked-target` covers any write to merged-cell followers, array-formula cells and protected
 * sheets. `read-only` comes only from an editor that does not accept edits.
 */
export type XlsxEditFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'content-mismatch'
  | 'overlapping-steps'
  | 'locked-target'
  | 'read-only'
  | 'unsupported'
  | 'invalid-step'
  | 'limit-exceeded';

export type XlsxEditFailure = OperationFailure<XlsxEditFailureCode, XlsxRangeTarget>;

export type XlsxEditRefusal = OperationRefusal<XlsxEditFailure>;

export interface XlsxEditReceipt {
  stepIndex: number;
  changed: boolean;
  target: XlsxRangeTarget;
  /** Row-major. */
  changedCells: XlsxCellAddress[];
}

/** What one step would do; it cannot be committed later. */
export interface XlsxEditPreview {
  stepIndex: number;
  target: XlsxRangeTarget;
  wouldChange: boolean;
  changedCellCount: number;
}

export interface XlsxEditCalculation {
  /** Dependents whose value moved. */
  changed: XlsxCellAddress[];
  cycleCells: XlsxCellAddress[];
  limitedCells: XlsxCellAddress[];
  /** Whether a list stopped at 10,000 cells. */
  truncated: boolean;
}

export type XlsxEditResult =
  | (EditSuccess<XlsxEditReceipt> & {
      source: XlsxEditSource;
      /** Sheets the batch or its recalculation changed, as ids in sheet order. */
      changedSheets: string[];
      calculation: XlsxEditCalculation;
    })
  | XlsxEditRefusal;

export type XlsxValidationResult = ValidationSuccess<XlsxEditPreview> | XlsxEditRefusal;

export interface XlsxReadRequest {
  /** Empty reads only the sheet catalog. */
  ranges: readonly XlsxRangeTarget[];
}

export interface XlsxSheetEntry {
  sheetId: string;
  name: string;
  /** False for sheets batches refuse to write: non-worksheets and protected worksheets. */
  editable: boolean;
}

export interface XlsxCellRead {
  a1: string;
  value: XlsxCellValue;
  formula: string | null;
  displayText: string;
}

export interface XlsxRangeRead {
  target: XlsxRangeTarget;
  /** Row-major. */
  cells: XlsxCellRead[][];
}

export interface XlsxReadCalculation {
  cycleCells: XlsxCellAddress[];
  limitedCells: XlsxCellAddress[];
  /** Whether a list stopped at 10,000 cells. */
  truncated: boolean;
}

export type XlsxReadResult =
  | {
      ok: true;
      version: string;
      sheets: XlsxSheetEntry[];
      ranges: XlsxRangeRead[];
      calculation: XlsxReadCalculation;
    }
  | XlsxEditRefusal;

/** An exact, case-sensitive substring search over display text. */
export interface XlsxFindRequest {
  text: string;
  /** Defaults to every sheet. */
  sheetIds?: readonly string[];
  /** Defaults to 100; at most 10,000. */
  limit?: number;
}

export interface XlsxFindMatch {
  cell: XlsxCellAddress;
  /** The cell's whole display text. */
  text: string;
}

export type XlsxFindResult =
  | { ok: true; version: string; matches: XlsxFindMatch[]; truncated: boolean }
  | XlsxEditRefusal;
