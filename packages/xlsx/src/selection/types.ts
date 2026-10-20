/**
 * Pure selection types. Plain data only — no DOM, no framework — so the model
 * is golden-testable. The chrome layer interprets the
 * declarative {@link SelectionAction}s the reducer emits.
 */

/**
 * A single cell coordinate. `row`/`col` are zero-based sheet indices, not px.
 */
export interface CellAddr {
  row: number;
  col: number;
}

/**
 * `anchor` is the fixed starting cell; `focus` is the active cell.
 * `range` overrides their bounds for region selection.
 */
export interface Selection {
  anchor: CellAddr;
  focus: CellAddr;
  /** Explicit bounds when the active cell is inside a selected region. */
  range?: CellRange;
}

/**
 * A normalized, inclusive rectangle of cells. `top <= bottom`, `left <= right`.
 * Derived from a {@link Selection} by {@link normalizeRange}; the natural shape
 * for hit-testing, overlay rects, and range reads.
 */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * A cardinal move direction for {@link moveFocus} and arrow-key handling.
 */
export type Direction = 'up' | 'down' | 'left' | 'right';

/**
 * Grid bounds a move must stay inside: the used range is `[0, rows)` × `[0,
 * cols)`, and `rowsPerPage` is the viewport height in whole rows for
 * PageUp/PageDown steps.
 */
export interface SelectionLimits {
  rows: number;
  cols: number;
  rowsPerPage: number;
}

/**
 * Options for {@link moveFocus}: whether to extend (keep the anchor) or collapse
 * to a single cell, how many cells to step, and the grid bounds to clamp into.
 */
export interface MoveOpts {
  extend?: boolean;
  step?: number;
  limits: SelectionLimits;
  hidden?: SelectionHidden;
}

/**
 * A keyboard event reduced to the fields the reducer reads. `metaKey`/`ctrlKey`
 * are both accepted so a caller need not normalize mac cmd vs win/linux ctrl.
 */
export interface KeyInput {
  key: string;
  shiftKey?: boolean;
  metaKey?: boolean;
  ctrlKey?: boolean;
  altKey?: boolean;
}

/**
 * The declarative outcome of a keystroke. The chrome interprets it: `move`
 * replaces the selection, `startEdit` opens the in-cell editor (seeding it with
 * `initialInput` when a printable key triggered it), `clear` deletes the
 * selected cells' contents, and `none` means the key was not handled here.
 */
export type SelectionAction =
  | { type: 'move'; selection: Selection }
  | { type: 'startEdit'; initialInput?: string }
  | { type: 'clear' }
  | { type: 'none' };

/** Precomputed model answers for one keyboard navigation request. */
export interface SelectionData {
  edge?: CellAddr | null;
  region?: CellRange | null;
}

export interface SelectionHidden {
  isRowHidden(row: number): boolean;
  isColHidden(col: number): boolean;
}
