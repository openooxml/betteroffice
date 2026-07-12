/**
 * Pure a11y types: a serializable mirror of the visible grid the chrome renders
 * as an offscreen `role="grid"` for screen readers. Plain data only — the
 * builder takes label templates as a param rather than importing i18n, so this
 * seam stays dependency-free.
 */

/**
 * Label templates the builder interpolates. Field names mirror the `a11y` keys
 * in the framework adapter so a caller can pass localized strings through.
 * Placeholders: `{address}`, `{value}`, `{row}`, `{column}`.
 */
export interface A11yStrings {
  gridLabel: string;
  cellLabel: string;
  cellLabelSelected: string;
  emptyCellLabel: string;
  emptyCellLabelSelected: string;
  rowHeaderLabel: string;
  columnHeaderLabel: string;
}

/**
 * One mirrored cell: its address, display text, selection state, and the fully
 * interpolated aria label (e.g. `"B2, 42, selected"`).
 */
export interface A11yCell {
  row: number;
  col: number;
  address: string;
  text: string;
  selected: boolean;
  label: string;
}

/**
 * A column header: the letter (`"A"`) and its aria label (`"Column A"`).
 */
export interface A11yColumnHeader {
  col: number;
  text: string;
  label: string;
}

/**
 * One mirrored row: its 1-based header label and the cells left-to-right.
 */
export interface A11yRow {
  row: number;
  header: string;
  cells: A11yCell[];
}

/**
 * The full offscreen grid mirror for the current frame.
 */
export interface A11yGrid {
  label: string;
  sheetName: string;
  columnHeaders: A11yColumnHeader[];
  rows: A11yRow[];
}
