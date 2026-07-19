/**
 * TS mirror of the Rust display list.
 *
 * Hand-mirrored from crates/xlsx-render/src/display_list.rs — keep in sync.
 * The renderer front-half (Rust) emits a target-agnostic list of draw commands;
 * a backend (Canvas2D in the browser, tiny-skia on servers) executes them. These
 * types are plain data with no DOM or framework dependency so the same list
 * flows across the wasm boundary as JSON and into either backend unchanged.
 */

/**
 * Axis-aligned rectangle in device-independent pixels.
 */
export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Horizontal alignment for a text command, matching the spreadsheet cell
 * alignment values the renderer resolves before emitting.
 */
export type TextAlign = 'left' | 'center' | 'right';

/**
 * Fill a solid rectangle — cell backgrounds, selection bands, header fills.
 */
export interface FillRectCmd {
  op: 'fillRect';
  x: number;
  y: number;
  w: number;
  h: number;
  /** css color string (`#rrggbb`, `rgba(...)`), resolved from theme + tint in Rust. */
  color: string;
}

/**
 * Stroke a straight line — gridlines, cell borders, frozen-pane dividers.
 */
export interface LineCmd {
  op: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
  color: string;
  /**
   * Stroke pattern. Absent/`undefined` is solid; `'dashed'`/`'dotted'` apply a
   * dash pattern; `'double'` draws two thin parallel passes offset by ±. Skip-
   * serialized at its default in Rust, so a solid line omits the field.
   */
  style?: 'dashed' | 'dotted' | 'double';
}

/**
 * Paint a single-line text run. Spreadsheet cell text is single-line and
 * clipped to the cell box, so `clip` is the cell rect and `align` places the
 * run within it. Text metrics are owned upstream in Rust.
 */
export interface TextCmd {
  op: 'text';
  x: number;
  y: number;
  text: string;
  fontSize: number;
  /** resolved font color (`#rrggbb`); a number-format color prefix wins upstream. */
  color: string;
  /** clip rectangle; the backend saves/clips/restores around the fill. */
  clip?: Rect;
  align?: TextAlign;
  /**
   * Font style facets resolved from the cell's style. All skip-serialized at
   * `false` in Rust, so an unstyled run omits them; treat absent as `false`.
   */
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** css/font family from the style font; the backend falls back to its default. */
  fontFamily?: string;
  /**
   * Preview text that is not the cell's committed value (a proposal ghost's
   * `new` text). Painted normally, but excluded from a11y text recovery.
   */
  ghost?: boolean;
}

/**
 * One draw command; discriminated on `op`.
 */
export type DrawCmd = FillRectCmd | LineCmd | TextCmd;

/**
 * Grid metadata for the frame: which sheet cells the visible tracks map to and
 * their viewport-local pixel boundaries. `rowOffsets[i]`/`colOffsets[i]` is the
 * leading edge of the i-th visible row/col in device-independent px from the
 * frame origin; both arrays have length `visible count + 1`, the last entry
 * being one-past-end (the trailing edge of the last visible track). The pure
 * hit-test and a11y seams read this to place clicks and mirror the grid without
 * re-deriving geometry.
 *
 * Hand-mirrored from crates/xlsx-render/src/display_list.rs — keep in sync.
 */
export interface GridMeta {
  startRow: number;
  startCol: number;
  rowOffsets: number[];
  colOffsets: number[];
}

/**
 * A full frame to paint: logical size plus the ordered command stream. `grid`
 * is optional so a synthetic or pre-grid-metadata frame still type-checks; the
 * hit-test and a11y builders treat its absence as "no addressable cells".
 */
export interface DisplayList {
  width: number;
  height: number;
  commands: DrawCmd[];
  grid?: GridMeta;
}
