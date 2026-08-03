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
  clip?: Rect;
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
  clip?: Rect;
}

export type GeometryPathCommand =
  | { type: 'move'; x: number; y: number }
  | { type: 'line'; x: number; y: number }
  | { type: 'quad'; cpx: number; cpy: number; x: number; y: number }
  | {
      type: 'cubic';
      cp1x: number;
      cp1y: number;
      cp2x: number;
      cp2y: number;
      x: number;
      y: number;
    }
  | { type: 'close' };

export interface PathStroke {
  color: string;
  width: number;
}

export interface PathCmd {
  op: 'path';
  commands: GeometryPathCommand[];
  fill: string;
  stroke?: PathStroke;
  clip?: Rect;
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
  highlight?: string;
  dashedUnderline?: boolean;
  /** css/font family from the style font; the backend falls back to its default. */
  fontFamily?: string;
  /**
   * Preview text that is not the cell's committed value (a proposal ghost's
   * `new` text). Painted normally, but excluded from a11y text recovery.
   */
  ghost?: boolean;
  chart?: boolean;
}

/**
 * One draw command; discriminated on `op`.
 */
export type DrawCmd = FillRectCmd | LineCmd | PathCmd | TextCmd;

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
  rowIndices?: number[];
  colIndices?: number[];
  rowOffsets: number[];
  colOffsets: number[];
}

export interface HyperlinkRegion {
  top: number;
  left: number;
  bottom: number;
  right: number;
  externalTarget?: string;
  location?: string;
  tooltip?: string;
}

/**
 * A chart's placement in the frame: the id that addresses it across the wasm
 * boundary (the chart part's package path), what a screen reader reads, whether
 * the renderer managed to draw it, its full viewport-local rect and the visible
 * part after pane clipping. A chart that degraded to a placeholder still gets a
 * region, so it stays an addressable object on the sheet.
 *
 * Hand-mirrored from crates/xlsx-render/src/display_list.rs — keep in sync.
 */
export interface ChartRegion {
  id: string;
  label: string;
  /**
   * the chart could not be drawn; a neutral box occupies its rect instead.
   * skip-serialized at `false` in Rust, so treat absent as `false`.
   */
  placeholder?: boolean;
  rect: Rect;
  /** `rect` intersected with the pane band it paints in — the hit area. */
  clip: Rect;
  /** whether the anchor can be repinned; an absolute one cannot. */
  movable: boolean;
}

/** Former name of {@link ChartRegion}, when it carried the label alone. */
export type ChartA11yAttrs = ChartRegion;

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
  hyperlinks?: HyperlinkRegion[];
  charts?: ChartRegion[];
}
