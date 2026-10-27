// @betteroffice/xlsx "./headless" public API surface
// Regenerate: bun scripts/api-surface.ts --write xlsx
// ---- a11y/index.d.ts ----
/**
 * Pure a11y grid builder: turns a display list plus the current selection into a
 * serializable {@link A11yGrid} the chrome renders offscreen for screen readers.
 * Cell text is recovered from the frame's clipped text commands, keyed by cell
 * box origin; geometry and addresses come from the frame's grid metadata. No
 * DOM, no i18n import — label templates arrive as a param.
 */
import type { DisplayList } from '../display-list/types';
import type { Selection } from '../selection/index';
import type { A11yGrid, A11yStrings } from './types';
export type { A11yCell, A11yChart, A11yColumnHeader, A11yGrid, A11yRow, A11yStrings, } from './types';
/**
 * Build the offscreen grid mirror for a frame. Returns an empty grid (no rows)
 * when the frame carries no grid metadata, so the chrome can render nothing
 * rather than guess geometry.
 */
export declare function buildA11yGrid(displayList: DisplayList, selection: Selection | null, sheetName: string, t: A11yStrings): A11yGrid;

// ---- a11y/types.d.ts ----
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
export interface A11yChart {
    label: string;
    /** the chart could not be drawn; a neutral box occupies its rect instead. */
    placeholder: boolean;
}
/**
 * The full offscreen grid mirror for the current frame.
 */
export interface A11yGrid {
    label: string;
    sheetName: string;
    columnHeaders: A11yColumnHeader[];
    rows: A11yRow[];
    charts: A11yChart[];
}

// ---- clipboard/index.d.ts ----
/**
 * Pure TSV codec for clipboard interchange, in the Excel dialect (tab-separated
 * fields, CRLF rows, `"`-quoted fields with doubled embedded quotes). No DOM —
 * the chrome hands the system clipboard the string this produces and feeds paste
 * text back through {@link fromTsv}.
 *
 * Security: a foreign spreadsheet treats
 * a pasted cell as a formula when its text leads with `=`, `+`, `-`, `@`, or a
 * control char that can break the field. We defang such *text* with a leading
 * single quote so it pastes literally — but we export our own real formulas
 * (`isFormula`) verbatim, because faithfully round-tripping a `=SUM(...)` the
 * user actually wrote is the point, and it is our value, not attacker text.
 */
/**
 * One source cell as the wasm range read returns it: the raw entered `input`
 * (a `=...` string when `isFormula`) plus whether the model holds it as a
 * formula.
 */
export interface CellInput {
    input: string;
    isFormula: boolean;
}
/**
 * Serialize a row-major grid of cells to a TSV string for the system clipboard,
 * defanging formula-looking text (see the module security note) while exporting
 * genuine formulas as-is.
 */
export declare function toTsv(cells: CellInput[][]): string;
/**
 * Parse clipboard TSV back into a row-major grid of raw strings. Handles the
 * Excel dialect: `"`-quoted fields with doubled quotes and embedded tabs/
 * newlines, CRLF or LF rows, and a single trailing newline (which does not add
 * an empty row). The formula-injection guard quote is left in place — it is part
 * of the literal text a defanged cell round-trips to.
 */
export declare function fromTsv(text: string): string[][];

// ---- display-list/types.d.ts ----
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
export type PatternType = 'gray125' | 'gray0625' | 'mediumGray' | 'darkGray' | 'lightGray' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis';
export interface PatternFillCmd {
    op: 'patternFill';
    x: number;
    y: number;
    w: number;
    h: number;
    pattern: PatternType;
    fg: string;
    bg: string;
    originX: number;
    originY: number;
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
    style?: 'dashed' | 'dotted' | 'double' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
    /**
     * Rail width of a `'double'` line: two rails this wide on the outer edges of
     * `width`. Absent, painters keep the legacy rails derived from `width`.
     */
    rail?: number;
    clip?: Rect;
}
export type GeometryPathCommand = {
    type: 'move';
    x: number;
    y: number;
} | {
    type: 'line';
    x: number;
    y: number;
} | {
    type: 'quad';
    cpx: number;
    cpy: number;
    x: number;
    y: number;
} | {
    type: 'cubic';
    cp1x: number;
    cp1y: number;
    cp2x: number;
    cp2y: number;
    x: number;
    y: number;
} | {
    type: 'close';
};
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
 * Wrapped cell text: the painter breaks the run into lines no wider than
 * `width` and steps their baselines by `lineHeight`, the block anchored by
 * `valign` on the single-line baseline `y`.
 */
export interface TextWrap {
    width: number;
    lineHeight: number;
    valign: 'top' | 'center' | 'bottom';
}
/**
 * Paint a text run clipped to the cell box: `clip` is the cell rect and
 * `align` places the run within it. A run without `wrap` is one line.
 */
export interface TextCmd {
    op: 'text';
    x: number;
    y: number;
    text: string;
    /** Font size in points. */
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
    numberFormat?: NumberFormatLayout;
    /**
     * Preview text that is not the cell's committed value (a proposal ghost's
     * `new` text). Painted normally, but excluded from a11y text recovery.
     */
    ghost?: boolean;
    chart?: boolean;
    rotationDeg?: number;
    rotationBox?: Rect;
    rotationValign?: 'top' | 'center' | 'bottom';
    stacked?: boolean;
    wrap?: TextWrap;
}
export interface NumberFormatLayout {
    left: number;
    width: number;
    parts: {
        offset: number;
        kind: 'padding' | 'fill';
        character: string;
    }[];
    general?: number;
    overflow?: boolean;
}
/**
 * One draw command; discriminated on `op`.
 */
export type DrawCmd = FillRectCmd | PatternFillCmd | LineCmd | PathCmd | TextCmd;
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
    hiddenRowSpans?: [number, number][];
    hiddenColSpans?: [number, number][];
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
 * boundary, what a screen reader reads, whether the renderer managed to draw
 * it, its full viewport-local rect and the visible part after pane clipping. A
 * chart that degraded to a placeholder still gets a region, so it stays an
 * addressable object on the sheet.
 *
 * Hand-mirrored from crates/xlsx-render/src/display_list.rs — keep in sync.
 */
export interface ChartRegion {
    /**
     * the drawing part and the anchor in it, unique within a sheet. Opaque —
     * never the chart part's path, which two anchors may share.
     */
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
    /** The cell budget omitted the bottom of this viewport. */
    truncated?: boolean;
}

// ---- headless.d.ts ----
/**
 * Headless entry — the DOM-free compute surface for server rendering, tests,
 * and workers. Re-exports only pure viewport math and the display-list types;
 * no canvas, no wasm loader, no framework. The seam-purity test walks the
 * import closure from this file and fails if anything impure creeps in.
 */
export type { Rect, TextAlign, FillRectCmd, PatternFillCmd, PatternType, LineCmd, GeometryPathCommand, PathStroke, PathCmd, TextCmd, DrawCmd, DisplayList, GridMeta, HyperlinkRegion, ChartRegion, ChartA11yAttrs, } from './display-list/types';
export type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './viewport/index';
export { trackCount, totalExtent, visibleRange, visibleCells, clampScroll, uniformOffsets, } from './viewport/index';
export type { CellAddr, Selection, CellRange, Direction, SelectionLimits, SelectionData, SelectionHidden, SelectionMerges, MoveOpts, KeyInput, SelectionAction, } from './selection/index';
export { selectionAt, normalizeRange, rangeContains, cellCount, extendTo, moveFocus, selectionKeyReducer, } from './selection/index';
export { cellAtPoint, cellRect, chartRegionAtPoint, rangeRect } from './hittest/index';
export type { HyperlinkDestination } from './hyperlinks/index';
export { hyperlinkAtCell, parseHyperlinkLocation, safeExternalHyperlink, } from './hyperlinks/index';
export type { CellInput } from './clipboard/index';
export { toTsv, fromTsv } from './clipboard/index';
export type { A11yStrings, A11yCell, A11yChart, A11yColumnHeader, A11yRow, A11yGrid, } from './a11y/index';
export { buildA11yGrid } from './a11y/index';
export type EngineBackend = 'wasm' | 'native';
/** Which engine runs this workbook code ('native' when @betteroffice/xlsx-native is installed and matches this version on Node, Bun or Deno, else 'wasm'). */
export declare function engineBackend(): EngineBackend;

// ---- hittest/index.d.ts ----
/**
 * Pure hit-testing over a frame's {@link GridMeta}: map a viewport-local point
 * to a cell, and map cells/ranges back to their pixel rects for overlay
 * placement. Binary search over the boundary offsets; DOM-free so click
 * handling and the selection overlay share one geometry source with the painter.
 */
import type { ChartRegion, GridMeta, Rect } from '../display-list/types';
import type { CellAddr, CellRange } from '../selection/index';
/**
 * The topmost chart whose visible region contains a viewport-local point, or
 * `null`. Mirrors `chart_at_point` in crates/xlsx-render/src/hit.rs: regions are
 * in paint order so the last match wins, and the clipped region bounds the hit,
 * so a chart scrolled under a frozen pane is not hit where it is hidden.
 *
 * Every coordinate is rounded to f32 first. The regions arrive as f32 off the
 * wire, and the engine compares against `clip.x + clip.w` computed in f32; the
 * same sum in f64 lands on the other side of the edge at reachable coordinates.
 */
export declare function chartRegionAtPoint(charts: readonly ChartRegion[] | undefined, x: number, y: number): ChartRegion | null;
/**
 * The cell under a viewport-local point, or `null` when the point is outside the
 * visible grid (or the frame carries no grid metadata).
 */
export declare function cellAtPoint(grid: GridMeta | undefined, x: number, y: number): CellAddr | null;
/**
 * The pixel rect of a single cell in viewport-local coordinates, or `null` when
 * the cell is not within the visible window.
 */
export declare function cellRect(grid: GridMeta | undefined, row: number, col: number): Rect | null;
/**
 * The pixel rect covering the visible part of a cell range, clipped to the
 * window. `null` when the range lies entirely outside the visible grid — so the
 * selection overlay is hidden rather than mispainted when scrolled away.
 */
export declare function rangeRect(grid: GridMeta | undefined, range: CellRange): Rect | null;

// ---- hyperlinks/index.d.ts ----
import type { DisplayList, HyperlinkRegion } from '../display-list/types';
export interface HyperlinkDestination {
    sheetName: string;
    row: number;
    col: number;
}
export declare function hyperlinkAtCell(displayList: DisplayList, row: number, col: number): HyperlinkRegion | null;
export declare function safeExternalHyperlink(link: HyperlinkRegion): string | null;
export declare function parseHyperlinkLocation(location: string, currentSheet: string): HyperlinkDestination | null;

// ---- selection/index.d.ts ----
/**
 * Pure selection seam: the anchor/focus model, its transforms, and the keyboard
 * reducer. DOM-free by construction (see the seam-purity test).
 */
export type { CellAddr, Selection, CellRange, Direction, SelectionLimits, SelectionData, SelectionHidden, SelectionMerges, MoveOpts, KeyInput, SelectionAction, } from './types';
export { selectionAt, normalizeRange, rangeContains, cellCount, extendTo, moveFocus, } from './model';
export { selectionKeyReducer } from './keyboard';

// ---- selection/keyboard.d.ts ----
/**
 * Keyboard reducer: maps a keystroke over the current selection to a
 * declarative {@link SelectionAction}. Pure — it neither mutates state nor
 * touches the DOM; the chrome owns the keydown listener and interprets the
 * action. Excel navigation semantics, with two deliberate choices documented at
 * the F2/Enter handlers below.
 */
import type { KeyInput, Selection, SelectionAction, SelectionLimits, SelectionData, SelectionHidden, SelectionMerges } from './types';
/**
 * Reduce a keystroke to a selection action against the current selection and
 * grid limits. Returns `{ type: 'none' }` for keys this layer does not own so
 * the chrome can fall through to its own handling.
 */
export declare function selectionKeyReducer(state: Selection, key: KeyInput, limits: SelectionLimits, data?: SelectionData, hidden?: SelectionHidden, merges?: SelectionMerges): SelectionAction;

// ---- selection/model.d.ts ----
/**
 * Selection model: pure transforms over {@link Selection} data. Every function
 * returns a fresh value and never mutates its input, so state lives with the
 * chrome and this layer stays a deterministic reducer core.
 */
import type { CellAddr, CellRange, Direction, MoveOpts, Selection, SelectionLimits, SelectionMerges } from './types';
/**
 * A collapsed selection at a single cell.
 */
export declare function selectionAt(addr: CellAddr): Selection;
/**
 * The explicit bounds or normalized rectangle spanned by the endpoints.
 */
export declare function normalizeRange(sel: Selection): CellRange;
/**
 * Whether a cell falls inside a selection's normalized range.
 */
export declare function rangeContains(sel: Selection, addr: CellAddr): boolean;
/**
 * Number of cells covered by a selection.
 */
export declare function cellCount(sel: Selection): number;
/**
 * Extend the selection to a new focus, holding the anchor fixed (shift-click /
 * shift-arrow semantics). The address is clamped into the grid when limits pass.
 */
export declare function extendTo(sel: Selection, addr: CellAddr, limits?: SelectionLimits): Selection;
/**
 * The anchor of the merged area covering `addr`, or `addr` itself.
 */
export declare function mergeAnchor(addr: CellAddr, merges?: SelectionMerges): CellAddr;
/**
 * Move the focus one (or `step`) cells in a direction. When `extend` is set the
 * anchor holds and the range grows; otherwise the whole selection collapses onto
 * the new focus. The result is clamped to the grid bounds in `opts.limits`. A
 * merged area counts as one cell: the move starts from its far edge, and a
 * collapsed selection that lands inside one lands on its anchor.
 */
export declare function moveFocus(sel: Selection, direction: Direction, opts: MoveOpts): Selection;

// ---- selection/types.d.ts ----
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
/** Sheet bounds and the viewport height in whole rows. */
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
    merges?: SelectionMerges;
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
export type SelectionAction = {
    type: 'move';
    selection: Selection;
} | {
    type: 'startEdit';
    initialInput?: string;
} | {
    type: 'clear';
} | {
    type: 'none';
};
/** Precomputed model answers for one keyboard navigation request. */
export interface SelectionData {
    merge?: CellRange | null;
    edge?: CellAddr | null;
    region?: CellRange | null;
    usedRange?: CellRange | null;
}
export interface SelectionHidden {
    isRowHidden(row: number): boolean;
    isColHidden(col: number): boolean;
}
/** Merged areas, so navigation treats each one as a single cell. */
export interface SelectionMerges {
    mergeAt(row: number, col: number): CellRange | null;
}

// ---- viewport/index.d.ts ----
/**
 * Pure viewport seam: virtualization math and its types. DOM-free by
 * construction (see the seam-purity test).
 */
export type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './types';
export { trackCount, totalExtent, visibleRange, visibleCells, clampScroll, uniformOffsets, } from './math';

// ---- viewport/math.d.ts ----
/**
 * Pure viewport math: visible-range computation against cumulative track
 * offsets, and scroll clamping. No DOM, no side effects — every function is a
 * deterministic transform of plain numbers so it is golden-testable.
 */
import type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './types';
/**
 * Number of tracks described by an offsets array.
 */
export declare function trackCount(offsets: TrackOffsets): number;
/**
 * Total pixel extent of a track axis (position of the trailing edge).
 */
export declare function totalExtent(offsets: TrackOffsets): number;
/**
 * The inclusive range of scrollable (non-frozen) tracks intersecting the
 * viewport. `frozen` tracks at the head are pinned and excluded here — the
 * scrollable window starts after them and is shortened by their pinned width.
 */
export declare function visibleRange(offsets: TrackOffsets, frozen: number, scroll: number, viewportExtent: number): VisibleRange;
/**
 * Visible column and row ranges for a frame, frozen tracks excluded.
 */
export declare function visibleCells(state: ViewportState, colOffsets: TrackOffsets, rowOffsets: TrackOffsets): VisibleCells;
/**
 * Clamp a viewport's scroll offsets into the valid range for the given content.
 * Returns the same state when already in range, a corrected copy otherwise.
 */
export declare function clampScroll(state: ViewportState, colOffsets: TrackOffsets, rowOffsets: TrackOffsets): ViewportState;
/**
 * Build a uniform offsets array of `count` tracks each `size` px wide.
 */
export declare function uniformOffsets(count: number, size: number): TrackOffsets;

// ---- viewport/types.d.ts ----
/**
 * Pure viewport types shared by the accessibility grid and canvas painter.
 */
/**
 * Cumulative pixel offsets for a track axis (columns or rows). `offsets[i]` is
 * the leading edge of track `i`; `offsets.length === trackCount + 1`, so the
 * last element is the total extent. Frozen tracks are the first `frozen` of
 * these — they are pinned and always visible.
 */
export type TrackOffsets = number[];
/**
 * Scroll + geometry state driving virtualization. `scrollX/scrollY` are in
 * content pixels of the scrollable (non-frozen) region; `width/height` is the
 * visible viewport in css pixels; `dpr` scales to device pixels at paint time.
 */
export interface ViewportState {
    scrollX: number;
    scrollY: number;
    width: number;
    height: number;
    dpr: number;
    frozenRows: number;
    frozenCols: number;
}
/**
 * An inclusive track index range. `last < first` (e.g. `{ first: 0, last: -1 }`)
 * denotes an empty range.
 */
export interface VisibleRange {
    first: number;
    last: number;
}
/**
 * The visible track window on both axes, frozen tracks excluded (they always
 * render). Feed to the renderer to build the display list for a frame.
 */
export interface VisibleCells {
    cols: VisibleRange;
    rows: VisibleRange;
}
