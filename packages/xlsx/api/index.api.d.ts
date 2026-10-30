// @betteroffice/xlsx "." public API surface
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

// ---- charts.d.ts ----
export interface ChartSeriesInfo {
    index: number;
    name: string | null;
    nameRef: string | null;
    categoryRef: string | null;
    valueRef: string | null;
}
export interface ChartInfo {
    id: string;
    index: number;
    part: string;
    type: string;
    anchorRange: string | null;
    /** false when the chart part cannot be read; error says why and the rest is empty. */
    editable: boolean;
    error: string | null;
    title: string | null;
    /** a title Excel fills in, with no text of its own. */
    autoTitle: boolean;
    richTitle: boolean;
    series: ChartSeriesInfo[];
}
export type ChartSeriesName = {
    text: string;
} | {
    ref: string;
};
export type ChartEdit = {
    action: 'set_chart_title';
    title: string | null;
} | {
    action: 'set_series';
    series: number;
    name?: ChartSeriesName;
    categories?: string;
    values?: string;
};
export type ChartMutation = {
    sheet: number;
    chart: number;
} & ChartEdit;

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

// ---- collaboration/asyncReplica.d.ts ----
import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
import type { AwarenessCursor, AwarenessPeer, CollaborationReplica, CollaborationUpdateOrigin } from './types';
/** @experimental */
export interface AsyncCollaborationReplica {
    readonly clientId: number;
    encodeStateVector(): Promise<Uint8Array>;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Promise<Uint8Array>;
    applyUpdate(update: Uint8Array): Promise<unknown>;
    /** Applies updates in order as one merge; rejects at the first refused update, keeping the ones before it. */
    applyUpdateChunks?(updates: readonly Uint8Array[]): Promise<unknown>;
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}
/** @experimental */
export interface AsyncReplicaAdapter extends CollaborationReplica {
    readonly async: AsyncCollaborationReplica;
}
export declare function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined;
/** @internal */
export declare function attachCollaborationPipe(adapter: AsyncReplicaAdapter, channel: CollaborationPipeChannel<AwarenessCursor, AwarenessPeer> | undefined): void;
/** @internal */
export declare function collaborationPipeFor(replica: CollaborationReplica): CollaborationPipeChannel<AwarenessCursor, AwarenessPeer> | undefined;
/** @experimental */
export declare function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter;

// ---- collaboration/types.d.ts ----
export type CollaborationUpdateOrigin = 'local' | 'remote';
export interface CollaborationReplica {
    readonly clientId: number;
    encodeStateVector(): Uint8Array;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    applyUpdate(update: Uint8Array): unknown;
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}
export type CollaborationTransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
} | {
    type: 'drain';
};
export interface CollaborationTransport {
    /** Starts one attempt; after close, reopen internally or require another provider.connect(). */
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    /** Returns true only when the owned frame was accepted; false requires a later drain event. */
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: CollaborationTransportEvent) => void): () => void;
}
export type CollaborationStatus = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
export interface CollaborationStatusChange {
    status: CollaborationStatus;
    synced: boolean;
}
export interface CollaborationUser {
    name: string;
    color: string;
}
export interface CollaborationUserOptions {
    name: string;
    color?: string;
}
export interface AwarenessCell {
    row: number;
    col: number;
}
export interface AwarenessCursor {
    sheet: string;
    anchor: AwarenessCell;
    head: AwarenessCell;
}
export interface AwarenessPayload {
    user: CollaborationUser;
    cursor: AwarenessCursor | null;
}
export interface AwarenessState extends AwarenessPayload {
    clientId: number;
    clock: number;
}
export interface AwarenessUpdate {
    clientId: number;
    clock: number;
    state: AwarenessPayload | null;
}
export interface AwarenessPeer extends AwarenessState {
    lastSeen: number;
    cursorMovedAt: number;
}
export interface CollaborationAwareness {
    readonly peers: readonly AwarenessPeer[];
    setCursor(cursor: AwarenessCursor | null): void;
    onAwareness(listener: AwarenessListener): () => void;
}
export type CollaborationErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export declare class CollaborationError extends Error {
    readonly code: CollaborationErrorCode;
    readonly cause: unknown;
    constructor(code: CollaborationErrorCode, message: string, cause?: unknown);
}
export interface CollaborationProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    maxPendingBytes?: number;
    user?: CollaborationUserOptions;
}
export type CollaborationStatusListener = (change: CollaborationStatusChange) => void;
export type CollaborationErrorListener = (error: CollaborationError) => void;
export type AwarenessListener = (peers: readonly AwarenessPeer[]) => void;

// ---- definedNames.d.ts ----
export type XlsxDefinedNameOp = {
    type: 'setDefinedName';
    name: string;
    refers_to: string;
    scope?: number | null;
    hidden?: boolean;
} | {
    type: 'deleteDefinedName';
    name: string;
    scope?: number | null;
};
export interface XlsxDefinedName {
    name: string;
    refersTo: string;
    scope: number | null;
    hidden: boolean;
}
export interface XlsxDefinedNamesRequest {
    /** Omitted selects all scopes; null selects workbook scope. */
    scope?: number | null;
    offset?: number;
    limit?: number;
}
export interface XlsxDefinedNamesResult {
    version: string;
    items: XlsxDefinedName[];
    total: number;
    nextOffset: number | null;
}

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

// ---- edits.d.ts ----
/**
 * Versioned reads and version-checked edit batches over one workbook. Every shape here is plain
 * JSON; the semantics live in Rust.
 *
 * Rows and columns are zero-based and ranges inclusive. Sheet ids come from the current catalog:
 * `sheet:{index}` standalone, the replica's sheet keys in collaboration. Versions and sheet ids
 * are session-scoped; neither survives save and reopen. Standalone ids are positional, so each is
 * valid only for the version it was read at.
 */
import type { EditSuccess, OperationFailure, OperationRefusal, ValidationSuccess } from '../../../shared/host-contracts/edits';
import type { XlsxSheetVisibility } from './exports';
import type { NumberFormatMutation, RangeStylePatch } from './wasm/loader';
export type { EditSuccess, OperationFailure, OperationRefusal, ValidationSuccess, } from '../../../shared/host-contracts/edits';
/** Provenance only: it neither grants permission nor selects history. */
export type XlsxEditSource = 'host' | 'agent';
/**
 * `separate`: exactly one undo step. `none`: no undo step, existing entries kept; standalone
 * undo replays inverse operations, so undoing an older step that wrote the same cell still
 * overwrites it.
 *
 * @experimental `none` only: its interaction with standalone undo may change in a minor release.
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
export type XlsxRangeAddress = {
    kind: 'a1';
    a1: string;
} | {
    kind: 'rowCol';
    start: XlsxCellPosition;
    end: XlsxCellPosition;
};
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
export type XlsxErrorValue = '#DIV/0!' | '#N/A' | '#NAME?' | '#NULL!' | '#NUM!' | '#REF!' | '#VALUE!' | '#SPILL!' | '#CALC!';
export type XlsxCellValue = {
    kind: 'empty';
} | {
    kind: 'number';
    value: number;
} | {
    kind: 'text';
    value: string;
} | {
    kind: 'bool';
    value: boolean;
} | {
    kind: 'error';
    value: XlsxErrorValue;
};
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
export type XlsxEditOperation = {
    /** What a user would type, parsed against each cell's current number format. */
    op: 'setCellInputs';
    target: XlsxRangeTarget;
    inputs: string[][];
} | {
    op: 'setCellValues';
    target: XlsxRangeTarget;
    values: XlsxCellValue[][];
} | {
    /** Formula source without the leading `=`, stored as a formula whatever the format. */
    op: 'setFormulas';
    target: XlsxRangeTarget;
    formulas: string[][];
} | {
    /** Relative mutations resolve against each cell's current format. */
    op: 'setNumberFormat';
    target: XlsxRangeTarget;
    format: NumberFormatMutation;
} | {
    op: 'patchStyle';
    target: XlsxRangeTarget;
    patch: RangeStylePatch;
};
export type XlsxEditStep = XlsxEditOperation & {
    expect?: {
        cells: XlsxCellGuard[][];
    };
};
/** Requests above 16 MiB and results above 64 MiB are refused with `limit-exceeded`. */
export interface XlsxEditRequest {
    /** The version the targets were read at; a changed workbook refuses with `stale-version`. */
    expectVersion: string;
    /** Defaults to `host`. */
    source?: XlsxEditSource;
    /** Defaults to `separate`. */
    history?: XlsxEditHistory;
    /** Overrides the handle clock for this request; without either clock, NOW() has no clock. */
    calculation?: {
        nowSerial?: number;
    };
    steps: readonly XlsxEditStep[];
}
/**
 * `locked-target` covers any write to merged-cell followers, array-formula cells and locked
 * cells of protected sheets, and formats on protected sheets that do not allow formatting cells.
 * `read-only` comes only from an editor that does not accept edits.
 */
export type XlsxEditFailureCode = 'stale-version' | 'missing-target' | 'content-mismatch' | 'overlapping-steps' | 'locked-target' | 'read-only' | 'unsupported' | 'invalid-step' | 'limit-exceeded';
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
export type XlsxEditResult = (EditSuccess<XlsxEditReceipt> & {
    source: XlsxEditSource;
    /** Sheets the batch or its recalculation changed, as ids in sheet order. */
    changedSheets: string[];
    calculation: XlsxEditCalculation;
}) | XlsxEditRefusal;
export type XlsxValidationResult = ValidationSuccess<XlsxEditPreview> | XlsxEditRefusal;
export interface XlsxReadRequest {
    /** Empty reads only the sheet catalog. */
    ranges: readonly XlsxRangeTarget[];
}
/** Stored text with xstring escapes decoded; printable sections use dialog field tokens and normalized line breaks. */
export interface XlsxHeaderFooterText {
    raw: string;
    left: string;
    center: string;
    right: string;
}
/** Absent or empty header and footer parts are null. */
export interface XlsxHeaderFooterPair {
    header: XlsxHeaderFooterText | null;
    footer: XlsxHeaderFooterText | null;
}
/** Printable variants; even and first are null unless their respective flags are enabled. */
export interface XlsxSheetHeaderFooter {
    differentFirst: boolean;
    differentOddEven: boolean;
    scaleWithDoc: boolean;
    alignWithMargins: boolean;
    odd: XlsxHeaderFooterPair;
    even: XlsxHeaderFooterPair | null;
    first: XlsxHeaderFooterPair | null;
}
export interface XlsxSheetEntry {
    /** Null means known absent or empty; omitted means source facts are unknown. */
    headerFooter?: XlsxSheetHeaderFooter | null;
    visibility?: XlsxSheetVisibility;
    sheetId: string;
    name: string;
    /**
     * False for non-worksheets, which batches refuse to write, and protected worksheets, which take
     * content only in unlocked cells and formats only when they allow formatting cells.
     */
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
export type XlsxReadResult = {
    ok: true;
    version: string;
    sheets: XlsxSheetEntry[];
    ranges: XlsxRangeRead[];
    calculation: XlsxReadCalculation;
} | XlsxEditRefusal;
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
export type XlsxFindResult = {
    ok: true;
    version: string;
    matches: XlsxFindMatch[];
    truncated: boolean;
} | XlsxEditRefusal;

// ---- exports.d.ts ----
/**
 * Read-only structured export of workbook content and its Markdown projection. Every shape
 * here is plain JSON; the walker, limits and Markdown renderer live in Rust.
 *
 * Export never recalculates: formula results are the stored values (`calculation.policy:
 * 'asStored'`, `freshness: 'unverified'`). Anchors name positions by sheet id, index and name
 * and A1 in the exported version or snapshot; none follows later row, column or sheet edits.
 */
import type { OperationRefusal } from '../../../shared/host-contracts/edits';
import type { ExportCompletion, ExportDiagnostic, MarkdownAnchor } from '../../../shared/host-contracts/exports';
import type { XlsxCellValue } from './edits';
export type { ExportCompletion, ExportDiagnostic, ExportSeverity, MarkdownAnchor, } from '../../../shared/host-contracts/exports';
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
export type XlsxAnchor = {
    kind: 'sheet';
    sheet: XlsxSheetIdentity;
} | {
    kind: 'cell';
    sheet: XlsxSheetIdentity;
    a1: string;
} | {
    kind: 'range';
    sheet: XlsxSheetIdentity;
    a1: string;
} | {
    kind: 'definedName';
    name: string;
    localSheet: XlsxSheetIdentity | null;
    /** Position in the workbook's defined-name list. */
    ordinal: number;
} | {
    kind: 'sourcePart';
    part: string;
    partSha256: string;
    path: number[];
};
/** Retained source XML: part, SHA-256 of its bytes, element-child ordinals from its root. */
export interface XlsxSourcePart {
    part: string;
    partSha256: string;
    path: number[];
}
export type XlsxExportDiagnosticCode = 'hidden-content-excluded' | 'visibility-unknown' | 'unsupported-sheet' | 'unsupported-content' | 'object-placeholder' | 'unreadable-object' | 'provenance-unavailable' | 'comments-omitted' | 'rich-text-omitted' | 'formatting-approximate' | 'formula-cache-unverified' | 'formula-result-missing' | 'calculation-failure' | 'value-unavailable' | 'truncated' | 'markdown-lossy';
export type XlsxExportDiagnostic = ExportDiagnostic<XlsxExportDiagnosticCode, XlsxAnchor>;
/** A cell value; `unavailable` stands for one JSON cannot carry, such as a non-finite number. */
export type XlsxExportValue = XlsxCellValue | {
    kind: 'unavailable';
    reason: string;
};
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
    merge: {
        range: string;
        origin: boolean;
    } | null;
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
export type XlsxObjectKind = 'chart' | 'picture' | 'shape' | 'group' | 'connector' | 'diagram' | 'graphicFrame' | 'contentPart' | 'unknown';
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
export type XlsxSheetVisibility = 'visible' | 'hidden' | 'veryHidden' | 'unknown';
/**
 * One sheet. `truncated` means export stopped inside it: its lists are complete in the order
 * hidden spans, merges, tables, hyperlinks, objects, cells, up to where it stopped.
 */
export interface XlsxExportSheet extends ExportCompletion {
    id: string;
    anchor: XlsxAnchor;
    kind: 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet' | 'other';
    visibility: XlsxSheetVisibility;
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
    calculation: {
        policy: 'asStored';
        freshness: 'unverified';
    };
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
export type XlsxExportResult<T> = {
    ok: true;
    version: string;
    content: T;
} | OperationRefusal<XlsxExportFailure>;
export type { XlsxHeaderFooterText, XlsxHeaderFooterPair, XlsxSheetHeaderFooter } from './edits';

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

// ---- index.d.ts ----
/**
 * `@betteroffice/xlsx` — the framework-free core: display-list types, the Canvas2D
 * backend, pure viewport math, and the wasm loader. Framework chrome lives in
 * the adapter packages (`@betteroffice/xlsx-react`); nothing here touches react/vue
 * (lint-enforced).
 */
export * from './workerSession';
/** @experimental */
export { WorkbookPeerHydrationError } from './session/peerHydrationError';
/** @experimental */
export { WorkbookEditPeerFailedError } from './session/editPeer';
export type { Rect, TextAlign, FillRectCmd, PatternFillCmd, PatternType, LineCmd, GeometryPathCommand, PathStroke, PathCmd, TextCmd, TextWrap, DrawCmd, DisplayList, GridMeta, HyperlinkRegion, ChartRegion, ChartA11yAttrs, } from './display-list/types';
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
export { paintDisplayList } from './render/canvas2d';
export { initWasm, decodeFontFace, isWasmAvailable, isPngExportAvailable, isProposalsAvailable, openWorkbook, wasmVersion, engineBackend, StaleProposalError, DisplayTooLargeError, getDisplayListCellLimit, exportXlsxMarkdown, exportXlsxStructured, renderXlsxMarkdown, } from './wasm/loader';
export type { WasmInitInput, WasmInitOptions, EngineBackend, OpenWorkbookOptions, WorkbookCalculationContext, Viewport, PrintMetrics, SheetInfo, WorkbookHandle, XlsxTextMatch, XlsxTextSearchOptions, WorkbookUpdateInfo, WorkbookUpdateListener, WorkbookUpdateOrigin, CellEdit, CellPosition, CellInputEdit, EditResult, EditProfile, ProfiledEditResult, DisplayListProfile, ProfiledDisplayList, CalculationStatus, BorderPatch, BorderPreset, BorderStyle, CapturedFormat, CapturedColor, CapturedFill, CapturedBorderStyle, CapturedBorderEdge, CapturedCellFormat, CellPoint, ConditionalFormatInfo, XlsxAutoFilterOp, ConditionalFormatOp, ConditionalFormatRuleInput, DifferentialFormatInput, HistoryState, HorizontalAlignment, MergedRange, WorkbookHyperlink, NumberFormat, NumberFormatMutation, Proposal, ProposalCell, ProposalEdit, RangeStylePatch, SelectionFormatting, StaleProposalTarget, StyleProperty, TextWrapping, VerticalAlignment, } from './wasm/loader';
export type { EditSuccess, OperationFailure, OperationRefusal, ValidationSuccess, XlsxCellAddress, XlsxCellGuard, XlsxCellPosition, XlsxCellRead, XlsxCellValue, XlsxEditCalculation, XlsxEditFailure, XlsxEditFailureCode, XlsxEditHistory, XlsxEditOperation, XlsxEditPreview, XlsxEditReceipt, XlsxEditRefusal, XlsxEditRequest, XlsxEditResult, XlsxEditSource, XlsxEditStep, XlsxErrorValue, XlsxFindMatch, XlsxFindRequest, XlsxFindResult, XlsxRangeAddress, XlsxRangeRead, XlsxRangeTarget, XlsxReadCalculation, XlsxReadRequest, XlsxReadResult, XlsxSheetEntry, XlsxHeaderFooterText, XlsxHeaderFooterPair, XlsxSheetHeaderFooter, XlsxValidationResult, } from './edits';
export type { ExportCompletion, ExportDiagnostic, ExportSeverity, MarkdownAnchor, XlsxAnchor, XlsxExportCell, XlsxExportDefinedName, XlsxExportDiagnostic, XlsxExportDiagnosticCode, XlsxExportFailure, XlsxExportFailureCode, XlsxExportHyperlink, XlsxExportMerge, XlsxExportObject, XlsxExportOptions, XlsxExportResult, XlsxExportScope, XlsxExportSheet, XlsxExportTable, XlsxExportValue, XlsxFormulaResult, XlsxMarkdownContent, XlsxMarkdownOptions, XlsxObjectKind, XlsxSheetIdentity, XlsxSheetVisibility, XlsxSourcePart, XlsxStructuredContent, } from './exports';
export type { XlsxDefinedNameOp, XlsxDefinedName, XlsxDefinedNamesRequest, XlsxDefinedNamesResult } from './definedNames';
export type { ChartInfo, ChartSeriesInfo, ChartEdit, ChartMutation, ChartSeriesName } from './charts';

// ---- render/canvas2d.d.ts ----
/**
 * Canvas2D backend for the display list — the browser paint target.
 *
 * DOM canvas types are used deliberately here; this file is the browser backend
 * and is excluded from the pure seam (viewport/ + display-list/). The function
 * is otherwise pure: it reads a plain display list and issues draw calls, with
 * no allocation of app state and no reads back from the canvas.
 */
import type { DisplayList } from '../display-list/types';
/**
 * Paint a display list into a 2D context. `dpr` maps device-independent list
 * coordinates onto the backing store, so callers size the canvas at
 * `width * dpr` × `height * dpr`. `origin` is in backing-store pixels.
 */
export declare function paintDisplayList(ctx: CanvasRenderingContext2D, dl: DisplayList, dpr: number, origin?: {
    x: number;
    y: number;
}): void;

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

// ---- session/client.d.ts ----
import { type PipeCommandPayload } from '../../../../shared/collaboration-pipe/protocol';
import type { AwarenessCursor } from '../collaboration/types';
import { type SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type Promisified } from '../../../../shared/office-session/types';
import { type FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { type OpenWorkbookOptions, type Viewport, type WorkbookHandle } from '../wasm/loader';
import { type WorkbookFrame, type WorkbookFrameOptions, type WorkbookInitialPaint, type WorkbookSessionEvents, type WorkbookSessionMethods, type WorkbookSessionOpenOptions, type WorkbookSessionState } from './methods';
import type { CommandBasis, CommandIdentity, CommandStatus, WorkbookAck, WorkbookCommand, WorkbookCommandFrame, WorkbookCommandSelection, WorkbookEditState } from './commandProtocol';
import { type WorkbookEditSession } from './editSession';
import { type AsyncReplicaAdapter } from '../collaboration/asyncReplica';
/**
 * Options for opening a workbook in a dedicated worker.
 * @experimental
 */
export interface OpenWorkbookSessionOptions extends OpenWorkbookOptions {
    worker?: () => Worker;
    wasm?: ArrayBuffer | WebAssembly.Module;
    png?: boolean;
    /** Retain source bytes and a compiled module for edit peer hydration. */
    retainPeerHydration?: boolean;
    /** Aborting closes the session worker while the open is still in flight. */
    signal?: AbortSignal;
    initialViewport?: WorkbookSessionOpenOptions['initialViewport'];
    includeEditState?: boolean;
    initialUpdate?: Uint8Array;
}
/**
 * Async workbook access and its current local projection.
 * @experimental
 */
export interface WorkbookSession {
    readonly pngExport?: boolean;
    readonly replica?: AsyncReplicaAdapter;
    bindCollaboration?(apply: (update: Uint8Array) => Promise<unknown>, applyChunks?: (updates: Uint8Array[]) => Promise<unknown>, applyPipe?: (payload: PipeCommandPayload<AwarenessCursor>) => Promise<unknown>): () => void;
    readonly state: WorkbookSessionState;
    readonly initialPaint?: WorkbookInitialPaint;
    readonly commandIdentity: CommandIdentity;
    readonly confirmedBasis: CommandBasis;
    readonly throughSeq: number;
    readonly presentedThroughSeq: number;
    readonly editState: Readonly<Partial<WorkbookEditState>>;
    readonly selection: WorkbookCommandSelection | undefined;
    readonly editSession?: WorkbookEditSession;
    submit(command: WorkbookCommand): Promise<CommandStatus>;
    onAck(listener: (ack: WorkbookAck) => void): () => void;
    markPresented(frame: WorkbookCommandFrame | Pick<WorkbookCommandFrame, 'session' | 'incarnation' | 'coversThroughSeq'>): boolean;
    whenPresented(seq: number): Promise<void>;
    readonly call: Promisified<Omit<WorkbookSessionMethods, 'open' | 'openRecovery' | 'dispose' | 'frame' | 'command' | 'edit' | 'checkpoint' | 'commandFrame' | 'decodeFontFace' | 'renderPng' | 'renderRangePng' | 'printDisplayList'>> & {
        renderPng(viewport: Viewport, options?: WorkbookFrameOptions): Promise<Uint8Array>;
        renderRangePng(options?: {
            range?: string;
            scale?: number;
        }, frameOptions?: WorkbookFrameOptions): Promise<Uint8Array>;
        printDisplayList(...args: Parameters<WorkbookHandle['printDisplayList']>): Promise<ReturnType<WorkbookHandle['printDisplayList']>>;
        command(command: WorkbookCommand): Promise<WorkbookAck>;
        /**
         * @experimental A newer `frame` call replaces a queued one, which rejects with an error
         * named `SessionSuperseded`.
         */
        frame(viewport: Viewport, options?: WorkbookFrameOptions): Promise<WorkbookFrame>;
    };
    save(): Promise<Uint8Array>;
    on<K extends keyof WorkbookSessionEvents>(name: K, listener: (payload: WorkbookSessionEvents[K]) => void): () => void;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    readonly failure: SessionFailure | undefined;
    dispose(): Promise<void>;
}
export declare const workbookFontRuntimes: WeakMap<WorkbookSession, {
    decoder: FontDecoder;
}>;
/** @internal */
export declare function hydratePeer(session: WorkbookSession): Promise<WorkbookHandle>;
/**
 * Opens a worker session, transferring copies so caller buffers stay usable.
 * @experimental
 */
export declare function openWorkbookSession(bytes: Uint8Array | ArrayBuffer, options?: OpenWorkbookSessionOptions): Promise<WorkbookSession>;
export declare function createWorkbookSession(bytes: Uint8Array | ArrayBuffer, options: OpenWorkbookSessionOptions, transport: SessionTransport): Promise<WorkbookSession>;

// ---- session/clientRecovery.d.ts ----
import type { Checkpoint } from '../../../../shared/office-session/commands/types';
import { type SessionTransport } from '../../../../shared/office-session/transport';
import type { OpenWorkbookSessionOptions, WorkbookSession } from './client';
import type { CommandStatus, WorkbookAck, WorkbookCommand, WorkbookSharedCommand } from './commandProtocol';
import { type WorkbookCheckpoint } from './methods';
/** @experimental */
export declare class WorkbookRecoveryReplayError extends Error {
    readonly seq: number;
    readonly reason: string;
    readonly code = "workbook-recovery-refused";
    constructor(seq: number, reason: string);
}
/** @experimental */
export interface WorkbookSessionRecovery {
    readonly throughSeq: number;
    readonly retainedCommands: number;
    readonly checkpointIdleMs: number;
    capture(allowed: () => boolean): Promise<boolean>;
    save(assertCurrent: () => void, pending?: readonly Pick<WorkbookCommand, 'kind' | 'payload'>[]): Promise<Uint8Array>;
}
interface RecoverySession {
    replay(command: WorkbookSharedCommand, pending?: boolean): Promise<CommandStatus>;
    save(): Promise<Uint8Array>;
    dispose(): Promise<void>;
}
/** @experimental */
export declare const workbookSessionRecovery: WeakMap<WorkbookSession, WorkbookSessionRecovery>;
export declare const workbookRecoveryBackend: {
    open(bytes: Uint8Array, options: OpenWorkbookSessionOptions, checkpoint?: Checkpoint, version?: string): Promise<RecoverySession>;
};
export declare function createWorkbookRecoverySession(bytes: Uint8Array, options: OpenWorkbookSessionOptions, transport: SessionTransport, checkpoint?: Checkpoint, version?: string): Promise<RecoverySession>;
export declare function createWorkbookSessionRecovery(bytes: Uint8Array, options: OpenWorkbookSessionOptions, checkpoint: (allowed: () => boolean) => Promise<WorkbookCheckpoint | undefined>, epoch: number, liveCalculation?: boolean, version?: string, fail?: (error: unknown) => void): {
    recovery: WorkbookSessionRecovery;
    normalizePayload: (payload: WorkbookCommand["payload"]) => WorkbookCommand["payload"];
    admit(input: WorkbookSharedCommand): void;
    acknowledge(ack: WorkbookAck): void;
    dispose(): void;
};
export {};

// ---- session/commandProtocol.d.ts ----
import { type PipeCommandPayload } from '../../../../shared/collaboration-pipe/protocol';
import type { AwarenessCursor } from '../collaboration/types';
import type { XlsxRangeTarget } from '../edits';
import type { CommandAck as SharedCommandAck, CommandConnection, CommandEnvelope as SharedCommandEnvelope, CommandOutcome, CommandTicket, EchoHooks, FrameEnvelope, SubmitOptions, TicketOutcome } from '../../../../shared/office-session/commands/types';
import type { CalculationStatus, CellPoint, MergedRange, SelectionFormatting, Viewport, WorkbookHandle } from '../wasm/loader';
import type { WorkbookFrame, WorkbookSheetSummary, WorkbookWireFrame } from './methods';
import { type WorkbookReplayMethod, type WorkbookReplayOp } from './replay';
export interface CommandIdentity {
    session: string;
    incarnation: string;
}
export interface CommandBasis {
    epoch: number;
    version: string;
}
export interface CommandEnvelope<K extends string, P> extends CommandIdentity {
    seq: number;
    basis: CommandBasis;
    afterSeq?: number;
    kind: K;
    payload: P;
}
export interface CommandStatus<R = unknown> {
    seq: number;
    status: 'applied' | 'noop' | 'refused' | 'cancelled';
    reason?: string;
    committed?: boolean;
    affectedSeqs?: number[];
    result?: R;
}
export interface CommandAck<S, E, F> {
    throughSeq: number;
    version: string;
    selection?: S;
    editState?: Partial<E>;
    statuses: CommandStatus[];
    frame?: F;
}
export interface WorkbookCommandSelection extends XlsxRangeTarget {
    activeCell?: CellPoint;
}
export interface WorkbookCommandContext extends CommandIdentity {
    basis: CommandBasis;
}
interface WorkbookIntent {
    inputTime?: number;
    undoGroupId?: string;
    selection?: WorkbookCommandSelection;
}
type CommandArgs<K extends WorkbookReplayMethod> = K extends 'propose' ? [string, string | null, (Parameters<WorkbookHandle['propose']>[2][number] & {
    target?: XlsxRangeTarget;
})[]] : Extract<WorkbookReplayOp, {
    method: K;
}>['args'];
type ReplayPayload<K extends WorkbookReplayMethod> = Extract<WorkbookReplayOp, {
    method: K;
}> extends infer O ? O extends WorkbookReplayOp ? Pick<O, 'calculation'> & {
    args: CommandArgs<K>;
} & WorkbookIntent & {
    target?: XlsxRangeTarget;
} & (K extends 'setActiveSheet' ? {
    expectVersion?: string;
} : {}) : never : never;
export type WorkbookCommandKind = WorkbookReplayMethod | 'select' | 'setReadOnly' | 'applyUpdate' | 'collabPipe';
export type WorkbookToolbarPayload = {
    id: 'bold' | 'italic' | 'strikethrough' | 'paintFormat' | 'clear' | 'fontFamily' | 'fontSize' | 'fontSizeStep' | 'decimalPlaces' | 'textColor' | 'fillColor' | 'borderColor' | 'numberFormat' | 'borderPreset' | 'borderStyle' | 'merge' | 'horizontalAlignment' | 'verticalAlignment' | 'textWrapping';
    args: unknown;
    target: XlsxRangeTarget;
} | {
    id: 'refreshProposals';
    args: unknown;
    target?: never;
};
export interface WorkbookFormatTokenResult {
    formatToken: string;
}
export type WorkbookCommandPayload<K extends WorkbookCommandKind> = K extends WorkbookReplayMethod ? ReplayPayload<K> | (K extends 'applyOps' ? WorkbookIntent & {
    toolbar: WorkbookToolbarPayload;
} : never) | (K extends 'applyFormat' ? WorkbookIntent & {
    target: XlsxRangeTarget;
    formatToken: string;
} | WorkbookIntent & {
    target: XlsxRangeTarget;
    capturedFrom: XlsxRangeTarget;
} : never) : K extends 'collabPipe' ? PipeCommandPayload<AwarenessCursor> : K extends 'select' ? {
    selection: WorkbookCommandSelection;
} : K extends 'applyUpdate' ? {
    update: Uint8Array;
    updates?: Uint8Array[];
} : {
    readOnly: boolean;
};
export interface WorkbookCommandFrameRequest {
    sheet: number;
    viewport: Viewport;
    zoom: number;
    dpr: number;
    generation: number;
    requestGeneration?: number;
}
export type WorkbookCommand = {
    [K in WorkbookCommandKind]: CommandEnvelope<K, WorkbookCommandPayload<K>> & {
        frame?: WorkbookCommandFrameRequest;
    };
}[WorkbookCommandKind];
export interface WorkbookEditState {
    historyReset?: boolean;
    canUndo: boolean;
    canRedo: boolean;
    formatting: SelectionFormatting;
    formattingAvailable: boolean;
    selectionEditable: boolean;
    mergedRanges: MergedRange[];
    calculation: CalculationStatus;
    sheets: WorkbookSheetSummary[];
    activeSheet: number;
    readOnly: boolean;
}
interface CommandFrameTags extends CommandIdentity {
    coversThroughSeq: number;
    requestGeneration: number;
    zoom: number;
    dpr: number;
}
export type WorkbookCommandFrame = WorkbookFrame & CommandFrameTags;
export type WorkbookCommandWireFrame = WorkbookWireFrame & CommandFrameTags;
export type WorkbookAck = CommandAck<WorkbookCommandSelection, WorkbookEditState, WorkbookCommandFrame>;
export type WorkbookWireAck = CommandAck<WorkbookCommandSelection, WorkbookEditState, WorkbookCommandWireFrame> & {
    shared?: WorkbookSharedWireAck;
};
export interface WorkbookEditPayload {
    payload: WorkbookCommand['payload'];
    frame?: WorkbookCommandFrameRequest;
    compatibility?: WorkbookCommand;
    dispatchOrder?: number;
}
export type WorkbookSharedCommand = SharedCommandEnvelope<WorkbookEditPayload>;
export type WorkbookSharedOutcome = CommandOutcome & {
    workbook?: CommandStatus;
};
export type WorkbookSharedAck<F = WorkbookCommandFrame> = Omit<SharedCommandAck<WorkbookCommandSelection, Partial<WorkbookEditState>, F>, 'outcomes'> & {
    engineVersion: string;
    outcomes: WorkbookSharedOutcome[];
    compatibility?: CommandAck<WorkbookCommandSelection, WorkbookEditState, F>;
};
export type WorkbookSharedWireAck = WorkbookSharedAck<WorkbookCommandWireFrame>;
export type WorkbookEditConnection = Omit<CommandConnection<WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame>, 'initial' | 'client'> & {
    initial: WorkbookSharedAck;
    client: Omit<CommandConnection<WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame>['client'], 'call'> & {
        call: {
            edit(command: WorkbookSharedCommand): Promise<WorkbookSharedAck>;
            checkpoint(): Promise<import('../../../../shared/office-session/commands/types').Checkpoint>;
        };
    };
};
export type WorkbookEditFrame = FrameEnvelope<WorkbookCommandFrame>;
export type WorkbookTicketOutcome = TicketOutcome & {
    workbook?: CommandStatus;
};
export type WorkbookCommandTicket = Omit<CommandTicket, 'ack'> & {
    readonly ack: Promise<WorkbookTicketOutcome>;
};
export type WorkbookSubmitOptions<X = unknown> = SubmitOptions<X>;
export type WorkbookEchoHooks<X = unknown> = EchoHooks<X, WorkbookCommandFrame>;
export declare function isWorkbookCommandFrameRequest(value: unknown): value is WorkbookCommandFrameRequest;
export declare function validateWorkbookCommandEnvelope(value: unknown): asserts value is WorkbookCommand;
export declare function validateWorkbookCommandPayload(command: WorkbookCommand, pipeFrameBytes?: ReadonlyMap<number, number>): void;
export declare class WorkbookCommandRefusal extends Error {
    readonly reason: string;
    readonly result?: unknown | undefined;
    constructor(reason: string, result?: unknown | undefined);
}
export {};

// ---- session/editPeer.d.ts ----
import { type WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';
import { type WorkbookReplayMethod } from './replay';
export interface WorkbookEditPeerOptions {
    session: WorkbookSession;
    peer: WorkbookHandle;
    now?(): number;
    randomSeed?(): number;
    onError?(error: Error): void;
}
export type WorkbookEditPeer = Pick<WorkbookHandle, WorkbookReplayMethod> & {
    readonly state: 'ready' | 'failed';
    readonly error: Error | undefined;
    readonly acknowledgedSequence: number;
    readonly sentSequence: number;
    flush(): Promise<void>;
    save(): Promise<ArrayBuffer>;
    recoverySave(): {
        bytes: ArrayBuffer;
        recovery: true;
    };
    dispose(): void;
};
export declare class WorkbookEditPeerFailedError extends Error {
    readonly failure: Error;
    constructor(failure: Error);
}
/** @internal */
export declare function createWorkbookEditPeer(options: WorkbookEditPeerOptions): WorkbookEditPeer;

// ---- session/editSession.d.ts ----
import type { EchoHooks, EditSession, FrameEnvelope, RecoveryHooks, WorkerEditSessionOptions } from '../../../../shared/office-session/commands/types';
import { SessionFailure } from '../../../../shared/office-session/types';
import { type CommandIdentity, type WorkbookAck, type WorkbookCommand, type WorkbookCommandFrame, type WorkbookCommandSelection, type WorkbookEditConnection, type WorkbookEditPayload, type WorkbookEditState, type WorkbookSharedCommand, type WorkbookCommandTicket, type WorkbookSubmitOptions } from './commandProtocol';
type SessionOptions<X> = WorkerEditSessionOptions<unknown, WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame, X>;
export interface WorkbookEditSession<X = unknown> extends EditSession<WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, X> {
    readonly engineVersion: string;
    readonly failure: SessionFailure | undefined;
    readonly presentationRevision: number | undefined;
    submit(kind: string, payload: WorkbookEditPayload, options?: WorkbookSubmitOptions<X>): WorkbookCommandTicket;
    command(input: WorkbookCommand): Promise<WorkbookAck>;
    runAfterEdits<T>(call: () => Promise<T>): Promise<T>;
    onAck(listener: (ack: WorkbookAck) => void): () => void;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    bindPresenter(echo: EchoHooks<X, WorkbookCommandFrame>): () => void;
    present(frame: FrameEnvelope<WorkbookCommandFrame>): void;
    markPresented(frame: Pick<WorkbookCommandFrame, 'session' | 'incarnation' | 'coversThroughSeq'> & Partial<WorkbookCommandFrame>): boolean;
}
export declare function createWorkbookEditSession<X = unknown>(options: Omit<SessionOptions<X>, 'connection' | 'source' | 'recovery' | 'maxRecoveryAttempts' | 'transactionalPresentation' | 'echoOverflow' | 'sameCoverageFrames'> & {
    connection: WorkbookEditConnection;
    identity?: CommandIdentity;
    source?: unknown;
    recovery?: RecoveryHooks<unknown, WorkbookEditPayload, WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame>;
    onAdmit?(command: WorkbookSharedCommand): void;
    normalizePayload?(payload: WorkbookCommand['payload']): WorkbookCommand['payload'];
}): WorkbookEditSession<X>;
export {};

// ---- session/methods.d.ts ----
import type { FontDecoderMethods } from '../../../../shared/office-session/fontDecoder';
import type { XlsxDefinedNamesRequest, XlsxDefinedNamesResult } from '../definedNames';
import type { PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { MethodPolicies, SessionState } from '../../../../shared/office-session';
import type { Direction, SelectionData } from '../selection/types';
import type { DisplayList, Rect } from '../display-list/types';
import type { AwarenessPeer, CollaborationUpdateOrigin } from '../collaboration/types';
import type { XlsxExportOptions, XlsxExportResult, XlsxMarkdownContent, XlsxMarkdownOptions, XlsxStructuredContent } from '../exports';
import type { WorkbookCommand, WorkbookCommandContext, WorkbookCommandFrame, WorkbookCommandSelection, WorkbookCommandWireFrame, WorkbookEditState, WorkbookWireAck, WorkbookSharedCommand, WorkbookSharedWireAck, WorkbookCommandFrameRequest } from './commandProtocol';
import type { Checkpoint, FrameEnvelope } from '../../../../shared/office-session/commands/types';
import type { XlsxEditRequest, XlsxEditResult, XlsxFindRequest, XlsxFindResult, XlsxReadRequest, XlsxReadResult, XlsxValidationResult } from '../edits';
import type { CalculationStatus, CapturedFormat, CellEdit, CellPosition, MergedRange, OpenWorkbookOptions, PrintMetrics, SheetInfo, Viewport } from '../wasm/loader';
/** @internal */
export interface WorkbookCollaborationPipeReply {
    collabPipe?: PipeOutputBatch<AwarenessPeer>[];
}
/** @experimental */
export declare class WorkbookPngUnavailableError extends Error {
    readonly code = "png-unavailable";
    constructor();
}
export type WorkbookCheckpoint = Checkpoint & {
    encodeMs?: number;
};
/** @experimental Unzoomed sheet pixels. */
export type WorkbookSheetView = Pick<SheetInfo, 'contentWidth' | 'contentHeight' | 'frozenRows' | 'frozenCols' | 'initialScrollX' | 'initialScrollY'> & {
    sheet: number;
    version: string;
    frozenWidth: number;
    frozenHeight: number;
};
/** @experimental Absolute sheet rect and freeze-relative scroll position, in unzoomed pixels. */
export interface WorkbookCellGeometry {
    sheet: number;
    version: string;
    rect: Rect;
    scrollPosition: CellPosition;
}
/** @experimental Row-major editable cell text. */
export interface WorkbookCellInputs {
    sheet: number;
    version: string;
    cells: CellEdit[][];
}
/** @experimental */
export interface WorkbookFrameOptions {
    sheet?: number;
}
/** @experimental */
export interface WorkbookFrame {
    displayList: DisplayList;
    version: string;
    epoch: number;
    sequence: number;
    sheet: number;
    viewport: Viewport;
    mergedRanges?: MergedRange[];
}
/** @experimental */
export type WorkbookWireFrame = Omit<WorkbookFrame, 'displayList'> & {
    displayList: ArrayBuffer;
};
/**
 * Structured-cloneable workbook open options.
 * @experimental
 */
export interface WorkbookSessionOpenOptions extends OpenWorkbookOptions {
    wasm?: ArrayBuffer | WebAssembly.Module;
    png?: boolean;
    initialViewport?: WorkbookInitialPaintRequest;
    includeEditState?: boolean;
    initialUpdate?: Uint8Array;
    /** @internal */
    recoveryVersion?: string;
}
/** @experimental */
export interface WorkbookInitialPaintRequest {
    sheet?: number;
    width: number;
    height: number;
    zoom: number;
    dpr: number;
    requestGeneration: number;
}
/** @experimental */
export interface WorkbookInitialPaint<F = WorkbookCommandFrame> {
    sheetView: WorkbookSheetView;
    frame: F;
    editState: WorkbookEditState;
}
/** @experimental */
export type WorkbookSessionOpenReply = WorkbookSessionState & {
    pngExport?: boolean;
    clientId?: number;
    throughSeq?: number;
    revision?: number;
    historyReset?: boolean;
    initialPaint?: WorkbookInitialPaint<WorkbookCommandWireFrame>;
    editState?: WorkbookEditState;
    selection?: WorkbookCommandSelection;
};
/**
 * Sheet metadata; indices are zero-based.
 * @experimental
 */
export interface WorkbookSheetSummary {
    id: string;
    index: number;
    name: string;
}
/**
 * Workbook projection; version counts applied batches.
 * @experimental
 */
export type WorkbookSessionState = Omit<SessionState, 'stage'> & {
    format: 'xlsx';
    stage: 'preview' | 'ready' | 'failed';
    sheets: WorkbookSheetSummary[];
    activeSheet: number;
    commandContext?: WorkbookCommandContext;
};
/**
 * Workbook change notifications after applied batches.
 * @experimental
 */
export interface WorkbookSessionEvents {
    changed: {
        version: number;
        dirty: boolean;
    };
    collaborationUpdate: {
        update: Uint8Array;
        origin: CollaborationUpdateOrigin;
    };
}
/**
 * Workbook RPC methods; engine version tokens remain strings.
 * @experimental
 */
export type WorkbookSessionMethods = FontDecoderMethods & {
    open(bytes: ArrayBuffer, options?: WorkbookSessionOpenOptions): WorkbookSessionOpenReply;
    openRecovery(checkpoint: Checkpoint, options?: Pick<WorkbookSessionOpenOptions, 'wasm' | 'png'>): WorkbookSessionOpenReply;
    version(): string;
    collaborationStateVector(): Uint8Array;
    collaborationStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    listDefinedNames(request?: XlsxDefinedNamesRequest): XlsxDefinedNamesResult;
    readCells(request: XlsxReadRequest): XlsxReadResult;
    findText(request: XlsxFindRequest): XlsxFindResult;
    validateEdits(request: XlsxEditRequest): XlsxValidationResult;
    applyEdits(request: XlsxEditRequest): XlsxEditResult;
    command(command: WorkbookCommand): WorkbookWireAck;
    edit(command: WorkbookSharedCommand): WorkbookSharedWireAck;
    checkpoint(preserveHistory?: boolean): WorkbookCheckpoint;
    commandFrame(request?: WorkbookCommandFrameRequest): FrameEnvelope<WorkbookCommandWireFrame> | undefined;
    frame(viewport: Viewport, options?: WorkbookFrameOptions): WorkbookWireFrame;
    sheetView(sheet: number): WorkbookSheetView;
    cellGeometry(sheet: number, row: number, col: number): WorkbookCellGeometry;
    cellInputs(sheet: number, range: string): WorkbookCellInputs;
    selectionData(sheet: number, row: number, col: number, direction?: Direction | 'usedRange' | 'rowEnd' | 'merge'): SelectionData;
    captureFormat(sheet: number, range: string): CapturedFormat;
    sheets(): WorkbookSheetSummary[];
    calculationStatus(): CalculationStatus;
    renderPng(viewport: Viewport, options?: WorkbookFrameOptions): ArrayBuffer;
    renderRangePng(options?: {
        range?: string;
        scale?: number;
    }, frameOptions?: WorkbookFrameOptions): ArrayBuffer;
    exportStructured(options?: XlsxExportOptions): XlsxExportResult<XlsxStructuredContent>;
    exportMarkdown(options?: XlsxExportOptions, markdownOptions?: XlsxMarkdownOptions): XlsxExportResult<XlsxMarkdownContent>;
    printDisplayList(sheet: number, range: string, metrics: PrintMetrics, gridlines: boolean): ArrayBuffer;
    save(): ArrayBuffer;
    dispose(): void;
};
export declare const WORKBOOK_SESSION_METHODS: {
    decodeFontFace: true;
    open: true;
    openRecovery: true;
    version: true;
    collaborationStateVector: true;
    collaborationStateAsUpdate: true;
    listDefinedNames: true;
    readCells: true;
    findText: true;
    validateEdits: true;
    applyEdits: true;
    command: true;
    edit: true;
    checkpoint: true;
    commandFrame: true;
    frame: true;
    sheetView: true;
    cellGeometry: true;
    cellInputs: true;
    selectionData: true;
    captureFormat: true;
    sheets: true;
    calculationStatus: true;
    renderPng: true;
    renderRangePng: true;
    exportStructured: true;
    exportMarkdown: true;
    printDisplayList: true;
    save: true;
    dispose: true;
};
export declare const WORKBOOK_SESSION_POLICIES: MethodPolicies<WorkbookSessionMethods>;

// ---- session/peerHydrationError.d.ts ----
export type WorkbookPeerHydrationErrorCode = 'mutation-outside-replay' | 'version-mismatch' | 'missing-hydration' | 'missing-module';
export declare class WorkbookPeerHydrationError extends Error {
    readonly code: WorkbookPeerHydrationErrorCode;
    readonly refusal: {
        code: WorkbookPeerHydrationErrorCode;
    };
    constructor(code: WorkbookPeerHydrationErrorCode, message: string);
}

// ---- session/replay.d.ts ----
import type { MethodPolicies } from '../../../../shared/office-session';
import type { StaleProposalTarget, Viewport, WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';
import type { WorkbookFrame, WorkbookWireFrame } from './methods';
export declare const WORKBOOK_REPLAY_MUTATORS: {
    editCell: true;
    editCells: true;
    applyEdits: true;
    applyOps: true;
    propose: true;
    acceptProposal: true;
    rejectProposal: true;
    patchRangeStyle: true;
    setNumberFormat: true;
    applyFormat: true;
    moveChart: true;
    moveRange: true;
    undo: true;
    redo: true;
    setActiveSheet: true;
};
export type WorkbookReplayMethod = keyof typeof WORKBOOK_REPLAY_MUTATORS;
export type WorkbookReplayOp = {
    [K in WorkbookReplayMethod]: {
        method: K;
        args: Parameters<WorkbookHandle[K]>;
    };
}[WorkbookReplayMethod] & {
    calculation?: WorkbookCalculationContext;
};
export interface WorkbookReplayEnvelope {
    sequence: number;
    calculation: WorkbookCalculationContext;
    op: WorkbookReplayOp;
    staleProposal?: {
        cells: string[];
        targets: StaleProposalTarget[];
    };
}
export interface WorkbookReplayReply {
    sequence: number;
    revision: number;
    version: number;
    result: ReturnType<WorkbookHandle[WorkbookReplayMethod]>;
}
export type WorkbookInternalSessionMethods = {
    beginPeerSnapshot(records: number, bytes: number): {
        version: string;
        sequence: number;
    };
    pullPeerSnapshot(): ArrayBuffer[] | undefined;
    endPeerSnapshot(discard: boolean): void;
    attachPeer(version: string, sequence: number): void;
    detachPeer(): void;
    replay(envelope: WorkbookReplayEnvelope): WorkbookReplayReply;
    preview(viewport: Viewport, sheet: number, ops: readonly WorkbookReplayOp[]): WorkbookWireFrame;
};
export declare const WORKBOOK_INTERNAL_SESSION_METHODS: {
    readonly beginPeerSnapshot: true;
    readonly pullPeerSnapshot: true;
    readonly endPeerSnapshot: true;
    readonly attachPeer: true;
    readonly detachPeer: true;
    readonly replay: true;
    readonly preview: true;
};
export declare const WORKBOOK_INTERNAL_SESSION_POLICIES: MethodPolicies<WorkbookInternalSessionMethods>;
export declare const workbookSessionInternals: WeakMap<WorkbookSession, {
    replay(envelope: WorkbookReplayEnvelope): Promise<WorkbookReplayReply>;
    initialVersion?: string;
    attachPeer?(version: string): Promise<void>;
    detachPeer?(): Promise<void>;
    editPeerAttached: boolean;
    cellInput?(sheet: number, row: number, col: number): Promise<string>;
    initialCalculation?: WorkbookCalculationContext | null;
    preview?(viewport: Viewport, sheet: number, ops: readonly WorkbookReplayOp[]): Promise<WorkbookFrame>;
}>;
export declare function validateWorkbookReplayEnvelope(value: unknown): asserts value is WorkbookReplayEnvelope;
export declare function workbookReplayRefused(result: unknown): boolean;
export declare function applyWorkbookReplayOp(handle: WorkbookHandle, op: WorkbookReplayOp): WorkbookReplayReply['result'];

// ---- shared/collaboration-pipe/main.d.ts ----
import { type PipeCursor, type PipeErrorCode, type PipeOp, type PipeOutput, type PipePeer, type PipeProviderOptions } from './protocol';
type Status = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
type TransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
} | {
    type: 'drain';
};
export interface PipeHostTransport {
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: TransportEvent) => void): () => void;
}
/** One provider's line to the worker that runs it. */
export interface CollaborationPipeEndpoint<C = PipeCursor> {
    /** Runs `ops` in the worker, in order after earlier posts; settles once the worker has run them. */
    post(ops: PipeOp<C>[], transfer: ArrayBuffer[]): Promise<unknown>;
    close(): void;
}
/** What a worker replica offers a provider that can run in its worker. */
export interface CollaborationPipeChannel<C = PipeCursor, P = PipePeer> {
    /** `receive` takes the worker's output for this provider; `lost` reports the worker is gone for it. */
    open(receive: (out: readonly PipeOutput<P>[]) => void, lost: (reason: string) => void): CollaborationPipeEndpoint<C>;
}
export interface PipedCollaborationOptions<E, P = PipePeer> {
    worker: PipeProviderOptions;
    maxPendingBytes: number;
    hasUser: boolean;
    /** Moves a worker peer's timestamps onto this context's clock; omitted when they are wall-clock times. */
    rebasePeer?(peer: P, offset: number): P;
    makeError(code: PipeErrorCode, message: string, cause?: unknown): E;
    /** `cause` itself when it already is an `E`, else a new `E` with `context` and the cause's message. */
    normalize(code: PipeErrorCode, context: string, cause: unknown): E;
}
/**
 * The main-thread side of a provider whose protocol runs in the document worker. It owns the host transport's
 * lifecycle as a provider does, forwards the transport's events with frame bytes untouched, one post per burst,
 * and sends what the worker answers. Worker output older than main's latest lifecycle call is dropped.
 */
export declare class PipedCollaboration<E extends Error, C = PipeCursor, P = PipePeer> {
    private readonly transport;
    private readonly options;
    private readonly endpoint;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly peerListeners;
    private nextListenerId;
    private connectionStatus;
    private isSynced;
    private peerList;
    private seq;
    /** The op of main's latest lifecycle call; older worker output describes a connection main has left. */
    private floor;
    /** The op of the transport's latest open or close; sends, status and peers from before it describe a connection the host ended. */
    private sendFloor;
    /** Received frame bytes queued or in flight to the worker. */
    private inboundBytes;
    private flushingFrames;
    private queue;
    private posting;
    private flushQueued;
    private heldTimer;
    /** The queue holds an op that should not wait for held updates. */
    private urgent;
    private unsubscribeTransport;
    private subscription;
    private wantsConnection;
    private connectAttempt;
    private leaveFrame;
    private readonly pendingFrames;
    private queuedBytes;
    private isDestroyed;
    private closed;
    constructor(channel: CollaborationPipeChannel<C, P>, transport: PipeHostTransport, options: PipedCollaborationOptions<E, P>);
    get status(): Status;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly P[];
    setCursor(cursor: C | null, broadcast: boolean): void;
    onPeers(listener: (peers: readonly P[]) => void): () => void;
    onStatus(listener: (change: {
        status: Status;
        synced: boolean;
    }) => void): () => void;
    onError(listener: (error: E) => void): () => void;
    connect(): void;
    disconnect(): void;
    destroy(): void;
    private transportError;
    /** Leaves the connection main holds: drops the subscription, disconnects, clears queued bytes and peers. */
    private release;
    /** Fails main's connection with `error`, and the worker's with it. */
    private fail;
    private sendLeave;
    private enqueue;
    private schedule;
    private flush;
    private forward;
    private receive;
    private send;
    private queueFrame;
    private flushPending;
    private clearQueue;
    /**
     * The worker no longer runs this provider. Hands the host transport back at once; reports the loss a task later,
     * so a host that destroys the provider on `onReplica(null)` hears nothing.
     */
    private lose;
    private close;
    private clearPeers;
    private setStatus;
    private emitPeers;
    private report;
}
export {};

// ---- shared/collaboration-pipe/protocol.d.ts ----
/**
 * Wire format between a main-thread collaboration facade and the provider its document worker runs.
 * Main owns the host transport's lifecycle and forwards its bytes untouched; the worker runs the protocol and
 * answers with frames to send, state and the transport calls its own failures need.
 */
export type PipeErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export interface PipeError {
    code: PipeErrorCode;
    message: string;
}
export interface PipeUser {
    name: string;
    color?: string;
}
export interface PipeProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    syncChunkBytes?: number;
    maxSyncBytes?: number;
    user?: PipeUser;
}
/** The docx cursor; the default cursor type `C` of the pipe types. */
export interface PipeCursor {
    story: string;
    anchor: Uint8Array;
    head: Uint8Array;
}
export type PipeTransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: ArrayBuffer;
} | {
    type: 'close';
    reason?: string;
};
/**
 * Main to worker, in order; `seq` increases. `connect`, `disconnect`, `destroy` and `fail` replay what main already
 * did to the host transport, so the worker sends nothing and asks for no transport call while it runs them.
 */
export type PipeOp<C = PipeCursor> = {
    op: 'create';
    seq: number;
    options: PipeProviderOptions;
} | {
    op: 'connect';
    seq: number;
} | {
    op: 'disconnect';
    seq: number;
} | {
    op: 'destroy';
    seq: number;
} | {
    op: 'fail';
    seq: number;
    error: PipeError;
} | {
    op: 'cursor';
    seq: number;
    cursor: C | null;
    broadcast: boolean;
} | {
    op: 'event';
    seq: number;
    event: PipeTransportEvent;
};
/** The docx peer; the default peer type `P` of the pipe types. */
export interface PipePeer {
    clientId: number;
    clock: number;
    user: {
        name: string;
        color: string;
    };
    cursor: PipeCursor | null;
    lastSeenAt: number;
    cursorMovedAt: number;
    inferredCursor: {
        clientId: number;
        story: string;
        paraId: string;
        endOffset: number;
    } | null;
}
/**
 * Worker to main, in order. `through` is the last op the worker had run; main drops output older than its own
 * latest lifecycle call.
 */
export type PipeOutput<P = PipePeer> = {
    do: 'unsubscribe';
    through: number;
} | {
    do: 'disconnect';
    through: number;
} | {
    do: 'send';
    data: ArrayBuffer;
    through: number;
}
/** The awareness leave main sends when the host disconnects first; null while there is none to send. */
 | {
    do: 'leave';
    data: ArrayBuffer | null;
    through: number;
} | {
    do: 'status';
    status: 'disconnected' | 'connecting' | 'connected' | 'destroyed';
    synced: boolean;
    through: number;
} | {
    do: 'peers';
    peers: P[];
    timeOrigin: number;
    through: number;
} | {
    do: 'error';
    error: PipeError;
    through: number;
}
/** The worker has no provider under this id. */
 | {
    do: 'lost';
};
export interface PipeOutputBatch<P = PipePeer> {
    provider: number;
    out: PipeOutput<P>[];
}
export interface PipeCommandPayload<C = PipeCursor> {
    provider: number;
    ops: PipeOp<C>[];
}
export interface PipePayloadShape {
    /** Checks a non-null cursor; its byte fields stay within `maxFrameBytes`. */
    validCursor(cursor: unknown, maxFrameBytes: number): boolean;
    /** The keys a `create` op's options may carry. */
    optionKeys: ReadonlySet<string>;
}
/** Buffers a pipe payload holds, for transfer. */
export declare function pipeOpBuffers<C>(ops: readonly PipeOp<C>[]): ArrayBuffer[];
export declare function ownedBuffer(bytes: Uint8Array): ArrayBuffer;
export declare const PIPE_OPTION_KEYS: ReadonlySet<string>;
/** The docx cursor check. */
export declare function validPipeCursor(cursor: unknown, maxFrameBytes: number): boolean;
/** Shape check for a pipe command payload; frames past `maxFrameBytes` fail it. */
export declare function validPipePayload<C = PipeCursor>(payload: unknown, maxFrameBytes: number, shape: PipePayloadShape): payload is PipeCommandPayload<C>;
/** Epoch time of this context's provider clock zero: performance.now() when present, else Date.now(). */
export declare function clockOrigin(): number;

// ---- shared/host-contracts/edits.d.ts ----
/**
 * Result envelopes shared by the formats' host edit APIs. Targets, steps, receipts and failure
 * codes stay format-owned; these shapes only fix how every format reports success and refusal.
 */
/** Why a host operation was refused, as data. The document is unchanged. */
export interface OperationFailure<Code extends string, Target> {
    code: Code;
    /** The request step that failed, when one did. */
    stepIndex?: number;
    /** The earlier step an `overlapping-steps` failure conflicts with. */
    conflictingStepIndex?: number;
    /** The failing step's target as requested or resolved. */
    target?: Target;
    message: string;
}
/** A policy refusal carrying the version the document is still at. */
export interface OperationRefusal<Failure> {
    ok: false;
    version: string;
    failure: Failure;
}
/**
 * An accepted edit request. `applied` is false when every step was a no-op; the version then
 * stays `baseVersion` and no history or change notification was produced.
 */
export interface EditSuccess<Receipt> {
    ok: true;
    baseVersion: string;
    version: string;
    applied: boolean;
    /** One receipt per request step, in request order. */
    receipts: Receipt[];
}
/** A validated edit request. Validation reserves nothing and cannot be committed later. */
export interface ValidationSuccess<Preview> {
    ok: true;
    baseVersion: string;
    wouldApply: boolean;
    previews: Preview[];
}

// ---- shared/host-contracts/exports.d.ts ----
/**
 * Shapes shared by the formats' read-only structured exports. Content trees, anchors, story or
 * sheet selectors, diagnostic codes and limits stay format-owned; these only fix how every format
 * reports diagnostics, truncation and Markdown markers.
 */
export type ExportSeverity = 'info' | 'warning' | 'error';
/** Something an export omitted or could not represent, anchored where it applies. */
export interface ExportDiagnostic<Code extends string, Anchor> {
    code: Code;
    severity: ExportSeverity;
    anchor: Anchor | null;
    message: string;
}
/** Whether an export stopped at its limits before the end of the content. */
export interface ExportCompletion {
    truncated: boolean;
}
/** A Markdown marker comment and the anchor of the content it precedes. */
export interface MarkdownAnchor<Anchor> {
    marker: string;
    anchor: Anchor;
}

// ---- shared/office-session/client.d.ts ----
import { type HostMessage } from './protocol';
import type { SessionTransport } from './transport';
import { SessionFailure, type Promisified, type SessionEvents, type SessionMethods } from './types';
export declare const OFFICE_SESSION_SILENCE_MS = 60000;
export declare function requestWasmCompile(transport: SessionTransport): void;
export interface SessionClientOptions<M extends SessionMethods> {
    methods: {
        readonly [K in keyof M]-?: true;
    };
    silenceMs?: number;
    now?(): number;
    timer?(callback: () => void, ms: number): () => void;
    onWasmModule?(url: string, module: WebAssembly.Module): void;
    onWasmModuleMessage?(message: Extract<HostMessage, {
        kind: 'wasm-module';
    }>): void;
}
export interface SessionClient<M extends SessionMethods, E extends SessionEvents> {
    readonly call: Promisified<M>;
    callWithTransfer<K extends keyof M & string>(method: K, args: Parameters<M[K]>, transfer: Transferable[]): Promise<Awaited<ReturnType<M[K]>>>;
    on<K extends keyof E & string>(name: K, listener: (payload: E[K]) => void): () => void;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    readonly failure: SessionFailure | undefined;
    dispose(): Promise<void>;
}
/** Creates a typed client over an existing transport. */
export declare function createSessionClient<M extends SessionMethods, E extends SessionEvents>(transport: SessionTransport, options: SessionClientOptions<M>): SessionClient<M, E>;

// ---- shared/office-session/commands/spec.d.ts ----
import { type TransferResult } from '../protocol';
import type { CommandTicket, EditSession, SubmitOptions } from './types';
/** @internal */
export type BaseRefusal = 'read-only' | 'stale' | 'superseded' | 'not-ready' | 'busy' | 'disposed' | 'invalid-payload' | 'unsupported' | 'dependency-refused' | 'unknown-outcome';
/** @internal */
export declare const BASE_REFUSALS: readonly BaseRefusal[];
/** @internal */
export type Refusal<F extends string = never> = {
    code: BaseRefusal | F;
    detail?: string;
};
/** @internal */
export type Dependency = 'doc' | 'selection' | 'layout';
/** @internal Worker state at the read boundary. */
export type At = {
    incarnation: number;
    epoch: number;
    throughSeq: number;
    doc: number;
    selection: number;
    layout: number;
};
/** @internal */
export type Read<R, F extends string = never> = {
    ok: true;
    value: R;
    at: At;
} | {
    ok: false;
    refusal: Refusal<F>;
};
/** @internal */
export type Admission = 'view' | 'local-write' | 'remote-update' | 'mode';
declare const payloadType: unique symbol;
declare const queryType: unique symbol;
declare const REFUSAL: unique symbol;
/** @internal Use `satisfies Record<Kind, CommandSpec>` for exhaustive command kinds. */
export interface CommandSpec<P = unknown, F extends string = string> {
    readonly admission: Admission;
    readonly refusals: readonly F[];
    readonly validate?: (payload: unknown) => payload is P;
    readonly [payloadType]?: P;
}
/** @internal */
export type StaleMode = 'answer-at-fence' | {
    readonly reject: readonly Dependency[];
};
/** @internal */
export interface QuerySpec<A extends unknown[] = unknown[], R = unknown, F extends string = string> {
    readonly stale: StaleMode;
    readonly refusals: readonly F[];
    readonly [queryType]?: (...args: A) => R;
}
/** @internal */
export declare function command<P>(): <const F extends string = never, const X extends object = {}>(spec: {
    admission: Admission;
    refusals?: readonly F[];
    validate?(payload: unknown): payload is P;
} & X) => CommandSpec<P, F> & X;
/** @internal */
export declare function query<A extends unknown[], R>(): <const F extends string = never, const X extends object = {}>(spec: {
    stale: StaleMode;
    refusals?: readonly F[];
} & X) => QuerySpec<A, R, F> & X;
/** @internal */
export type CommandTable = Record<string, CommandSpec<any, string>>;
/** @internal */
export type QueryTable = Record<string, QuerySpec<any[], any, string>>;
/** @internal */
export type PayloadOf<S> = S extends CommandSpec<infer P, any> ? P : never;
/** @internal */
export type ArgsOf<S> = S extends QuerySpec<infer A, any, any> ? A : never;
/** @internal */
export type ResultOf<S> = S extends QuerySpec<any, infer R, any> ? R : never;
/** @internal */
export type RefusalOf<S> = S extends {
    readonly refusals: readonly (infer F extends string)[];
} ? F : never;
/** @internal */
export type CommandPayload<C extends CommandTable> = PayloadOf<C[keyof C]>;
/** @internal */
export interface RefusalResult<F extends string> {
    readonly [REFUSAL]: true;
    readonly refusal: Refusal<F>;
}
/** @internal */
export type QueryResult<R, F extends string> = R | TransferResult<R> | RefusalResult<F>;
/** @internal `satisfies QueryTable` preserves kinds for exhaustive `QueryHandlers<typeof table>`. */
export type QueryHandlers<Q extends QueryTable> = {
    [K in keyof Q]: (args: string extends keyof Q ? any : ArgsOf<Q[K]>, context: unknown) => QueryResult<ResultOf<Q[K]>, RefusalOf<Q[K]>> | Promise<QueryResult<ResultOf<Q[K]>, RefusalOf<Q[K]>>>;
};
/** @internal */
export declare function refuse<const F extends string>(code: BaseRefusal | F, detail?: string): RefusalResult<F>;
/** @internal */
export declare function isRefusalResult(value: unknown): value is RefusalResult<string>;
/** @internal Edge validation for unknown kinds and payloads. */
export declare function checkCommand<C extends CommandTable>(table: C, kind: string, payload: unknown): Refusal | undefined;
/** @internal */
export type QueryFence = {
    session: string;
    incarnation: number;
    epoch: number;
    afterSeq: number;
};
/** @internal */
export type QueryReply = Read<unknown, string>;
/** @internal */
export type QueryMethods = {
    query(fence: QueryFence, kind: string, args: unknown[]): QueryReply;
};
/** @internal */
export declare function isQueryReply(value: unknown): value is QueryReply;
/** @internal */
export type HoldReason = 'composition' | 'draft' | 'gesture' | 'mode';
/** @internal */
export interface Hold {
    readonly reason: HoldReason;
    readonly released: boolean;
    release(): void;
}
/** @internal */
export type ReadOptions = {
    within?: Hold;
    signal?: AbortSignal;
};
/** @internal */
export type SessionStage = 'ready' | 'recovering' | 'failed' | 'disposed';
/** @internal */
export interface SessionReads<Q extends QueryTable> {
    read<K extends keyof Q & string>(kind: K, args: ArgsOf<Q[K]>, options?: ReadOptions): Promise<Read<ResultOf<Q[K]>, RefusalOf<Q[K]>>>;
    hold(reason: HoldReason, options?: {
        signal?: AbortSignal;
    }): Hold;
    readonly stage: SessionStage;
}
/** @internal */
export type WorkerEditSession<P, S, E, X, Q extends QueryTable = {}> = EditSession<P, S, E, X> & SessionReads<Q>;
/** @internal Per-kind payload view of the same session object. */
export type CommandSession<C extends CommandTable, S, E, X, Q extends QueryTable = {}> = Omit<WorkerEditSession<CommandPayload<C>, S, E, X, Q>, 'submit'> & {
    submit<K extends keyof C & string>(kind: K, payload: PayloadOf<C[K]>, options?: SubmitOptions<X>): CommandTicket;
};
export {};

// ---- shared/office-session/commands/types.d.ts ----
import { type DeferredReply, type TransferResult } from '../protocol';
import type { MethodPolicy, SessionFailure } from '../types';
import type { CommandTable, QueryFence, QueryHandlers, QueryReply, QueryTable, WorkerEditSession } from './spec';
/** @internal Reply error name for an indeterminate command outcome. */
export declare const COMMAND_OUTCOME_UNKNOWN = "CommandOutcomeUnknown";
/** @internal Shared refusal and cancellation reasons; formats may supply other strings. */
export type CommandRefusalReason = 'stale-epoch' | 'stale-incarnation' | 'wrong-session' | 'stale-version' | 'dependency-refused' | 'unknown-outcome' | 'disposed';
/** @internal Error data preserved by the existing RPC reply protocol. */
export type CommandOutcomeUnknownError = Error & {
    name: typeof COMMAND_OUTCOME_UNKNOWN;
    refusal: {
        seq: number;
    };
};
/** @internal Command with a confirmed basis and a sequence assigned at send time. */
export type CommandEnvelope<P> = {
    session: string;
    incarnation: number;
    seq: number;
    basis: {
        epoch: number;
        version: number;
    };
    afterSeq?: number;
    expectVersion?: boolean;
    kind: string;
    payload: P;
    inputTime?: number;
    lastInputTime?: number;
    group?: string;
    replay?: true;
};
/** @internal Settled worker outcome for one command sequence. */
export type CommandOutcome = {
    seq: number;
    status: 'applied' | 'noop' | 'refused' | 'cancelled' | 'duplicate';
    reason?: string;
    committed?: false;
    affectedSeqs?: number[];
};
/** @internal Full or incremental frame identified by command coverage and revision. */
export type FrameEnvelope<F> = {
    coversThroughSeq: number;
    revision: number;
    baseRevision?: number;
    payload: F;
};
/** @internal Acknowledgement with authoritative state and an optional covering frame. */
export type CommandAck<S, E, F> = {
    session: string;
    incarnation: number;
    epoch: number;
    throughSeq: number;
    version: number;
    outcomes: CommandOutcome[];
    selection?: S;
    editState?: E;
    frame?: FrameEnvelope<F>;
    revisions?: {
        selection: number;
        layout: number;
    };
};
/** @internal Snapshot captured at a settled command boundary. */
export type Checkpoint = {
    state: Uint8Array;
    epoch: number;
    version: number;
    throughSeq: number;
};
/** @internal Methods registered on the existing session RPC. */
export type CommandMethods<P, S, E, F> = {
    edit(envelope: CommandEnvelope<P>): CommandAck<S, E, F>;
    checkpoint(): Checkpoint;
};
/** @internal Local outcome; unsent cancellations have no sequence. */
export type TicketOutcome = Omit<CommandOutcome, 'seq'> & {
    seq?: number;
};
/** @internal Non-rejecting presentation settlement for tickets and coverage waiters. */
export type PresentationStatus = 'presented' | 'dropped';
/** @internal Immediate local identity with a live send-time sequence getter. */
export type CommandTicket = {
    readonly id: number;
    readonly seq: number | undefined;
    readonly ack: Promise<TicketOutcome>;
    readonly presented: Promise<PresentationStatus>;
};
/** @internal Main-side command lane, distinct from resident scheduler lanes. */
export type CommandLane = 'input' | 'collab' | 'background';
/** @internal Submission metadata, causal ticket dependency, and optional local echo. */
export type SubmitOptions<X> = {
    lane?: CommandLane;
    after?: CommandTicket;
    expectVersion?: boolean;
    inputTime?: number;
    lastInputTime?: number;
    group?: string;
    coalesceKey?: string;
    fence?: boolean;
    echo?: X;
};
/** @internal Unsent command accepted by the outbox and coalescing hook. */
export type CommandDraft<P, X> = {
    kind: string;
    payload: P;
    options: SubmitOptions<X>;
};
/** @internal Format-owned echo associated with a local ticket and optional sent sequence. */
export type EchoEntry<X> = {
    ticket: CommandTicket;
    seq?: number;
    value: X;
};
/** @internal Synchronous painters; present installs, retires, and redraws atomically. */
export type EchoHooks<X, F> = {
    paint(pending: readonly EchoEntry<X>[]): void;
    drop(entries: readonly EchoEntry<X>[], reason: string): void;
    present(frame: FrameEnvelope<F>, retired: readonly EchoEntry<X>[], remaining: readonly EchoEntry<X>[]): void | boolean;
};
/** @internal Painter timing within one synchronous submit, acknowledgement, or frame event. */
export type EchoOverrunSample = {
    hook: keyof EchoHooks<unknown, unknown>;
    elapsedMs: number;
    eventTotalMs: number;
};
/** @internal Latest confirmed worker state and installed frame coverage. */
export type EditSessionState<S, E> = {
    epoch: number;
    version: number;
    throughSeq: number;
    presentedThroughSeq: number;
    selection: S | undefined;
    editState: E | undefined;
};
/** @internal Shared asynchronous editor contract for worker and legacy backends. */
export interface EditSession<P, S, E, X> {
    readonly backend: 'worker' | 'legacy-sync' | 'legacy-in-thread';
    readonly state: EditSessionState<S, E>;
    submit(kind: string, payload: P, options?: SubmitOptions<X>): CommandTicket;
    attachEcho?(tickets: readonly CommandTicket[], preview: (draft: CommandDraft<P, X>, originalIntentCount: number) => X | undefined): boolean;
    whenPresented(seq: number): Promise<PresentationStatus>;
    subscribe(listener: (state: EditSessionState<S, E>) => void): () => void;
    flush(): Promise<void>;
    dispose(): Promise<void>;
}
/** @internal Initial acknowledgement and structural subset of the existing RPC client. */
export type CommandConnection<P, S, E, F> = {
    initial: CommandAck<S, E, F>;
    client: {
        readonly call: {
            edit(envelope: CommandEnvelope<P>): Promise<CommandAck<S, E, F>>;
            checkpoint(): Promise<Checkpoint>;
            /** @internal Posts queries after all edits previously handed to this connection. */
            query?(fence: QueryFence, kind: string, args: unknown[]): Promise<QueryReply>;
        };
        readonly failure: SessionFailure | undefined;
        onFailure(listener: (failure: SessionFailure) => void): () => void;
        /** @internal Subscribes to frames posted by this connection's worker. */
        onFrame?(listener: (frame: FrameEnvelope<F>) => void): () => void;
        onAck?(listener: (ack: CommandAck<S, E, F>) => void): () => void;
        dispose(): Promise<void>;
    };
};
/** @internal Confirmed basis and live version provided for format-owned rebasing. */
export type ApplyContext = {
    basis: CommandEnvelope<unknown>['basis'];
    currentVersion: number;
    replayPreviouslyApplied: boolean;
};
/** @internal Format apply result; noop, refusal, cancellation, and applied with committed: false must preserve the version. */
export type ApplyResult<S, E> = {
    status: 'applied';
    committed?: false;
    selection?: S;
    editState?: E;
} | {
    status: 'noop';
    selection?: S;
    editState?: E;
} | {
    status: 'cancelled';
    reason?: string;
    selection?: S;
    editState?: E;
} | {
    status: 'refused';
    reason: string;
    selection?: S;
    editState?: E;
    affectedSeqs?: number[];
};
/** @internal Worker identity, replay ledger, version access, and format hooks. */
export type CommandHostOptions<P, S, E, F> = {
    session: string;
    incarnation: number;
    epoch: number;
    throughSeq?: number;
    retainedOutcomes?: number;
    /** @internal Refuses pruned duplicates with this reason; absent preserves duplicate outcomes. */
    prunedDuplicateReason?: string;
    refusedSeqs?: readonly number[];
    replayAppliedSeqs?: readonly number[];
    commands?: CommandTable;
    queries?: {
        specs: QueryTable;
        handlers: QueryHandlers<QueryTable>;
    };
    readOnly?(): boolean;
    layoutRevision?(): number;
    checkRefusals?: boolean;
    selectionEquals?(previous: S | undefined, next: S | undefined): boolean;
    editStateEquals?(previous: E | undefined, next: E | undefined): boolean;
    getVersion(): number;
    apply(envelope: CommandEnvelope<P>, context: ApplyContext): ApplyResult<S, E> | Promise<ApplyResult<S, E>>;
    deferFrame?(envelope: CommandEnvelope<P>, result: ApplyResult<S, E>): boolean;
    frame(input: {
        coversThroughSeq: number;
        version: number;
    }): FrameEnvelope<F> | TransferResult<FrameEnvelope<F>> | undefined | Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>> | undefined>;
    snapshot(): Uint8Array | Promise<Uint8Array>;
};
/** @internal Edit and checkpoint handlers with admission-time scheduler policies. */
export interface CommandHost<P, S, E, F> {
    handler(context: unknown, envelope: CommandEnvelope<P>): Promise<CommandAck<S, E, F> | TransferResult<CommandAck<S, E, F>>>;
    policy(envelope: CommandEnvelope<P>): MethodPolicy;
    checkpoint(context: unknown): Promise<Checkpoint | TransferResult<Checkpoint>>;
    checkpointPolicy(): MethodPolicy;
    query?(context: unknown, fence: QueryFence, kind: string, args: unknown[]): Promise<QueryReply | TransferResult<QueryReply>> | DeferredReply<QueryReply>;
    queryPolicy?(fence: QueryFence, kind: string, args: unknown[]): MethodPolicy;
    revise?(change?: {
        selection?: S;
        editState?: E;
    }): CommandAck<S, E, F> | undefined;
    /** @internal Builds a frame at a command boundary after admitted edits. */
    frame(): Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>>>;
}
/** @internal Factory contract implemented by host.ts. */
export type CreateCommandHost = <P, S, E, F>(options: CommandHostOptions<P, S, E, F>) => Required<CommandHost<P, S, E, F>>;
/** @internal Queued merged draft and all tickets sharing its eventual sequence and outcome. */
export type OutboxEntry<P, X> = {
    readonly id: number;
    readonly draft: CommandDraft<P, X>;
    readonly tickets: readonly CommandTicket[];
};
/** @internal Pure outbox coalescing and input fairness configuration. */
export type OutboxOptions<P, X> = {
    coalesce?(prev: CommandDraft<P, X>, next: CommandDraft<P, X>): CommandDraft<P, X> | undefined;
    inputBurstCap?: number;
};
/** @internal Unsent queue; takeNext removes an eligible entry subject to dependencies and fences. */
export interface Outbox<P, X> {
    readonly size: number;
    readonly entries: readonly OutboxEntry<P, X>[];
    enqueue(draft: CommandDraft<P, X>, ticket: CommandTicket): OutboxEntry<P, X>;
    takeNext(eligible: (entry: OutboxEntry<P, X>) => boolean, hold?: (entry: OutboxEntry<P, X>) => boolean): OutboxEntry<P, X> | undefined;
    remove(predicate: (entry: OutboxEntry<P, X>) => boolean): readonly OutboxEntry<P, X>[];
    clear(): readonly OutboxEntry<P, X>[];
}
/** @internal Factory contract implemented by outbox.ts; never assigns command sequences. */
export type CreateOutbox = <P, X>(options: OutboxOptions<P, X>) => Outbox<P, X>;
/** @internal Result of synchronously reconciling a candidate frame. */
export type FrameDisposition = 'presented' | 'stale' | 'base-mismatch';
/** @internal Pure presenter hooks, echo bound, and injectable painter clock. */
export type PresenterOptions<X, F> = {
    echo?: EchoHooks<X, F>;
    /** @internal Releases rejected frame resources; accepted frames belong to the painter. */
    discardFrame?(frame: FrameEnvelope<F>): void;
    sameCoverageFrames?: boolean;
    /** @internal Commits presentation only after the echo hook accepts it; off by default. */
    transactionalPresentation?: boolean;
    /** @internal Handles echo capacity overflow; defaults to declining new entries. */
    echoOverflow?: 'decline' | 'evict-oldest';
    maxEchoEntries?: number;
    now?(): number;
    requestFullFrame(): void;
    onEchoOverrun?(sample: EchoOverrunSample): void;
};
/** @internal Coverage and echo reconciler; reset invalidates coverage while retaining echo and waiters. */
export interface Presenter<X, F> {
    readonly presentedThroughSeq: number;
    readonly revision: number | undefined;
    readonly pending: readonly EchoEntry<X>[];
    event<T>(callback: () => T): T;
    add(entry: EchoEntry<X>): boolean;
    attach(entries: readonly EchoEntry<X>[]): boolean;
    replace(tickets: readonly CommandTicket[], entry?: EchoEntry<X>): void;
    drop(tickets: readonly CommandTicket[], reason: string): void;
    present(frame: FrameEnvelope<F>): FrameDisposition;
    whenPresented(seq: number): Promise<PresentationStatus>;
    reset(): void;
    dispose(): void;
}
/** @internal Factory contract implemented by presenter.ts. */
export type CreatePresenter = <X, F>(options: PresenterOptions<X, F>) => Presenter<X, F>;
/** @internal Recovery progress, replay divergence, and format history reset signal. */
export type RecoveryEvent = {
    phase: 'recovering' | 'recovered' | 'diverged' | 'failed';
    incarnation: number;
    attempt: number;
    diverged: number[];
    historyReset: boolean;
    failure?: SessionFailure;
};
/** @internal Format reopen and checkpoint hooks used by main-side recovery. */
export type RecoveryHooks<O, P, S, E, F> = {
    reopen(input: {
        source: O;
        checkpoint?: Checkpoint;
        session: string;
        incarnation: number;
        epoch: number;
        refusedSeqs: readonly number[];
        replayAppliedSeqs: readonly number[];
    }): Promise<CommandConnection<P, S, E, F>>;
    checkpoint?(connection: CommandConnection<P, S, E, F>): Promise<Checkpoint>;
    onRecovery?(event: RecoveryEvent): void;
};
/** @internal Worker-backed session limits, recovery, reconciliation, and injected scheduling. */
export type WorkerEditSessionOptions<O, P, S, E, F, X, Q extends QueryTable = {}> = {
    source: O;
    connection: CommandConnection<P, S, E, F>;
    queries?: Q;
    holdTimeoutMs?: number;
    timer?(callback: () => void, ms: number): () => void;
    recovery: RecoveryHooks<O, P, S, E, F>;
    echo?: EchoHooks<X, F>;
    /** @internal Releases rejected frame resources; accepted frames belong to the painter. */
    discardFrame?(frame: FrameEnvelope<F>): void;
    sameCoverageFrames?: boolean;
    /** @internal Commits presentation only after the echo hook accepts it; off by default. */
    transactionalPresentation?: boolean;
    /** @internal Handles echo capacity overflow; defaults to declining new entries. */
    echoOverflow?: 'decline' | 'evict-oldest';
    coalesce?(prev: CommandDraft<P, X>, next: CommandDraft<P, X>): CommandDraft<P, X> | undefined;
    /** @internal Keeps a matching draft queued while the matching command it follows is in flight, so a burst coalesces into one send. */
    holdForCoalescing?(draft: CommandDraft<P, X>): boolean;
    requestFullFrame(connection: CommandConnection<P, S, E, F>): Promise<CommandAck<S, E, F>>;
    maxInFlight?: number;
    inputBurstCap?: number;
    maxEchoEntries?: number;
    maxRecoveryAttempts?: number;
    checkpointPolicy?(stats: {
        applied: number;
        bytes: number;
    }): boolean;
    commandBytes?(envelope: CommandEnvelope<P>): number;
    /** @internal May share caller data only with a transport that copies before returning. */
    copyCommandForDispatch?(envelope: CommandEnvelope<P>): CommandEnvelope<P>;
    now?(): number;
    enqueue?(callback: () => void): void;
    onEchoOverrun?(sample: EchoOverrunSample): void;
    onAck?(ack: CommandAck<S, E, F>): void;
};
/** @internal Factory contract implemented by session.ts. */
export type CreateWorkerEditSession = <O, P, S, E, F, X, Q extends QueryTable = {}>(options: WorkerEditSessionOptions<O, P, S, E, F, X, Q>) => WorkerEditSession<P, S, E, X, Q>;
/** @internal Immediate legacy mutation with host-confirmed rendering and unchanged backend choice. */
export type LegacyEditSessionOptions<P, S, E, X> = {
    backend: 'legacy-sync' | 'legacy-in-thread';
    initial: EditSessionState<S, E>;
    apply(kind: string, payload: P, options?: SubmitOptions<X>): {
        outcome: Omit<CommandOutcome, 'seq'>;
        state: EditSessionState<S, E>;
    };
    render(throughSeq: number, presented: (coversThroughSeq: number) => void): void;
    dispose?(): void | Promise<void>;
};
/** @internal Factory contract implemented by legacy.ts. */
export type CreateLegacyEditSession = <P, S, E, X>(options: LegacyEditSessionOptions<P, S, E, X>) => EditSession<P, S, E, X>;
/** @internal Validates envelope metadata without asserting the format payload type. */
export declare function isCommandEnvelope(value: unknown): value is CommandEnvelope<unknown>;
/** @internal Validates acknowledgement metadata while leaving format-owned values unknown. */
export declare function isCommandAck(value: unknown): value is CommandAck<unknown, unknown, unknown>;
/** @internal Identifies the unknown-outcome reply error and its preserved sequence. */
export declare function isCommandOutcomeUnknownError(value: unknown): value is CommandOutcomeUnknownError;

// ---- shared/office-session/fontDecoder.d.ts ----
export type FontDecoder = (bytes: ArrayBuffer | Uint8Array, options: {
    codec: string;
    dict?: ArrayBuffer | Uint8Array;
    maxOutput?: number;
}) => Promise<{
    font: ArrayBuffer;
    payload?: ArrayBuffer;
}>;
export type FontDecodeResult = Awaited<ReturnType<FontDecoder>>;
export type FontDecoderMethods = {
    decodeFontFace(bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], wasm?: ArrayBuffer | WebAssembly.Module): FontDecodeResult;
};
export declare function rpcFontDecoder(call: (bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], transfer: Transferable[]) => Promise<FontDecodeResult>): FontDecoder;
export declare function fontBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer>;
export declare function copyDecodedFace(face: {
    font: Uint8Array;
    payload?: Uint8Array;
}): FontDecodeResult;

// ---- shared/office-session/host.d.ts ----
import { type TaskExecutor } from './scheduler';
import type { SessionTransport } from './transport';
import { SessionFailure, type MethodHandlers, type MethodPolicies, type SessionEvents, type SessionMethods, type SessionScheduler } from './types';
export interface SessionHostOptions<M extends SessionMethods, C> {
    handlers: MethodHandlers<M, C>;
    policies: MethodPolicies<M>;
    context: C;
    executor?: TaskExecutor;
    onDispose?(): void;
    /** Names a trap's cause, raw or already wrapped by a handler; without one it fails the session as `trap`. */
    trapFailure?(trap: SessionFailure): SessionFailure | undefined;
}
export interface SessionHost<E extends SessionEvents> {
    readonly scheduler: SessionScheduler;
    emit<K extends keyof E & string>(name: K, payload: E[K], transfer?: Transferable[]): void;
    /** Tells the client the running call is still working; safe to call from a long synchronous engine call. */
    progress(): void;
}
/** Runs method handlers serially through the resident scheduler. */
export declare function createSessionHost<M extends SessionMethods, E extends SessionEvents, C>(transport: SessionTransport, options: SessionHostOptions<M, C>): SessionHost<E>;

// ---- shared/office-session/index.d.ts ----
export * from './types';
export * from './protocol';
export * from './transport';
export * from './client';
export * from './host';
export * from './scheduler';

// ---- shared/office-session/protocol.d.ts ----
import type { SessionFailureCode } from './types';
export type ClientMessage = {
    protocol: 1;
    kind: 'wasm-compile';
} | {
    protocol: 1;
    kind: 'call';
    id: number;
    method: string;
    args: unknown[];
} | {
    protocol: 1;
    kind: 'dispose';
};
export interface ReplyError {
    name: string;
    message: string;
    refusal?: unknown;
}
export type HostMessage = {
    protocol: 1;
    kind: 'wasm-module';
    url: string;
    module: WebAssembly.Module;
    hydration?: string;
    version?: string;
    sequence?: number;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: true;
    value: unknown;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: false;
    error: ReplyError;
} | {
    protocol: 1;
    kind: 'event';
    name: string;
    payload: unknown;
} | {
    protocol: 1;
    kind: 'progress';
    id: number;
} | {
    protocol: 1;
    kind: 'failure';
    code: SessionFailureCode;
    message: string;
    diagnostics?: string;
};
declare const TRANSFER: unique symbol;
declare const DEFERRED: unique symbol;
export interface DeferredReply<T> {
    readonly [DEFERRED]: true;
    readonly promise: Promise<T | TransferResult<T>>;
}
/** Releases the foreground lane while a reply is pending. */
export declare function deferReply<T>(promise: Promise<T | TransferResult<T>>): DeferredReply<T>;
export declare function isDeferredReply(value: unknown): value is DeferredReply<unknown>;
export interface TransferResult<T> {
    readonly [TRANSFER]: true;
    readonly value: T;
    readonly transfer: Transferable[];
}
/** Transfers ownership of a handler result when replying. */
export declare function transferable<T>(value: T, transfer: Transferable[]): TransferResult<T>;
export declare function isTransferResult(value: unknown): value is TransferResult<unknown>;
/** `Object.hasOwn` for ES2020 targets. */
export declare function hasOwn(value: object, key: PropertyKey): boolean;
export declare function isClientMessage(value: unknown): value is ClientMessage;
export declare function isHostMessage(value: unknown): value is HostMessage;
export {};

// ---- shared/office-session/scheduler.d.ts ----
export type SchedulerLane = 'input' | 'collab' | 'interactive' | 'background';
declare const FOREGROUND_LANES: readonly ["input", "collab", "interactive"];
type ForegroundLane = (typeof FOREGROUND_LANES)[number];
export declare const INPUT_SLICE_MS = 8;
export declare const IDLE_SLICE_MS = 24;
export declare const INPUT_ACTIVE_MS = 1000;
export interface SchedulerMessage {
    lane: ForegroundLane;
    run(): Promise<void> | void;
    userInput?: boolean;
    holdsIdleTasks?: boolean;
    reorderable?: boolean;
    mutates?: boolean;
    reframes?: boolean;
    key?: string;
    replaceableBy?: string;
    supersede?(): void;
}
export type TaskStep = 'done' | 'yield';
export interface SchedulerTask {
    kind: string;
    version: number;
    generation: number;
    run(budgetMs: number): Promise<TaskStep> | TaskStep;
    idleAfterInputMs?: number;
    onStale?(): 'cancel' | 'continue';
    cancel?(): void;
    fail?(error: unknown): void;
}
export interface PureTask<Input = unknown, Result = unknown> {
    kind: string;
    version: number;
    generation: number;
    input: Input;
    transfer?: Transferable[];
    compute(input: Input): Result;
    install(result: Result): Promise<void> | void;
    cancel?(reason?: unknown): void;
}
export interface TaskExecutor {
    run<Input, Result>(job: {
        kind: string;
        input: Input;
        transfer?: Transferable[];
    }): Promise<Result>;
}
export interface SchedulerHost {
    now(): number;
    turn(callback: () => void): void;
    timer(callback: () => void, ms: number): () => void;
    executor?: TaskExecutor;
    failed?(error: unknown): void;
}
export interface ResidentScheduler {
    submit(message: SchedulerMessage): void;
    schedule(task: SchedulerTask): void;
    dispatch<Input, Result>(task: PureTask<Input, Result>): void;
    budget(): number;
    readonly version: number;
    readonly generation: number;
    bump(changes: {
        version?: boolean;
        generation?: boolean;
    }): void;
    pending(): {
        foreground: number;
        background: number;
    };
}
export declare function createResidentScheduler(host: SchedulerHost): ResidentScheduler;
export {};

// ---- shared/office-session/transport.d.ts ----
/** Endpoint seam: ordered reliable delivery; ports and sockets may copy instead of transfer. */
export interface SessionTransport {
    post(message: unknown, transfer?: Transferable[]): void;
    listen(listener: (message: unknown) => void): () => void;
    onError(listener: (error: unknown) => void): () => void;
    close(): void;
}
export interface SessionScope {
    postMessage(message: unknown, transfer: Transferable[]): void;
    addEventListener(type: string, listener: EventListener): void;
    removeEventListener(type: string, listener: EventListener): void;
    close?(): void;
    start?(): void;
}
/** Adapts an existing worker without constructing one. */
export declare function createWorkerTransport(worker: Worker): SessionTransport;
/** Adapts a worker scope or message port. */
export declare function createScopeTransport(scope: SessionScope): SessionTransport;

// ---- shared/office-session/types.d.ts ----
import type { DeferredReply, TransferResult } from './protocol';
import type { ResidentScheduler } from './scheduler';
export declare const SESSION_SUPERSEDED = "SessionSuperseded";
export type SessionScheduler = Pick<ResidentScheduler, 'schedule' | 'dispatch' | 'bump' | 'budget' | 'pending' | 'version' | 'generation'>;
/** Format-owned, structured-cloneable methods. */
export type SessionMethods = Record<string, (...args: any[]) => unknown>;
export type Promisified<M extends SessionMethods> = {
    [K in keyof M]: (...args: Parameters<M[K]>) => Promise<Awaited<ReturnType<M[K]>>>;
};
export type MethodHandlers<M extends SessionMethods, C> = {
    [K in keyof M]: (context: C, ...args: Parameters<M[K]>) => Awaited<ReturnType<M[K]>> | TransferResult<Awaited<ReturnType<M[K]>>> | DeferredReply<Awaited<ReturnType<M[K]>>> | Promise<Awaited<ReturnType<M[K]>> | TransferResult<Awaited<ReturnType<M[K]>>> | DeferredReply<Awaited<ReturnType<M[K]>>>>;
};
export interface MethodPolicy {
    lane: 'input' | 'collab' | 'interactive';
    mutates?: boolean;
    reframes?: boolean;
    reorderable?: boolean;
    key?: string;
    replaceableBy?: string;
    userInput?: boolean;
    holdsIdleTasks?: boolean;
}
export type MethodPolicies<M extends SessionMethods> = {
    [K in keyof M]: MethodPolicy | ((...args: Parameters<M[K]>) => MethodPolicy);
};
export type SessionEvents = Record<string, unknown>;
export interface SessionState {
    format: string;
    stage: 'preview' | 'ready';
    version: number;
    dirty: boolean;
}
export type SessionFailureCode = 'crash' | 'trap' | 'out-of-memory' | 'message' | 'silence' | 'disposed';
/** A terminal session failure. */
export declare class SessionFailure extends Error {
    readonly code: SessionFailureCode;
    readonly diagnostics?: string | undefined;
    constructor(code: SessionFailureCode, message: string, diagnostics?: string | undefined);
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

// ---- wasm/generated/xlsx_wasm.d.ts ----
/* tslint:disable */
/* eslint-disable */
export interface DecodedFontFace { font: Uint8Array; payload?: Uint8Array; }


/**
 * a workbook handle exposed to js; wraps the pure `Session`.
 */
export class XlsxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    /**
     * accept a proposal as one agent transaction; returns the edit envelope
     * plus `proposalId`, or a `{"code":"staleProposal",...}` error when the base moved.
     */
    acceptProposalJson(args: string): string;
    adoptPeerVersion(version: string): void;
    /**
     * apply an edit batch all-or-nothing; volatile functions see only the request's clock.
     */
    applyEditsJson(request: string): string;
    applyFormatJson(args: string): string;
    /**
     * apply a raw op list as one user transaction; returns `SheetInfo` json.
     */
    applyOpsJson(transaction_json: string): string;
    applyOpsProfiledJson(transaction_json: string): string;
    applyUpdateJson(update: Uint8Array): string;
    beginPeerSnapshot(records: number, bytes: number): void;
    calculationStatusJson(): string;
    captureFormatJson(args: string): string;
    /**
     * the editable representation of one cell.
     */
    cellJson(args: string): string;
    cellPositionJson(args: string): string;
    cellRectJson(args: string): string;
    /**
     * the chart under a viewport-local point, or `null`.
     */
    chartAtPointJson(args: string): string;
    chartsJson(sheet: number): string;
    /**
     * Stop observation and discard queued events.
     */
    clearUpdateObservation(): void;
    conditionalFormatsJson(sheet: number): string;
    /**
     * @experimental
     */
    displayListForJson(sheet: number, viewport_json: string): string;
    /**
     * serialized `DisplayList` for a serialized `Viewport`.
     */
    displayListJson(viewport_json: string): string;
    displayListProfiledJson(viewport_json: string): string;
    /**
     * The session-scoped version of the committed workbook state.
     */
    documentVersion(): string;
    /**
     * Poll one event: origin byte (`0` local, `1` remote), then update; empty means none.
     */
    drainUpdateEvent(): Uint8Array;
    /**
     * enter one cell edit; returns updated `SheetInfo` json.
     */
    editCellJson(args: string): string;
    editCellProfiledJson(args: string): string;
    /**
     * enter a batch of cell edits as one undo step; returns `SheetInfo` json.
     */
    editCellsJson(args: string): string;
    editChartsJson(edits: string): string;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    endPeerSnapshot(): void;
    exportMarkdownJson(options: string, markdown_options: string): string;
    /**
     * Export the committed workbook with the version it was read at; nothing recalculates.
     */
    exportStructuredJson(options: string): string;
    findTextJson(request: string): string;
    historyStateJson(): string;
    listDefinedNamesJson(request: string): string;
    /**
     * the pending proposals: `{"proposals":[...]}`.
     */
    listProposalsJson(): string;
    mergedRangesJson(args: string): string;
    /**
     * slide a chart by a pixel delta as one undo step.
     */
    moveChartJson(args: string): string;
    moveRangeJson(args: string): string;
    nextPeerSnapshotChunk(): Uint8Array | undefined;
    /**
     * Open a replica with a positive, safe-integer client ID.
     */
    static openCollaborative(bytes: Uint8Array, client_id: number): XlsxDocument;
    static openRecoveryCheckpoint(bytes: Uint8Array, through_seq: number): XlsxDocument;
    static openWithCalculationJson(bytes: Uint8Array, context: string): XlsxDocument;
    static openWithPeerHydrationJson(bytes: Uint8Array, hydration: string): XlsxDocument;
    /**
     * open a workbook from raw `.xlsx` bytes.
     */
    static open(bytes: Uint8Array): XlsxDocument;
    patchRangeStyleJson(args: string): string;
    peerHydrationJson(): string;
    printDisplayListJson(args: string): string;
    /**
     * register an agent proposal (preview only); returns the stored `Proposal` json.
     */
    proposeJson(args: string): string;
    /**
     * a rectangular block of cells for clipboard copy.
     */
    rangeCellsJson(args: string): string;
    readCellsJson(request: string): string;
    recoveryCheckpoint(through_seq: number, preserve_history?: boolean | null): Uint8Array;
    /**
     * redo the last undone transaction; same shape as `undoJson`.
     */
    redoJson(): string;
    /**
     * reject a proposal by id; returns `{"removed":bool}`.
     */
    rejectProposalJson(args: string): string;
    /**
     * render the current sheet viewport to png bytes (raster feature only).
     */
    renderPng(viewport_json: string): Uint8Array;
    /**
     * render an a1 range (default: used range) at an optional scale to png.
     */
    renderRangePng(args: string): Uint8Array;
    /**
     * serialize the current workbook back to `.xlsx` bytes.
     */
    saveBytes(): Uint8Array;
    searchTextJson(args: string): string;
    selectionDataJson(args: string): string;
    selectionFormattingJson(args: string): string;
    /**
     * switch the active sheet by index.
     */
    setActiveSheet(index: number): void;
    setCalculationContextJson(context: string): void;
    setRangeNumberFormatJson(args: string): string;
    sheetCount(): number;
    sheetInfoForJson(sheet: number): string;
    /**
     * serialized `SheetInfo`: stable IDs, names, active index, content extent.
     */
    sheetInfoJson(): string;
    /**
     * Start queuing origin-prefixed Yrs update events for polling.
     */
    startUpdateObservation(): void;
    /**
     * undo the last transaction; returns `{"applied":bool,"sheetInfo":{...}}`.
     */
    undoJson(): string;
    validateEditsJson(request: string): string;
    /**
     * crate version string.
     */
    static version(): string;
    visibleMergedRangesJson(sheet: number, viewport_json: string): string;
    readonly clientId: number;
}

export class XlsxSnapshotBuilder {
    free(): void;
    [Symbol.dispose](): void;
    advance(records: number, bytes: number): boolean;
    finish(): XlsxDocument;
    constructor();
    push(chunk: Uint8Array): void;
}

/**
 * bytes of the last allocation the engine could not get, or 0.
 */
export function allocationFailure(): number;

/**
 * Decodes one font pack face (codec "zstd" or "woff2t+zstd") with this engine's linked decoder.
 * Returns owned bytes and an optional transformed payload. Non-zstd input passes through.
 * The default output limit is 64 MiB; woff2t decoding is always capped at 64 MiB.
 */
export function decodeFontFace(bytes: Uint8Array, codec: string, dict?: Uint8Array | null, maxOutput?: number | null): DecodedFontFace;

export function displayListCellLimit(): number;

export function exportXlsxMarkdownJson(bytes: Uint8Array, options: string, markdown_options: string): string;

/**
 * Export `.xlsx` bytes as read: stored formula results, no clock.
 */
export function exportXlsxStructuredJson(bytes: Uint8Array, options: string): string;

/**
 * wasm panics abort, so a trap reaches JS as a bare `RuntimeError:
 * unreachable executed`. Log the panic first so it stays diagnosable.
 */
export function install_panic_hook(): void;

export function renderXlsxMarkdownJson(content: string, options: string): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly __externref_table_alloc: () => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbg_xlsxdocument_free: (a: number, b: number) => void;
    readonly __wbg_xlsxsnapshotbuilder_free: (a: number, b: number) => void;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_start: () => void;
    readonly allocationFailure: () => number;
    readonly decodeFontFace: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number];
    readonly displayListCellLimit: () => number;
    readonly exportXlsxMarkdownJson: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly exportXlsxStructuredJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly install_panic_hook: () => void;
    readonly memory: WebAssembly.Memory;
    readonly renderXlsxMarkdownJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly xlsxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_adoptPeerVersion: (a: number, b: number, c: number) => [number, number];
    readonly xlsxdocument_applyEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyFormatJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyOpsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyOpsProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_beginPeerSnapshot: (a: number, b: number, c: number) => [number, number];
    readonly xlsxdocument_calculationStatusJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_captureFormatJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_cellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_cellPositionJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_cellRectJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_chartAtPointJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_chartsJson: (a: number, b: number) => [number, number, number, number];
    readonly xlsxdocument_clearUpdateObservation: (a: number) => void;
    readonly xlsxdocument_clientId: (a: number) => number;
    readonly xlsxdocument_conditionalFormatsJson: (a: number, b: number) => [number, number, number, number];
    readonly xlsxdocument_displayListForJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly xlsxdocument_displayListJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_displayListProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_documentVersion: (a: number) => [number, number];
    readonly xlsxdocument_drainUpdateEvent: (a: number) => [number, number, number, number];
    readonly xlsxdocument_editCellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editCellProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editChartsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly xlsxdocument_encodeStateVector: (a: number) => [number, number];
    readonly xlsxdocument_endPeerSnapshot: (a: number) => void;
    readonly xlsxdocument_exportMarkdownJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly xlsxdocument_exportStructuredJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_findTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_historyStateJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_listDefinedNamesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_mergedRangesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_moveChartJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_moveRangeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_nextPeerSnapshotChunk: (a: number) => [number, number, number];
    readonly xlsxdocument_open: (a: number, b: number) => [number, number, number];
    readonly xlsxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly xlsxdocument_openRecoveryCheckpoint: (a: number, b: number, c: number) => [number, number, number];
    readonly xlsxdocument_openWithCalculationJson: (a: number, b: number, c: number, d: number) => [number, number, number];
    readonly xlsxdocument_openWithPeerHydrationJson: (a: number, b: number, c: number, d: number) => [number, number, number];
    readonly xlsxdocument_patchRangeStyleJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_peerHydrationJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_printDisplayListJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_rangeCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_readCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_recoveryCheckpoint: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderPng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderRangePng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly xlsxdocument_searchTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_selectionDataJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_selectionFormattingJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_setActiveSheet: (a: number, b: number) => [number, number];
    readonly xlsxdocument_setCalculationContextJson: (a: number, b: number, c: number) => [number, number];
    readonly xlsxdocument_setRangeNumberFormatJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_sheetCount: (a: number) => number;
    readonly xlsxdocument_sheetInfoForJson: (a: number, b: number) => [number, number, number, number];
    readonly xlsxdocument_sheetInfoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly xlsxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_validateEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_version: () => [number, number];
    readonly xlsxdocument_visibleMergedRangesJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly xlsxsnapshotbuilder_advance: (a: number, b: number, c: number) => [number, number, number];
    readonly xlsxsnapshotbuilder_finish: (a: number) => [number, number, number];
    readonly xlsxsnapshotbuilder_new: () => number;
    readonly xlsxsnapshotbuilder_push: (a: number, b: number, c: number) => [number, number];
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;

// ---- wasm/loader.d.ts ----
import type { ChartInfo, ChartMutation } from '../charts';
/**
 * Wasm loader for the xlsx-wasm core.
 *
 * Call {@link initWasm} once before opening workbooks. The default browser path
 * streams the packaged wasm asset; non-fetch environments can pass bytes or a
 * precompiled module. Callers never see the JSON-string boundary.
 */
import type { XlsxDefinedNamesRequest, XlsxDefinedNamesResult } from '../definedNames';
import type { Direction, SelectionData } from '../selection/types';
import type { InitInput } from './generated/xlsx_wasm.js';
import { type FontDecoder } from '../../../../shared/office-session/fontDecoder';
export declare const decodeFontFace: FontDecoder;
import type { CollaborationReplica, CollaborationUpdateOrigin } from '../collaboration/types';
import type { ChartRegion, DisplayList, Rect } from '../display-list/types';
import type { XlsxCellAddress, XlsxEditRequest, XlsxEditResult, XlsxFindRequest, XlsxFindResult, XlsxReadRequest, XlsxReadResult, XlsxValidationResult } from '../edits';
import type { XlsxExportOptions, XlsxExportResult, XlsxMarkdownContent, XlsxMarkdownOptions, XlsxStructuredContent } from '../exports';
/**
 * A scrolled window into a sheet. `x`/`y` are content-pixel offsets into the
 * non-frozen body. Mirrors the Rust `Viewport`.
 */
export interface Viewport {
    x: number;
    y: number;
    width: number;
    height: number;
}
export interface PrintMetrics {
    dpi: number;
    maxDigitWidth: number;
    defaultRowHeightPt: number;
    defaultColumnWidth?: number;
    fontSizePt: number;
    fontFamily: string;
    fontAscent: number;
    fontDescent: number;
}
export type XlsxAutoFilterOp = {
    type: 'setAutoFilter';
    sheet: number;
    range: string | null;
};
/**
 * Chrome-facing sheet metadata: stable IDs, tab names, active index, and the
 * scrollable content extent of the active sheet. Mirrors the Rust `SheetInfo`.
 */
export interface SheetInfo {
    autoFilter?: string | null;
    sheetIds: string[];
    sheetNames: string[];
    activeSheet: number;
    contentWidth: number;
    contentHeight: number;
    frozenRows: number;
    frozenCols: number;
    initialScrollX: number;
    initialScrollY: number;
}
export interface CellPosition {
    x: number;
    y: number;
}
/**
 * Result of any mutating call: whether it changed the workbook and the
 * (possibly grown) sheet metadata. Mirrors the Rust `EditResult`.
 *
 * `changed` lists the a1 addresses (on the active sheet) of *other* cells whose
 * displayed value moved as a fallout of the edit — the dependents the recalc
 * pass recomputed. It excludes the directly edited cell(s) and is omitted by
 * cores built before the recalc engine, so treat it as additive.
 */
export interface EditResult {
    applied: boolean;
    sheetInfo: SheetInfo;
    changed?: string[];
    limitedCells?: string[];
}
/** Facade stage latencies of one profiled mutation, in ms. */
export interface EditProfile {
    validateMs: number;
    applyMs: number;
    recalcMs: number;
    resultMs: number;
}
export interface ProfiledEditResult extends EditResult {
    profile: EditProfile;
}
export interface DisplayListProfile {
    buildMs: number;
    encodeMs: number;
}
export interface ProfiledDisplayList {
    displayList: DisplayList;
    profile: DisplayListProfile;
}
export interface CalculationStatus {
    limitedCells: string[];
    /** The last recalculation ran out of its budget; the cells it could not evaluate keep their saved values. */
    budgetExhausted?: boolean;
}
export type NumberFormat = 'automatic' | 'plainText' | 'number' | 'percent' | 'scientific' | 'currency' | 'date' | 'time' | 'custom';
export type BorderPreset = 'all' | 'inner' | 'horizontal' | 'vertical' | 'outer' | 'left' | 'top' | 'right' | 'bottom' | 'none';
export type BorderStyle = 'solid' | 'dashed' | 'dotted' | 'double';
export type HorizontalAlignment = 'left' | 'center' | 'right';
export type VerticalAlignment = 'top' | 'middle' | 'bottom';
export type TextWrapping = 'overflow' | 'wrap' | 'clip';
export type StyleProperty = 'bold' | 'italic' | 'strikethrough' | 'fontFamily' | 'fontSize' | 'textColor' | 'fillColor' | 'borders' | 'horizontalAlignment' | 'verticalAlignment' | 'textWrapping';
export interface BorderPatch {
    preset?: BorderPreset;
    style?: BorderStyle;
    color?: string;
}
export interface RangeStylePatch {
    bold?: boolean;
    italic?: boolean;
    strikethrough?: boolean;
    fontFamily?: string;
    fontSize?: number;
    textColor?: string;
    fillColor?: string;
    border?: BorderPatch;
    horizontalAlignment?: HorizontalAlignment;
    verticalAlignment?: VerticalAlignment;
    textWrapping?: TextWrapping;
    clear?: StyleProperty[];
}
export type NumberFormatMutation = Exclude<NumberFormat, 'custom'> | 'increaseDecimal' | 'decreaseDecimal' | {
    type: 'custom';
    pattern: string;
};
export interface SelectionFormatting {
    numberFormat?: NumberFormat;
    numberFormatPattern?: string;
    fontFamily?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    strikethrough?: boolean;
    textColor?: string;
    fillColor?: string;
    borderPreset?: BorderPreset;
    borderStyle?: BorderStyle;
    borderColor?: string;
    horizontalAlignment?: HorizontalAlignment;
    verticalAlignment?: VerticalAlignment;
    textWrapping?: TextWrapping;
}
export interface CapturedFormat<TFormat = unknown> {
    rows: number;
    columns: number;
    formats: TFormat[];
    sourceStyles?: (number | null)[];
    sourceFormats?: unknown[];
}
export type CapturedColor = 'Auto' | {
    Rgb: string;
} | {
    Indexed: number;
} | {
    Theme: {
        idx: number;
        tint: number;
    };
};
export type CapturedFill = 'None' | {
    Solid: CapturedColor;
} | {
    Pattern: {
        pattern: import('../display-list/types').PatternType;
        fg?: CapturedColor | null;
        bg?: CapturedColor | null;
    };
};
export type CapturedBorderStyle = 'Thin' | 'Medium' | 'Thick' | 'Dashed' | 'Dotted' | 'Double' | 'Hair' | 'MediumDashed' | 'DashDot' | 'MediumDashDot' | 'DashDotDot' | 'MediumDashDotDot' | 'SlantDashDot';
export interface CapturedBorderEdge {
    style: CapturedBorderStyle;
    color?: CapturedColor | null;
}
export interface CapturedCellFormat {
    font: {
        name?: string | null;
        size_pt?: number | null;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        strike?: boolean;
        color?: CapturedColor | null;
    };
    fill: CapturedFill;
    border: {
        left?: CapturedBorderEdge | null;
        right?: CapturedBorderEdge | null;
        top?: CapturedBorderEdge | null;
        bottom?: CapturedBorderEdge | null;
    };
    numberFormat: {
        kind: 'builtin';
        id: number;
    } | {
        kind: 'custom';
        pattern: string;
    };
    alignment: {
        h?: 'General' | 'Left' | 'Center' | 'Right' | 'Fill' | 'Justify' | 'CenterContinuous' | 'Distributed' | null;
        v?: 'Top' | 'Center' | 'Bottom' | 'Justify' | 'Distributed' | null;
        wrap_text?: boolean;
        shrink_to_fit?: boolean;
    };
}
export interface CellPoint {
    row: number;
    col: number;
}
export interface MergedRange {
    start: CellPoint;
    end: CellPoint;
}
export interface WorkbookHyperlink {
    range: MergedRange;
    externalTarget?: string;
    location?: string;
    tooltip?: string;
    display?: string;
}
export interface DifferentialFormatInput {
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
    fill?: string;
}
export interface ConditionalFormatRuleInput {
    type: 'cellIs' | 'expression';
    operator?: string;
    formulas: string[];
}
export interface ConditionalFormatInfo {
    id: string;
    range: string;
    type: string;
    operator?: string;
    formulas: string[];
    priority: number;
    dxfId?: number;
    format: DifferentialFormatInput;
    editable: boolean;
}
export type ConditionalFormatOp = {
    type: 'addConditionalFormat';
    sheet: number;
    range: MergedRange;
    kind: 'cellIs' | 'expression';
    operator?: string;
    formulas: string[];
    format: DifferentialFormatInput;
} | {
    type: 'removeConditionalFormat';
    sheet: number;
    id: string;
};
export interface HistoryState {
    canUndo: boolean;
    canRedo: boolean;
    undoDepth: number;
    redoDepth: number;
}
export type WorkbookUpdateOrigin = CollaborationUpdateOrigin;
/**
 * `fromPeer` is true for every event a peer update caused through `applyUpdate`, including
 * the `'local'` re-encoding a replica forwards after adopting a whole-document state.
 */
export interface WorkbookUpdateInfo {
    fromPeer: boolean;
}
export type WorkbookUpdateListener = (update: Uint8Array, origin: WorkbookUpdateOrigin, info: WorkbookUpdateInfo) => void;
export interface WorkbookCalculationContext {
    nowSerial: number;
    randSeed: number;
}
export interface OpenWorkbookOptions {
    /** Open a Yrs-backed replica that can accept peer updates. */
    collaborative?: boolean;
    /** Peer-unique positive safe integer. Generated securely when omitted. */
    clientId?: number;
    calculation?: WorkbookCalculationContext;
}
/**
 * The editable view of one cell: A1 address, the exact string the user would
 * edit (formulas as `=...`, guarded literals with a leading `'`), and whether
 * that string is a formula. Mirrors the Rust `CellEdit`.
 */
export interface CellEdit {
    a1: string;
    input: string;
    isFormula: boolean;
    isArrayFormula?: boolean;
}
export interface XlsxTextSearchOptions {
    /** Defaults to false. */
    caseSensitive?: boolean;
    /** Maximum matches; defaults to 1000. */
    limit?: number;
}
/** Zero-based sheet, row, and column. */
export interface XlsxTextMatch {
    sheet: number;
    sheetId: string;
    sheetName: string;
    row: number;
    col: number;
    a1: string;
    /** Formatted display text. */
    text: string;
}
/** One cell of a batch edit: target coordinates plus the raw user input. */
export interface CellInputEdit {
    row: number;
    col: number;
    input: string;
    copyFrom?: CellPoint;
}
/**
 * One edit inside a proposal. Same shape as {@link CellInputEdit} plus the
 * sheet index — a proposal's cells carry their own sheet on the wire, so it
 * can span sheets rather than being pinned to the active one.
 */
export interface ProposalEdit extends CellInputEdit {
    sheet: number;
    numberFormat?: NumberFormatMutation;
}
/**
 * One cell of a pending proposal: where it lands and the *formatted display
 * texts* before/after applying it. `oldText`/`newText` are what the grid shows,
 * not the raw input — the ghost decoration paints `newText` directly. Mirrors
 * the Rust proposal cell.
 */
export interface ProposalCell {
    sheet: number;
    row: number;
    col: number;
    a1: string;
    oldText: string;
    newText: string;
}
/**
 * A pending, un-applied set of cell changes an agent has staged for human
 * review. `note` is the agent's rationale (may be null); `cells` are the
 * per-cell before/after previews. Applying it is one undo step.
 */
export interface Proposal {
    id: string;
    agentId: string;
    note: string | null;
    cells: ProposalCell[];
}
/** A drifted proposal cell: its sheet index and catalog id beside the coordinates. */
export interface StaleProposalTarget extends XlsxCellAddress {
    sheet: number;
}
/**
 * Thrown by {@link WorkbookHandle.acceptProposal} when the workbook changed
 * under a proposal since it was staged (an edit touched one of its base cells)
 * and `force` was not set. `cells` are the a1 addresses that moved, so the UI
 * can name them and offer a force-apply; `targets` also carry each one's sheet.
 */
export declare class StaleProposalError extends Error {
    readonly cells: string[];
    readonly targets: StaleProposalTarget[];
    constructor(cells: string[], targets?: StaleProposalTarget[]);
}
/** A display-list viewport exceeds the counted cell limit. */
export declare class DisplayTooLargeError extends Error {
    readonly cells: number;
    readonly maxCells: number;
    readonly code = "displayTooLarge";
    constructor(cells: number, maxCells: number);
}
/**
 * A typed handle over an open workbook, hiding the wasm object and its JSON
 * boundary. Call {@link WorkbookHandle.dispose} to free the wasm memory.
 */
export interface WorkbookHandle extends CollaborationReplica {
    readonly clientId: number;
    setCalculationContext(context: WorkbookCalculationContext | null): void;
    /** Available in both modes; encodes this handle's current Yrs state vector. */
    encodeStateVector(): Uint8Array;
    /** Available in both modes; pass a peer vector to encode only the missing state. */
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    /** Apply a peer update. Standalone handles throw the Rust `NotCollaborative` error. */
    applyUpdate(update: Uint8Array): EditResult;
    /** Observe owned update bytes from local commits and accepted remote updates. */
    onUpdate(listener: WorkbookUpdateListener): () => void;
    sheetInfo(): SheetInfo;
    sheetCount(): number;
    /** Metadata for `sheet`, with that index in `activeSheet`; leaves the workbook unchanged. */
    sheetInfoFor(sheet: number): SheetInfo;
    calculationStatus(): CalculationStatus;
    /** Clamps tracks to the grid and truncates oversized frames at the scroll origin. */
    displayList(viewport: Viewport): DisplayList;
    /** `displayList` with build and encode time measured inside the core. */
    displayListProfiled(viewport: Viewport): ProfiledDisplayList;
    printDisplayList(sheet: number, range: string, metrics: PrintMetrics, gridlines: boolean): DisplayList;
    /**
     * the chart under a viewport-local point on the active sheet, or `null`,
     * resolved against the current model from the same anchor geometry the
     * display list is built from. the returned `label` is always empty: a hit
     * test resolves anchors only and never reads a chart part.
     *
     * this answers for the model as it is now. chrome painting a frame should
     * hit-test that frame's own `charts` with `chartRegionAtPoint` instead, or a
     * mutation the canvas has not drawn yet will move the answer off the pixels.
     */
    chartAtPoint(viewport: Viewport, x: number, y: number): ChartRegion | null;
    /**
     * slide the chart frame `chart` names — a {@link ChartRegion} id — by a
     * content-pixel delta, clamped to the grid, as one undo step. the new anchor
     * is written back on save. throws for a chart pinned to the sheet by an
     * absolute anchor.
     */
    moveChart(sheet: number, chart: string, dx: number, dy: number): EditResult;
    setActiveSheet(index: number): void;
    /**
     * apply one user input to a cell (parses number/bool/formula/text) and
     * recalc its transitive dependents. one undo step; moved dependents come back
     * in `EditResult.changed`.
     */
    editCell(sheet: number, row: number, col: number, input: string): EditResult;
    /** `editCell` with the facade's stage timings attached. */
    editCellProfiled(sheet: number, row: number, col: number, input: string): ProfiledEditResult;
    /** apply a batch of inputs (paste path) as one undo step; dependents recalc. */
    editCells(sheet: number, edits: CellInputEdit[]): EditResult;
    /** Move a range within one sheet and rewrite references in one native undo step. */
    moveRange(sheet: number, source: string, destination: string): EditResult;
    setHyperlink(sheet: number, hyperlink: WorkbookHyperlink): EditResult;
    removeHyperlink(sheet: number, range: MergedRange): EditResult;
    charts(sheet: number): ChartInfo[];
    editCharts(edits: ChartMutation[]): EditResult;
    conditionalFormats(sheet: number): ConditionalFormatInfo[];
    /** raw op-list escape hatch for structural ops (insert/delete rows, merges…). */
    applyOps(ops: unknown[]): EditResult;
    /** `applyOps` with the facade's stage timings attached. */
    applyOpsProfiled(ops: unknown[]): ProfiledEditResult;
    undo(): EditResult;
    redo(): EditResult;
    /** the editable view of one cell (formula bar / in-cell editor prefill). */
    cell(sheet: number, row: number, col: number): CellEdit;
    /** Searches formatted text in sheet and row order. */
    searchText(query: string, options?: XlsxTextSearchOptions): XlsxTextMatch[];
    cellPosition(sheet: number, row: number, col: number): CellPosition;
    /** Absolute sheet rectangle in unzoomed pixels. */
    cellRect(sheet: number, row: number, col: number): Rect;
    selectionData(sheet: number, row: number, col: number, direction?: Direction | 'usedRange' | 'rowEnd' | 'merge'): SelectionData;
    /** row-major editable views for a range, e.g. "A1:C3" (clipboard copy). */
    rangeCells(sheet: number, range: string): CellEdit[][];
    patchRangeStyle(sheet: number, range: string, patch: RangeStylePatch): EditResult;
    setNumberFormat(sheet: number, range: string, format: NumberFormatMutation): EditResult;
    selectionFormatting(sheet: number, range: string): SelectionFormatting;
    captureFormat(sheet: number, range: string): CapturedFormat;
    applyFormat(sheet: number, range: string, format: CapturedFormat): EditResult;
    mergedRanges(sheet: number, range: string): MergedRange[];
    visibleMergedRanges(sheet: number, viewport: Viewport): MergedRange[];
    historyState(): HistoryState;
    /**
     * render the current sheet viewport to png bytes via the native raster
     * backend — the same display list the canvas paints, rasterized server-side.
     * throws if the embedded wasm was built without png export (the `raster`
     * cargo feature); guard with {@link isPngExportAvailable}.
     */
    renderPng(viewport: Viewport): Uint8Array;
    /**
     * render an a1 range (default: the used range) at an optional scale to png.
     * dimensions are capped wasm-side; same availability guard as renderPng.
     */
    renderRangePng(opts?: {
        range?: string;
        scale?: number;
    }): Uint8Array;
    /** serialize the workbook back to .xlsx bytes. */
    save(): Uint8Array;
    /**
     * stage an agent's proposed edits for human review without applying them.
     * returns the {@link Proposal} with per-cell formatted before/after previews.
     * throws if the embedded wasm predates proposals; guard with
     * {@link WorkbookHandle.isProposalsAvailable}.
     */
    propose(agentId: string, note: string | null, edits: ProposalEdit[]): Proposal;
    /** every pending proposal, oldest first. empty when none or unsupported. */
    listProposals(): Proposal[];
    /**
     * apply a pending proposal as one undo step and recalc dependents. throws
     * {@link StaleProposalError} when the base changed since it was staged and
     * `force` is not set; pass `{ force: true }` to apply anyway.
     */
    acceptProposal(id: string, opts?: {
        force?: boolean;
    }): EditResult & {
        proposalId: string;
    };
    /** drop a pending proposal without applying it; false when the id was unknown. */
    rejectProposal(id: string): boolean;
    /** whether the embedded wasm core was built with the proposals api. */
    isProposalsAvailable(): boolean;
    /**
     * The session-scoped version of the committed workbook. Committed edits, peer updates, undo,
     * redo and a recalculation that changes values or what an export reports about results move
     * it; selection, the active sheet and proposals do not.
     */
    version(): string;
    listDefinedNames(request?: XlsxDefinedNamesRequest): XlsxDefinedNamesResult;
    /** Cells with the version they were read at; empty `ranges` reads the sheet catalog. */
    readCells(request: XlsxReadRequest): XlsxReadResult;
    /** Exact, case-sensitive search over display text, one match per cell. */
    findText(request: XlsxFindRequest): XlsxFindResult;
    /** Resolves, stages and rehearses a batch like `applyEdits`, changing nothing. */
    validateEdits(request: XlsxEditRequest): XlsxValidationResult;
    /**
     * Applies every step as one committed, recalculated change, or refuses with nothing changed.
     * Update listeners run once, after the call returns. Malformed requests throw.
     */
    applyEdits(request: XlsxEditRequest): XlsxEditResult;
    /**
     * Exports the committed workbook with the version it was read at: sparse cells with values,
     * formulas and display text, plus sheet metadata and diagnostics. Nothing is recalculated,
     * flushed or published. Options it cannot honor refuse; malformed ones throw.
     */
    exportStructured(options?: XlsxExportOptions): XlsxExportResult<XlsxStructuredContent>;
    /** {@link WorkbookHandle.exportStructured} rendered as Markdown from the same read. */
    exportMarkdown(options?: XlsxExportOptions, markdownOptions?: XlsxMarkdownOptions): XlsxExportResult<XlsxMarkdownContent>;
    dispose(): void;
}
/** @experimental */
export declare function workbookDisplayListJson(handle: WorkbookHandle, viewport: Viewport, sheet?: number): string;
export type WasmInitInput = InitInput | Promise<InitInput>;
export interface WasmInitOptions {
    png?: boolean;
}
/** Initialize the workbook engine. Concurrent calls share the same attempt. */
export declare function initWasm(input?: WasmInitInput, options?: WasmInitOptions): Promise<void>;
/**
 * The allocation the engine could not get, if any. The core aborts when an
 * allocation fails, so a trap after one means the engine ran out of memory.
 */
export declare function engineAllocationFailure(): {
    bytes: number;
    memoryBytes: number;
} | null;
/** Maximum counted cells per frame, including boundary and frozen tracks; call after `initWasm`. */
export declare function getDisplayListCellLimit(): number;
/**
 * Whether this environment exposes the WebAssembly runtime required by the core.
 */
export declare function isWasmAvailable(): boolean;
export type EngineBackend = 'wasm' | 'native';
/** Which engine runs this workbook code ('native' when @betteroffice/xlsx-native is installed and matches this version on Node, Bun or Deno, else 'wasm'). */
export declare function engineBackend(): EngineBackend;
/**
 * Whether the embedded wasm core was built with png export (the `raster` cargo
 * feature). The wasm-bindgen method only exists on the class when compiled in,
 * so chrome can disable an export control instead of calling and catching.
 */
export declare function isPngExportAvailable(): boolean;
/**
 * Whether the embedded wasm core exposes the proposals api (propose / accept /
 * reject / list). The methods only exist on the class when compiled in, so the
 * UI can hide proposal chrome and degrade gracefully against an older module.
 */
export declare function isProposalsAvailable(): boolean;
/**
 * Open a workbook from raw `.xlsx` bytes, initializing the core if needed.
 * Throws an `Error` if the bytes are not a readable workbook.
 */
export declare function openWorkbook(bytes: Uint8Array, options?: OpenWorkbookOptions): WorkbookHandle;
export declare function workbookRecoveryCheckpoint(handle: WorkbookHandle, throughSeq: number, preserveHistory?: boolean): Uint8Array;
export declare function openWorkbookRecoveryCheckpoint(bytes: Uint8Array, throughSeq: number, options?: OpenWorkbookOptions): WorkbookHandle;
export declare function workbookPeerSnapshot(handle: WorkbookHandle): {
    begin(records: number, bytes: number): void;
    next(): Uint8Array | undefined;
    end(): void;
};
export declare function createWorkbookSnapshotBuilder(options: OpenWorkbookOptions): {
    push(chunk: Uint8Array): void;
    advance(records: number, bytes: number): boolean;
    finish(): WorkbookHandle;
    dispose(): void;
};
export declare function workbookPeerHydration(handle: WorkbookHandle): string;
export declare function adoptWorkbookPeerVersion(handle: WorkbookHandle, version: string): void;
export declare function openWorkbookPeer(bytes: Uint8Array, options: OpenWorkbookOptions, hydration: string): WorkbookHandle;
/**
 * Exports `.xlsx` bytes as read, initializing the core if needed. Formula results are the
 * stored ones: nothing is recalculated and no clock is read, so the same bytes and options
 * always give the same content. Unreadable bytes and unusable options reject.
 */
export declare function exportXlsxStructured(bytes: Uint8Array, options?: XlsxExportOptions): Promise<XlsxStructuredContent>;
/** {@link exportXlsxStructured} rendered as Markdown. */
export declare function exportXlsxMarkdown(bytes: Uint8Array, options?: XlsxExportOptions, markdownOptions?: XlsxMarkdownOptions): Promise<XlsxMarkdownContent>;
/** Renders structured content as Markdown; content that does not validate rejects. */
export declare function renderXlsxMarkdown(content: XlsxStructuredContent, options?: XlsxMarkdownOptions): Promise<XlsxMarkdownContent>;
/**
 * The crate version string, for asserting wasm/js parity.
 */
export declare function wasmVersion(): string;

// ---- workerSession.d.ts ----
/** @experimental */
export { openWorkbookSession } from './session/client';
/** @experimental */
export type { OpenWorkbookSessionOptions, WorkbookSession } from './session/client';
/** @experimental */
export { workbookSessionRecovery, WorkbookRecoveryReplayError } from './session/clientRecovery';
/** @experimental */
export type { WorkbookSessionRecovery } from './session/clientRecovery';
export { createWorkbookEditSession } from './session/editSession';
export { WorkbookPngUnavailableError } from './session/methods';
export type { WorkbookEditSession } from './session/editSession';
/** @experimental */
export type { WorkbookCellGeometry, WorkbookCellInputs, WorkbookFrame, WorkbookFrameOptions, WorkbookInitialPaint, WorkbookInitialPaintRequest, WorkbookSessionOpenReply, WorkbookSessionEvents, WorkbookSessionMethods, WorkbookSessionOpenOptions, WorkbookSessionState, WorkbookSheetSummary, WorkbookSheetView, } from './session/methods';
