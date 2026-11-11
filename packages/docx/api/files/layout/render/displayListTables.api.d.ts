// @betteroffice/docx declaration layout/render/displayListTables.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Derive table-resize handle geometry from a built DisplayList.
 *
 * On the canvas path the DOM painter's `.layout-table-*-handle` elements are
 * parked/hidden, so the canvas table-resize overlay must place its column / row
 * / edge handles from the display list instead. This module walks the list and
 * returns, per painted table fragment, the geometry an adapter needs to mount
 * handles over the visible canvas:
 *
 *  - `originX` / `originY` — the fragment's page-local top-left. `originX` is
 *    taken from an EXACT edge source (an outer vertical table-border line)
 *    whenever one exists, so column handles land on the true column boundary
 *    rather than inset by cell padding. Only a fully borderless table falls back
 *    to the content bbox minus an estimated inset (documented, approximate).
 *  - `bottomY` — the fragment's bottom edge (exact when a bottom border exists,
 *    else content bbox bottom).
 *  - `rowBands` — per anchor-row content band `{ rowIndex, top, bottom }`,
 *    grouped by the `cell.row` carried on every in-cell primitive. Row-boundary
 *    handles sit in the gap between consecutive bands, and the `rowIndex` maps
 *    straight to `commitRowResize` (robust to vmerge, which suppresses interior
 *    border lines but keeps the anchor row on every primitive).
 *  - `tableKey` — the caller-supplied identity of the enclosing table; also the
 *    handle key the adapter resolves back to the table node.
 *
 * Grouping key: cell content primitives carry each cell-paragraph's block id,
 * NOT the table's, so grouping by block id would split a table into one fragment
 * per cell. The adapter therefore supplies `tableKeyOf`, mapping an in-cell doc
 * position to a stable per-table key (in the React adapter: the enclosing table
 * node's `before` position, resolved against the live doc). Primitives whose
 * position doesn't resolve to a table are skipped.
 *
 * Column-boundary x's and the right edge are intentionally NOT computed here:
 * they come from the doc's `columnWidths` (twips) anchored at `originX`,
 * which is exact and robust to horizontally-merged cells (colspan suppresses
 * the interior border line the display list would otherwise expose).
 *
 * Body region only. The HF caret + selection on canvas are now sourced from the
 * region-aware `hfRangeRects` / `hfCaretRects` queries (see
 * `CanvasPartSelectionOverlay`), but HF *table* resize handles remain a documented
 * gap: this walker takes `page.primitives` (body), and the overlay's
 * `tableKeyOf` / commit path resolves against the body doc. Closing it needs
 * (1) an optional region parameter here to walk `page.header|footer.primitives`,
 * (2) a `tableKeyOf` bound to the active HF EditorView, and (3) the same
 * nearest-viewport page pick the HF caret uses (the part paints on every page).
 * HF tables are rare, so this is deferred.
 *
 */
import type { DisplayList } from './displayList';
export interface TableRowBand {
    /** 0-based anchor-row index (from `cell.row`) */
    rowIndex: number;
    top: number;
    bottom: number;
}
export interface DisplayListTableFragment {
    pageIndex: number;
    /** caller-supplied identity of the enclosing table (also the handle key) */
    tableKey: string;
    /** smallest in-cell `docStart` seen for this table on this page */
    docStartSample: number;
    /** page-local left edge (exact when a border exposes it, else approximate) */
    originX: number;
    /** page-local top edge of the fragment's content */
    originY: number;
    /** page-local bottom edge of the fragment's content */
    bottomY: number;
    /** true when `originX` came from an exact edge source (border line) */
    originExact: boolean;
    /** per anchor-row content bands, ascending by `rowIndex` */
    rowBands: TableRowBand[];
    /**
     * One past the last source row the table paints (`table.rowCount`), when the
     * display list stamps it. Rows hidden from view are absent from `rowBands`.
     */
    rowCount?: number;
}
/** Resolve an in-cell doc position to its enclosing table's stable key, or null. */
export type TableKeyResolver = (docStart: number | undefined) => string | null;
export type DisplayListTableRegion = {
    kind?: 'body';
    rId?: undefined;
} | {
    kind: 'header' | 'footer';
    rId?: string | null;
};
export interface DisplayListSelectedCellSpec {
    /** caller-supplied identity of the enclosing table */
    tableKey: string;
    row: number;
    col: number;
    rowSpan: number;
    colSpan: number;
}
export interface DisplayListSelectedCellRect {
    pageIndex: number;
    tableKey: string;
    row: number;
    col: number;
    rowSpan: number;
    colSpan: number;
    x: number;
    y: number;
    w: number;
    h: number;
    exact: boolean;
    /** This box is a header row's copy on a continuation page. */
    repeatedHeader?: boolean;
    region: 'body' | 'header' | 'footer';
    rId?: string;
}
export interface DisplayListTableInsertHoverInput {
    list: DisplayList;
    pageIndex: number;
    /** Page-local pointer x/y in display-list px. */
    x: number;
    y: number;
    /** Client-space canvas rect for `pageIndex`, used to return DOM-path anchors. */
    canvasRect: Pick<DOMRect, 'left' | 'top' | 'width' | 'height'>;
    /** Display-list page size for `pageIndex`. */
    pageSize: {
        width: number;
        height: number;
    };
    tableKeyOf: TableKeyResolver;
    /** Resolve a table grid cell to its cell-start position. */
    cellPmPosOf: (tableKey: string, row: number, col: number) => number | null;
    region?: DisplayListTableRegion;
    edgeProximity?: number;
    /** Client px per CSS px of the button, which scales its offsets from the table. Default 1. */
    buttonZoom?: number;
}
export interface DisplayListTableInsertHoverHit {
    type: 'row' | 'column';
    clientX: number;
    clientY: number;
    cellPmPos: number;
}
export declare const DISPLAY_LIST_TABLE_INSERT_HIDE_DELAY_MS = 200;
/**
 * Whether the fragment paints the table's last row, so edge handles belong on
 * it. `modelRowCount` is the table's row count in the document; the display
 * list's own count wins, since it leaves out trailing rows hidden from view.
 */
export declare function displayListTableFragmentEndsTable(fragment: Pick<DisplayListTableFragment, 'rowBands' | 'rowCount'>, modelRowCount: number): boolean;
/**
 * All body table fragments across every page of the display list, each with the
 * geometry an adapter needs to mount canvas resize handles. `tableKeyOf` maps an
 * in-cell doc position to its enclosing table's stable key (see module docs).
 */
export declare function deriveDisplayListTableFragments(list: DisplayList, tableKeyOf: TableKeyResolver): DisplayListTableFragment[];
/**
 * [`deriveDisplayListTableFragments`] for the pages `start..=end` only. Each of
 * their tables still takes its identity from every page it covers.
 */
export declare function deriveDisplayListTableFragmentsOnPages(list: DisplayList, tableKeyOf: TableKeyResolver, start: number, end: number): DisplayListTableFragment[];
/**
 * Full selected-cell boxes for canvas overlays.
 *
 * The DOM painter highlights `.layout-table-cell` elements directly; in canvas
 * mode those cells are parked, so adapters pass the live cell selection as
 * grid coordinates and this helper resolves visible page-local boxes from the
 * display-list table geometry. Border lines provide exact grid edges when
 * present; borderless fragments fall back to the selected cell's content bbox
 * padded by the same conservative inset used by table-resize geometry.
 */
export declare function deriveDisplayListSelectedCellRects(list: DisplayList, selectedCells: DisplayListSelectedCellSpec[], tableKeyOf: TableKeyResolver, region?: DisplayListTableRegion): DisplayListSelectedCellRect[];
/**
 * Canvas/display-list version of the table insert "+" hover detector.
 *
 * Mirrors `detectTableInsertHover`'s edge threshold and button offsets, but
 * reads table/cell boxes from display-list geometry because the painter DOM is
 * parked while canvas is active.
 */
export declare function detectDisplayListTableInsertHover(input: DisplayListTableInsertHoverInput): DisplayListTableInsertHoverHit | null;
