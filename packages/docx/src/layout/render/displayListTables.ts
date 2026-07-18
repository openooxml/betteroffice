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
 * node's PM `before` position, resolved against the live doc). Primitives whose
 * position doesn't resolve to a table are skipped.
 *
 * Column-boundary x's and the right edge are intentionally NOT computed here:
 * they come from the PM doc's `columnWidths` (twips) anchored at `originX`,
 * which is exact and robust to horizontally-merged cells (colspan suppresses
 * the interior border line the display list would otherwise expose).
 *
 * Body region only. The HF caret + selection on canvas are now sourced from the
 * region-aware `hfRangeRects` / `hfCaretRects` queries (see
 * `CanvasHfSelectionOverlay`), but HF *table* resize handles remain a documented
 * gap: this walker takes `page.primitives` (body), and the overlay's
 * `tableKeyOf` / commit path resolves against the body PM doc. Closing it needs
 * (1) an optional region parameter here to walk `page.header|footer.primitives`,
 * (2) a `tableKeyOf` bound to the active HF EditorView, and (3) the same
 * nearest-viewport page pick the HF caret uses (the part paints on every page).
 * HF tables are rare, so this is deferred.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { DisplayList, DisplayPage, DisplayPrimitive, TableCellRef } from './displayList';
import { textRunRect, glyphRunRect, type GeoRect } from './displayListGeometry';

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
}

/** Resolve an in-cell doc position to its enclosing table's stable key, or null. */
export type TableKeyResolver = (docStart: number | undefined) => string | null;

export type DisplayListTableRegion =
  | { kind?: 'body'; rId?: undefined }
  | { kind: 'header' | 'footer'; rId?: string | null };

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
  pageSize: { width: number; height: number };
  tableKeyOf: TableKeyResolver;
  /** Resolve a table grid cell to its PM cell-start position. */
  cellPmPosOf: (tableKey: string, row: number, col: number) => number | null;
  region?: DisplayListTableRegion;
  edgeProximity?: number;
}

export interface DisplayListTableInsertHoverHit {
  type: 'row' | 'column';
  clientX: number;
  clientY: number;
  cellPmPos: number;
}

// default OOXML cell margin is ~0.08in ≈ 108 twips ≈ 7.5px at 96dpi; used only
// to back off the content bbox when a table exposes no border edge.
const ESTIMATED_CELL_INSET_PX = 7.5;
// how far outside a fragment's content bbox a border line may sit and still be
// claimed by it (cell padding + border width). Adjacent tables are separated by
// far more than this in normal flow.
const LINE_ASSOC_TOLERANCE_PX = 16;
// For full-cell overlays we need the far border too, not just the near edge.
// Wide cells with short text can put the right border well beyond the text bbox;
// this cap prevents unrelated tables elsewhere on the page from being claimed.
const GRID_LINE_ASSOC_MAX_DIST_PX = 320;
const TABLE_INSERT_EDGE_PROXIMITY_PX = 30;
export const DISPLAY_LIST_TABLE_INSERT_HIDE_DELAY_MS = 200;
const ROW_BUTTON_OFFSET_X = 24;
const ROW_BUTTON_OFFSET_Y = 10;
const COL_BUTTON_OFFSET_X = 10;
const COL_BUTTON_OFFSET_Y = 24;

function primitiveRect(p: DisplayPrimitive): GeoRect {
  let rect: GeoRect;
  switch (p.kind) {
    case 'text':
      rect = textRunRect(p);
      break;
    case 'glyphRun':
      rect = glyphRunRect(p);
      break;
    case 'rect':
    case 'image':
    case 'shape':
    case 'decoration':
      rect = { x: p.x, y: p.y, w: p.w, h: p.h };
      break;
    case 'line':
      rect = {
        x: Math.min(p.x1, p.x2),
        y: Math.min(p.y1, p.y2),
        w: Math.abs(p.x2 - p.x1),
        h: Math.abs(p.y2 - p.y1),
      };
      break;
  }
  const clip = primitiveClipRect(p);
  return clip ? (intersectRect(rect, clip) ?? { x: clip.x, y: clip.y, w: 0, h: 0 }) : rect;
}

function primitiveClipRect(p: DisplayPrimitive): GeoRect | null {
  const clip = p.clipGroup?.clip;
  if (!clip) return null;
  return {
    x: clip.x ?? 0,
    y: clip.y ?? 0,
    w: Math.max(0, clip.w ?? 0),
    h: Math.max(0, clip.h ?? 0),
  };
}

function intersectRect(a: GeoRect, b: GeoRect): GeoRect | null {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.w, b.x + b.w);
  const bottom = Math.min(a.y + a.h, b.y + b.h);
  if (right < x || bottom < y) return null;
  return { x, y, w: right - x, h: bottom - y };
}

interface TableIdentity {
  tableKey: string;
  docStartSample: number;
}

type TableIdentityMap = Map<string, TableIdentity>;

function primitiveTableIdentity(
  p: DisplayPrimitive,
  tableKeyOf: TableKeyResolver,
  identities: TableIdentityMap
): TableIdentity | null {
  if (p.docStart != null) {
    const tableKey = tableKeyOf(p.docStart);
    if (tableKey != null) {
      const identity = { tableKey, docStartSample: p.docStart };
      if (p.table?.tableId && !identities.has(p.table.tableId))
        identities.set(p.table.tableId, identity);
      return identity;
    }
  }
  return p.table?.tableId ? (identities.get(p.table.tableId) ?? null) : null;
}

/**
 * Exact ownership association for a border/cut line: the Rust builder stamps
 * `table.tableId` (and the owning `cell`) on table lines, which resolves to a
 * fragment accumulator through the identity map. Returns null when the line
 * predates the ownership contract so callers keep the geometric fallback.
 */
function ownedLineAccum<T extends { tableKey: string }>(
  p: DisplayPrimitive,
  identities: TableIdentityMap,
  accums: Map<string, T>
): T | null {
  const tableId = p.table?.tableId;
  if (!tableId) return null;
  const identity = identities.get(tableId);
  if (!identity) return null;
  return accums.get(identity.tableKey) ?? null;
}

interface FragmentAccum {
  tableKey: string;
  docStartSample: number;
  contentMinX: number;
  contentMinY: number;
  contentMaxX: number;
  contentMaxY: number;
  /** exact left-edge candidates (outer vertical borders) */
  exactLeftX: number[];
  exactBottomY: number[];
  exactClipLeftX: number[];
  exactClipBottomY: number[];
  rows: Map<number, { top: number; bottom: number }>;
}

/** Derive every body table fragment's handle geometry for one page. */
function fragmentsForPage(
  pageIndex: number,
  primitives: DisplayPrimitive[],
  tableKeyOf: TableKeyResolver,
  identities: TableIdentityMap
): DisplayListTableFragment[] {
  const accums = new Map<string, FragmentAccum>();

  // Pass 1: gather in-cell content primitives per table (grouped by the
  // caller's table key, since each cell-paragraph carries its own block id).
  for (const p of primitives) {
    if (p.kind === 'line') continue;
    const cell = p.cell;
    if (!cell) continue;
    const identity = primitiveTableIdentity(p, tableKeyOf, identities);
    if (!identity) continue;
    const { tableKey, docStartSample } = identity;
    const r = primitiveRect(p);
    const exactClip = primitiveClipRect(p);
    const geometry = exactClip ?? r;
    let a = accums.get(tableKey);
    if (!a) {
      a = {
        tableKey,
        docStartSample,
        contentMinX: geometry.x,
        contentMinY: geometry.y,
        contentMaxX: geometry.x + geometry.w,
        contentMaxY: geometry.y + geometry.h,
        exactLeftX: [],
        exactBottomY: [],
        exactClipLeftX: [],
        exactClipBottomY: [],
        rows: new Map(),
      };
      accums.set(tableKey, a);
    }
    if (docStartSample < a.docStartSample) a.docStartSample = docStartSample;
    a.contentMinX = Math.min(a.contentMinX, geometry.x);
    a.contentMinY = Math.min(a.contentMinY, geometry.y);
    a.contentMaxX = Math.max(a.contentMaxX, geometry.x + geometry.w);
    a.contentMaxY = Math.max(a.contentMaxY, geometry.y + geometry.h);
    if (exactClip) {
      a.exactClipLeftX.push(exactClip.x);
      a.exactClipBottomY.push(exactClip.y + exactClip.h);
    }

    // per anchor-row band (skip vmerge continuation slices — not the anchor)
    if (!cell.continuation) {
      const band = a.rows.get(cell.row);
      if (band) {
        band.top = Math.min(band.top, geometry.y);
        band.bottom = Math.max(band.bottom, geometry.y + geometry.h);
      } else {
        a.rows.set(cell.row, { top: geometry.y, bottom: geometry.y + geometry.h });
      }
    }
  }

  if (accums.size === 0) return [];

  // Pass 2: claim table-border / table-cut lines. Lines emitted with ownership
  // metadata (`table.tableId` stamped by the Rust builder) associate exactly;
  // pre-contract lists without it fall back to the geometric claim. Vertical
  // lines feed the exact left edge; horizontal lines feed the exact bottom
  // edge.
  for (const p of primitives) {
    if (p.kind !== 'line') continue;
    if (p.role !== 'table-border' && p.role !== 'table-cut') continue;
    let best: FragmentAccum | null = ownedLineAccum(p, identities, accums);
    if (!best) {
      const cx = (p.x1 + p.x2) / 2;
      const cy = (p.y1 + p.y2) / 2;
      let bestDist = Infinity;
      for (const a of accums.values()) {
        const insideX =
          cx >= a.contentMinX - LINE_ASSOC_TOLERANCE_PX &&
          cx <= a.contentMaxX + LINE_ASSOC_TOLERANCE_PX;
        const insideY =
          cy >= a.contentMinY - LINE_ASSOC_TOLERANCE_PX &&
          cy <= a.contentMaxY + LINE_ASSOC_TOLERANCE_PX;
        if (!insideX || !insideY) continue;
        const dx = Math.max(0, a.contentMinX - cx, cx - a.contentMaxX);
        const dy = Math.max(0, a.contentMinY - cy, cy - a.contentMaxY);
        const dist = dx + dy;
        if (dist < bestDist) {
          bestDist = dist;
          best = a;
        }
      }
    }
    if (!best) continue;
    const vertical = Math.abs(p.x1 - p.x2) < 0.5;
    const horizontal = Math.abs(p.y1 - p.y2) < 0.5;
    if (vertical) best.exactLeftX.push(Math.min(p.x1, p.x2));
    if (horizontal) best.exactBottomY.push(Math.max(p.y1, p.y2));
  }

  const out: DisplayListTableFragment[] = [];
  for (const a of accums.values()) {
    if (a.rows.size === 0) continue;
    const hasExactLeft = a.exactLeftX.length > 0 || a.exactClipLeftX.length > 0;
    const originX = hasExactLeft
      ? Math.min(...a.exactLeftX, ...a.exactClipLeftX)
      : a.contentMinX - ESTIMATED_CELL_INSET_PX;
    const bottomY =
      a.exactBottomY.length > 0 || a.exactClipBottomY.length > 0
        ? Math.max(...a.exactBottomY, ...a.exactClipBottomY)
        : a.contentMaxY;
    const rowBands = [...a.rows.entries()]
      .map(([rowIndex, band]) => ({ rowIndex, top: band.top, bottom: band.bottom }))
      .sort((x, y) => x.rowIndex - y.rowIndex);
    out.push({
      pageIndex,
      tableKey: a.tableKey,
      docStartSample: a.docStartSample,
      originX,
      originY: a.contentMinY,
      bottomY,
      originExact: hasExactLeft,
      rowBands,
    });
  }
  return out;
}

/**
 * All body table fragments across every page of the display list, each with the
 * geometry an adapter needs to mount canvas resize handles. `tableKeyOf` maps an
 * in-cell doc position to its enclosing table's stable key (see module docs).
 */
export function deriveDisplayListTableFragments(
  list: DisplayList,
  tableKeyOf: TableKeyResolver
): DisplayListTableFragment[] {
  const out: DisplayListTableFragment[] = [];
  const identities = buildTableIdentityMap(list, tableKeyOf, { kind: 'body' });
  for (const page of list.pages) {
    out.push(...fragmentsForPage(page.pageIndex, page.primitives, tableKeyOf, identities));
  }
  return out;
}

interface RegionPrimitives {
  kind: 'body' | 'header' | 'footer';
  rId?: string;
  primitives: DisplayPrimitive[];
}

function primitivesForRegion(
  page: DisplayPage,
  region: DisplayListTableRegion
): RegionPrimitives | null {
  const kind = region.kind ?? 'body';
  if (kind === 'body') return { kind, primitives: page.primitives };
  const hf = kind === 'header' ? page.header : page.footer;
  if (!hf || hf.kind !== kind) return null;
  if (region.rId && hf.rId !== region.rId) return null;
  return { kind, rId: hf.rId, primitives: hf.primitives };
}

function buildTableIdentityMap(
  list: DisplayList,
  tableKeyOf: TableKeyResolver,
  region: DisplayListTableRegion
): TableIdentityMap {
  const identities: TableIdentityMap = new Map();
  for (const page of list.pages) {
    const regionPrimitives = primitivesForRegion(page, region);
    if (!regionPrimitives) continue;
    for (const primitive of regionPrimitives.primitives) {
      if (!primitive.cell || !primitive.table?.tableId || primitive.docStart == null) continue;
      const tableKey = tableKeyOf(primitive.docStart);
      if (tableKey == null) continue;
      const current = identities.get(primitive.table.tableId);
      if (!current || primitive.docStart < current.docStartSample) {
        identities.set(primitive.table.tableId, { tableKey, docStartSample: primitive.docStart });
      }
    }
  }
  return identities;
}

interface CellGridAccum {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
  exactClip?: GeoRect;
}

interface TableGridAccum {
  tableKey: string;
  contentMinX: number;
  contentMinY: number;
  contentMaxX: number;
  contentMaxY: number;
  cells: Map<string, CellGridAccum>;
  verticalLines: number[];
  horizontalLines: number[];
}

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function upsertCellGrid(
  a: TableGridAccum,
  cell: TableCellRef,
  r: GeoRect,
  exactClip: GeoRect | null
): void {
  const key = cellKey(cell.row, cell.col);
  const existing = a.cells.get(key);
  if (existing) {
    existing.minX = Math.min(existing.minX, r.x);
    existing.minY = Math.min(existing.minY, r.y);
    existing.maxX = Math.max(existing.maxX, r.x + r.w);
    existing.maxY = Math.max(existing.maxY, r.y + r.h);
    existing.rowSpan = Math.max(existing.rowSpan, cell.rowSpan);
    existing.colSpan = Math.max(existing.colSpan, cell.colSpan);
    if (exactClip) existing.exactClip = exactClip;
  } else {
    a.cells.set(key, {
      row: cell.row,
      col: cell.col,
      rowSpan: cell.rowSpan,
      colSpan: cell.colSpan,
      minX: r.x,
      minY: r.y,
      maxX: r.x + r.w,
      maxY: r.y + r.h,
      exactClip: exactClip ?? undefined,
    });
  }
}

function deriveTableGridsForPrimitives(
  primitives: DisplayPrimitive[],
  tableKeyOf: TableKeyResolver,
  identities: TableIdentityMap
): TableGridAccum[] {
  const accums = new Map<string, TableGridAccum>();

  for (const p of primitives) {
    if (p.kind === 'line') continue;
    const cell = p.cell;
    if (!cell) continue;
    const identity = primitiveTableIdentity(p, tableKeyOf, identities);
    if (!identity) continue;
    const tableKey = identity.tableKey;
    const r = primitiveRect(p);
    let a = accums.get(tableKey);
    if (!a) {
      a = {
        tableKey,
        contentMinX: r.x,
        contentMinY: r.y,
        contentMaxX: r.x + r.w,
        contentMaxY: r.y + r.h,
        cells: new Map(),
        verticalLines: [],
        horizontalLines: [],
      };
      accums.set(tableKey, a);
    }
    a.contentMinX = Math.min(a.contentMinX, r.x);
    a.contentMinY = Math.min(a.contentMinY, r.y);
    a.contentMaxX = Math.max(a.contentMaxX, r.x + r.w);
    a.contentMaxY = Math.max(a.contentMaxY, r.y + r.h);
    upsertCellGrid(a, cell, r, primitiveClipRect(p));
  }

  if (accums.size === 0) return [];

  for (const p of primitives) {
    if (p.kind !== 'line') continue;
    if (p.role !== 'table-border' && p.role !== 'table-cut') continue;

    const vertical = Math.abs(p.x1 - p.x2) < 0.5;
    const horizontal = Math.abs(p.y1 - p.y2) < 0.5;
    if (!vertical && !horizontal) continue;

    const lineMinX = Math.min(p.x1, p.x2);
    const lineMaxX = Math.max(p.x1, p.x2);
    const lineMinY = Math.min(p.y1, p.y2);
    const lineMaxY = Math.max(p.y1, p.y2);

    // ownership metadata associates the line exactly; geometric distance is
    // only the pre-contract fallback
    const owned = ownedLineAccum(p, identities, accums);
    if (owned) {
      if (vertical) owned.verticalLines.push(lineMinX);
      else owned.horizontalLines.push(lineMinY);
      continue;
    }

    for (const a of accums.values()) {
      if (vertical) {
        const yOverlaps =
          lineMaxY >= a.contentMinY - LINE_ASSOC_TOLERANCE_PX &&
          lineMinY <= a.contentMaxY + LINE_ASSOC_TOLERANCE_PX;
        if (!yOverlaps) continue;
        const x = lineMinX;
        const dx = Math.max(0, a.contentMinX - x, x - a.contentMaxX);
        if (dx <= GRID_LINE_ASSOC_MAX_DIST_PX) a.verticalLines.push(x);
      } else if (horizontal) {
        const xOverlaps =
          lineMaxX >= a.contentMinX - LINE_ASSOC_TOLERANCE_PX &&
          lineMinX <= a.contentMaxX + LINE_ASSOC_TOLERANCE_PX;
        if (!xOverlaps) continue;
        const y = lineMinY;
        const dy = Math.max(0, a.contentMinY - y, y - a.contentMaxY);
        if (dy <= GRID_LINE_ASSOC_MAX_DIST_PX) a.horizontalLines.push(y);
      }
    }
  }

  return [...accums.values()];
}

function sortedUnique(values: number[]): number[] {
  const out: number[] = [];
  for (const v of [...values].sort((a, b) => a - b)) {
    if (out.length === 0 || Math.abs(out[out.length - 1] - v) > 0.5) out.push(v);
  }
  return out;
}

function rowBoundaryMap(table: TableGridAccum): Map<number, number> {
  const rows = sortedUnique([...table.cells.values()].map((c) => c.row));
  const ys = sortedUnique(table.horizontalLines);
  const map = new Map<number, number>();
  if (rows.length > 0 && ys.length >= rows.length + 1) {
    for (let i = 0; i < rows.length; i++) {
      map.set(rows[i], ys[i]);
      map.set(rows[i] + 1, ys[i + 1]);
    }
  }
  return map;
}

function fallbackCellRect(cell: CellGridAccum | undefined): {
  x: number;
  y: number;
  w: number;
  h: number;
} | null {
  if (!cell) return null;
  if (cell.exactClip) return cell.exactClip;
  const x = cell.minX - ESTIMATED_CELL_INSET_PX;
  const y = cell.minY - ESTIMATED_CELL_INSET_PX;
  return {
    x,
    y,
    w: Math.max(1, cell.maxX - cell.minX + ESTIMATED_CELL_INSET_PX * 2),
    h: Math.max(1, cell.maxY - cell.minY + ESTIMATED_CELL_INSET_PX * 2),
  };
}

function rowBoundsFallback(
  table: TableGridAccum,
  row: number
): { top: number; bottom: number } | null {
  let top = Infinity;
  let bottom = -Infinity;
  for (const cell of table.cells.values()) {
    if (cell.row !== row) continue;
    top = Math.min(top, cell.exactClip?.y ?? cell.minY - ESTIMATED_CELL_INSET_PX);
    bottom = Math.max(
      bottom,
      cell.exactClip ? cell.exactClip.y + cell.exactClip.h : cell.maxY + ESTIMATED_CELL_INSET_PX
    );
  }
  if (!Number.isFinite(top) || !Number.isFinite(bottom) || bottom <= top) return null;
  return { top, bottom };
}

function cellBoundsFallback(
  table: TableGridAccum,
  row: number,
  col: number
): { left: number; right: number; top: number; bottom: number } | null {
  const fallback = fallbackCellRect(table.cells.get(cellKey(row, col)));
  if (!fallback) return null;
  return {
    left: fallback.x,
    right: fallback.x + fallback.w,
    top: fallback.y,
    bottom: fallback.y + fallback.h,
  };
}

/**
 * Full selected-cell boxes for canvas overlays.
 *
 * The DOM painter highlights `.layout-table-cell` elements directly; in canvas
 * mode those cells are parked, so adapters pass the live PM CellSelection as
 * grid coordinates and this helper resolves visible page-local boxes from the
 * display-list table geometry. Border lines provide exact grid edges when
 * present; borderless fragments fall back to the selected cell's content bbox
 * padded by the same conservative inset used by table-resize geometry.
 */
export function deriveDisplayListSelectedCellRects(
  list: DisplayList,
  selectedCells: DisplayListSelectedCellSpec[],
  tableKeyOf: TableKeyResolver,
  region: DisplayListTableRegion = { kind: 'body' }
): DisplayListSelectedCellRect[] {
  if (selectedCells.length === 0) return [];
  const selectedByTable = new Map<string, DisplayListSelectedCellSpec[]>();
  for (const cell of selectedCells) {
    const arr = selectedByTable.get(cell.tableKey);
    if (arr) arr.push(cell);
    else selectedByTable.set(cell.tableKey, [cell]);
  }

  const out: DisplayListSelectedCellRect[] = [];
  const identities = buildTableIdentityMap(list, tableKeyOf, region);
  for (const page of list.pages) {
    const regionPrims = primitivesForRegion(page, region);
    if (!regionPrims) continue;
    const tables = deriveTableGridsForPrimitives(regionPrims.primitives, tableKeyOf, identities);
    for (const table of tables) {
      const selected = selectedByTable.get(table.tableKey);
      if (!selected) continue;
      const xs = sortedUnique(table.verticalLines);
      const rowYs = rowBoundaryMap(table);

      for (const spec of selected) {
        const contentCell = table.cells.get(cellKey(spec.row, spec.col));
        let x = xs[spec.col];
        let right = xs[spec.col + spec.colSpan];
        let y = rowYs.get(spec.row);
        let bottom = rowYs.get(spec.row + spec.rowSpan);
        let exact = true;

        if (x == null || right == null || right <= x) {
          const fallback = contentCell?.exactClip ?? fallbackCellRect(contentCell);
          if (!fallback) continue;
          x = fallback.x;
          right = fallback.x + fallback.w;
          exact = contentCell?.exactClip != null;
        }
        if (y == null || bottom == null || bottom <= y) {
          const fallback = contentCell?.exactClip ?? fallbackCellRect(contentCell);
          if (!fallback) continue;
          y = fallback.y;
          bottom = fallback.y + fallback.h;
          exact &&= contentCell?.exactClip != null;
        }

        out.push({
          pageIndex: page.pageIndex,
          tableKey: table.tableKey,
          row: spec.row,
          col: spec.col,
          rowSpan: spec.rowSpan,
          colSpan: spec.colSpan,
          x,
          y,
          w: Math.max(1, right - x),
          h: Math.max(1, bottom - y),
          exact,
          region: regionPrims.kind,
          rId: regionPrims.rId,
        });
      }
    }
  }
  return out;
}

/**
 * Canvas/display-list version of the table insert "+" hover detector.
 *
 * Mirrors `detectTableInsertHover`'s edge threshold and button offsets, but
 * reads table/cell boxes from display-list geometry because the painter DOM is
 * parked while canvas is active.
 */
export function detectDisplayListTableInsertHover(
  input: DisplayListTableInsertHoverInput
): DisplayListTableInsertHoverHit | null {
  const {
    list,
    pageIndex,
    x,
    y,
    canvasRect,
    pageSize,
    tableKeyOf,
    cellPmPosOf,
    region = { kind: 'body' },
    edgeProximity = TABLE_INSERT_EDGE_PROXIMITY_PX,
  } = input;

  const page = list.pages[pageIndex];
  if (!page) return null;
  const regionPrims = primitivesForRegion(page, region);
  if (!regionPrims) return null;

  const scaleX = pageSize.width > 0 ? canvasRect.width / pageSize.width : 1;
  const scaleY = pageSize.height > 0 ? canvasRect.height / pageSize.height : 1;
  const toClient = (localX: number, localY: number) => ({
    clientX: canvasRect.left + localX * scaleX,
    clientY: canvasRect.top + localY * scaleY,
  });

  const identities = buildTableIdentityMap(list, tableKeyOf, region);
  const tables = deriveTableGridsForPrimitives(regionPrims.primitives, tableKeyOf, identities);
  for (const table of tables) {
    const xs = sortedUnique(table.verticalLines);
    const rowYs = rowBoundaryMap(table);
    const rows = sortedUnique([...table.cells.values()].map((cell) => cell.row));
    if (rows.length === 0) continue;

    const left = xs[0] ?? table.contentMinX - ESTIMATED_CELL_INSET_PX;
    const right = xs[xs.length - 1] ?? table.contentMaxX + ESTIMATED_CELL_INSET_PX;
    const top = rowYs.get(rows[0]) ?? table.contentMinY - ESTIMATED_CELL_INSET_PX;
    const lastRow = rows[rows.length - 1];
    const bottom = rowYs.get(lastRow + 1) ?? table.contentMaxY + ESTIMATED_CELL_INSET_PX;

    const nearLeftEdge = x < left + edgeProximity && x >= left - edgeProximity;
    const nearTopEdge = y < top + edgeProximity && y >= top - edgeProximity;
    if (!nearLeftEdge && !nearTopEdge) continue;

    if (nearLeftEdge && y >= top - edgeProximity && y <= bottom) {
      for (const row of rows) {
        const exactTop = rowYs.get(row);
        const exactBottom = rowYs.get(row + 1);
        const fallback =
          exactTop == null || exactBottom == null ? rowBoundsFallback(table, row) : null;
        const rowTop = exactTop ?? fallback?.top;
        const rowBottom = exactBottom ?? fallback?.bottom;
        if (rowTop == null || rowBottom == null || rowBottom <= rowTop) continue;
        if (y < rowTop || y > rowBottom) continue;

        const firstCell = [...table.cells.values()]
          .filter((cell) => cell.row === row)
          .sort((a, b) => a.col - b.col)[0];
        if (!firstCell) continue;
        const pmPos = cellPmPosOf(table.tableKey, row, firstCell.col);
        if (!pmPos) continue;
        const anchor = toClient(left, rowTop + (rowBottom - rowTop) / 2);
        return {
          type: 'row',
          clientX: anchor.clientX - ROW_BUTTON_OFFSET_X,
          clientY: anchor.clientY - ROW_BUTTON_OFFSET_Y,
          cellPmPos: pmPos,
        };
      }
    }

    if (nearTopEdge && x >= left - edgeProximity && x <= right) {
      const firstRow = rows[0];
      const cells = [...table.cells.values()]
        .filter((cell) => cell.row === firstRow)
        .sort((a, b) => a.col - b.col);
      for (const cell of cells) {
        const exactLeft = xs[cell.col];
        const exactRight = xs[cell.col + cell.colSpan];
        const fallback =
          exactLeft == null || exactRight == null
            ? cellBoundsFallback(table, cell.row, cell.col)
            : null;
        const cellLeft = exactLeft ?? fallback?.left;
        const cellRight = exactRight ?? fallback?.right;
        if (cellLeft == null || cellRight == null || cellRight <= cellLeft) continue;
        if (x < cellLeft || x > cellRight) continue;

        const pmPos = cellPmPosOf(table.tableKey, cell.row, cell.col);
        if (!pmPos) continue;
        const anchor = toClient(cellLeft + (cellRight - cellLeft) / 2, top);
        return {
          type: 'column',
          clientX: anchor.clientX - COL_BUTTON_OFFSET_X,
          clientY: anchor.clientY - COL_BUTTON_OFFSET_Y,
          cellPmPos: pmPos,
        };
      }
    }
  }

  return null;
}
