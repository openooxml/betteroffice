import { tableRowEdges, type TableCellBox, type TableRowEdge } from '../internals/tableRowEdges';
import type { PeerlessPropertyIntent } from './inputEvents';
import type { PeerlessFrameGeometry } from './viewStore';

/** Twips per layout pixel. */
const TWIPS_PER_PX = 15;
const MIN_COLUMN_PX = 20;
const MIN_ROW_PX = 1;

/** @internal A draggable border of one table fragment on one page, in page-local layout px. */
export interface TableResizeHandle {
  id: string;
  kind: 'col' | 'right';
  table: string;
  pageIndex: number;
  /** Grid boundary the border sits on: the right edge of column `columnEdge - 1`. */
  columnEdge: number;
  x: number;
  top: number;
  bottom: number;
  /** Narrowest drag in px before a neighbouring column reaches its minimum width. */
  minDelta: number;
  /** Widest drag in px, or Infinity for the right edge. */
  maxDelta: number;
  /** A cell of the table: the worker resolves the table from it, and its handle outlives other tables moving. */
  cell: string;
  /** Grid column of `cell` and columns of its table when the handle was built. */
  cellColumn: number;
  columns: number;
}

function isBodyStory(story: string): boolean {
  return story === 'body' || story.startsWith('body:');
}

function presentedObjects(geometry: PeerlessFrameGeometry | null | undefined) {
  const objects = geometry?.objects;
  return geometry && objects && objects.version === geometry.documentVersion &&
    objects.frameEpoch === geometry.frameEpoch && objects.layoutRevision === geometry.layoutRevision ? objects : null;
}

/**
 * Column borders and right edges of the body tables on `pageIndices`, from the
 * cell boxes the worker's frame geometry carries. Empty while that geometry
 * belongs to another frame or document version than the pages shown.
 */
export function tableResizeHandles(geometry: PeerlessFrameGeometry | null | undefined, pageIndices: ReadonlySet<number>): TableResizeHandle[] {
  const objects = presentedObjects(geometry);
  if (!objects) return [];
  const fragments = new Map<string, { table: string; pageIndex: number; columns: number; cell: string; cellColumn: number;
    gridTwips?: readonly number[]; edges: Map<number, number[]>; top: number; bottom: number }>();
  for (const object of objects.objects) {
    const placement = object.cell;
    if (object.kind !== 'cell' || !placement || !isBodyStory(object.story)) continue;
    for (const rect of object.rects) {
      if (!pageIndices.has(rect.pageIndex) || !(rect.width > 0) || !(rect.height > 0)) continue;
      const key = `${placement.table}@${rect.pageIndex}`;
      let fragment = fragments.get(key);
      if (!fragment) {
        fragment = { table: placement.table, pageIndex: rect.pageIndex, columns: placement.columns, cell: object.handle,
          cellColumn: placement.column, edges: new Map(), top: rect.y, bottom: rect.y + rect.height };
        fragments.set(key, fragment);
      }
      fragment.gridTwips ??= placement.gridTwips;
      fragment.top = Math.min(fragment.top, rect.y);
      fragment.bottom = Math.max(fragment.bottom, rect.y + rect.height);
      for (const [edge, x] of [[placement.column, rect.x], [placement.column + placement.colspan, rect.x + rect.width]] as const) {
        const seen = fragment.edges.get(edge);
        if (seen) seen.push(x);
        else fragment.edges.set(edge, [x]);
      }
    }
  }
  const handles: TableResizeHandle[] = [];
  for (const fragment of fragments.values()) {
    const origin = fragment.edges.get(0);
    if (fragment.gridTwips?.length === fragment.columns && origin) {
      let x = Math.min(...origin);
      fragment.edges = new Map([[0, [x]]]);
      fragment.gridTwips.forEach((width, column) => {
        x += width / TWIPS_PER_PX;
        fragment.edges.set(column + 1, [x]);
      });
    }
    const at = (edge: number) => {
      const seen = fragment.edges.get(edge);
      return seen ? seen.reduce((sum, x) => sum + x, 0) / seen.length : undefined;
    };
    const known = [...fragment.edges.keys()].sort((a, b) => a - b);
    for (let edge = 1; edge <= fragment.columns; edge += 1) {
      const x = at(edge);
      const before = known.filter((candidate) => candidate < edge).pop();
      const after = edge < fragment.columns ? known.find((candidate) => candidate > edge) : undefined;
      if (x === undefined || before === undefined || (edge < fragment.columns && after === undefined)) continue;
      const previous = at(before)!;
      const next = after === undefined ? undefined : at(after)!;
      const kind = edge === fragment.columns ? 'right' : 'col';
      handles.push({
        id: `${fragment.table}@${fragment.pageIndex}:${edge}`, kind, table: fragment.table, pageIndex: fragment.pageIndex,
        columnEdge: edge, x, top: fragment.top, bottom: fragment.bottom, cell: fragment.cell,
        cellColumn: fragment.cellColumn, columns: fragment.columns,
        minDelta: Math.min(0, MIN_COLUMN_PX - (x - previous)),
        maxDelta: next === undefined ? Infinity : Math.max(0, next - x - MIN_COLUMN_PX),
      });
    }
  }
  return handles;
}

/** @internal The handle whose border is nearest page-local `x`, `y` on `pageIndex`, within `tolerance` px. */
export function tableHandleAt(handles: readonly TableResizeHandle[], pageIndex: number, x: number, y: number, tolerance: number): TableResizeHandle | null {
  let nearest: TableResizeHandle | null = null;
  for (const handle of handles) {
    if (handle.pageIndex !== pageIndex || y < handle.top || y > handle.bottom) continue;
    const distance = Math.abs(x - handle.x);
    if (distance <= tolerance && (!nearest || distance < Math.abs(x - nearest.x))) nearest = handle;
  }
  return nearest;
}

/**
 * @internal The worker command for dragging `handle` by `deltaPx` layout px, or null for no change.
 * Null too once the table the drag began on is gone or its columns changed: a table's index is
 * not its identity, so another table moving into that index must not take the drag.
 */
export function tableResizeIntent(geometry: PeerlessFrameGeometry | null | undefined, handle: TableResizeHandle, deltaPx: number): PeerlessPropertyIntent | null {
  const objects = presentedObjects(geometry);
  const placement = objects?.objects.find((object) => object.handle === handle.cell)?.cell;
  const deltaTwips = Math.round(Math.min(Math.max(deltaPx, handle.minDelta), handle.maxDelta) * TWIPS_PER_PX);
  if (!objects || !placement || placement.column !== handle.cellColumn || placement.columns !== handle.columns || deltaTwips === 0) return null;
  return { kind: 'resizeObject', payload: { handle: handle.cell, expectVersion: objects.version,
    columnEdge: handle.columnEdge, deltaTwips } };
}

/** @internal A draggable row bottom border of one table, in page-local layout px. */
export interface TableRowHandle extends TableRowEdge<string> {
  id: string;
  minDelta: number;
}

/** Row bottom borders of the body tables ending on `pageIndices`, bound to a cell of each row. */
export function tableRowResizeHandles(geometry: PeerlessFrameGeometry | null | undefined, pageIndices: ReadonlySet<number>): TableRowHandle[] {
  const objects = presentedObjects(geometry);
  if (!objects) return [];
  const tables = new Map<string, { row: number; rowspan: number; boxes: TableCellBox[]; cell: string }[]>();
  for (const object of objects.objects) {
    const placement = object.cell;
    if (object.kind !== 'cell' || !placement?.exact || !isBodyStory(object.story)) continue;
    const cells = tables.get(placement.table) ?? [];
    cells.push({ row: placement.row, rowspan: placement.rowspan, boxes: object.rects, cell: object.handle });
    tables.set(placement.table, cells);
  }
  return [...tables].flatMap(([table, cells]) => tableRowEdges(cells, pageIndices).map((edge) => ({
    ...edge, id: `${table}@${edge.pageIndex}:row${edge.row}`, minDelta: Math.min(0, MIN_ROW_PX - edge.paintedHeight),
  })));
}

/** @internal The row border nearest page-local `x`, `y` on `pageIndex`, within `tolerance` px. */
export function tableRowHandleAt(handles: readonly TableRowHandle[], pageIndex: number, x: number, y: number, tolerance: number): TableRowHandle | null {
  let nearest: TableRowHandle | null = null;
  for (const handle of handles) {
    if (handle.pageIndex !== pageIndex || x < handle.left || x > handle.right) continue;
    const distance = Math.abs(y - handle.y);
    if (distance <= tolerance && (!nearest || distance < Math.abs(y - nearest.y))) nearest = handle;
  }
  return nearest;
}

/** @internal The worker command for dragging `handle` by `deltaPx` layout px, or null once its cell is gone or merged. */
export function tableRowResizeIntent(geometry: PeerlessFrameGeometry | null | undefined, handle: TableRowHandle, deltaPx: number): PeerlessPropertyIntent | null {
  const objects = presentedObjects(geometry);
  const placement = objects?.objects.find((object) => object.handle === handle.cell)?.cell;
  const rowDeltaTwips = Math.round(Math.max(deltaPx, handle.minDelta) * TWIPS_PER_PX);
  if (!objects || placement?.rowspan !== 1 || rowDeltaTwips === 0) return null;
  return { kind: 'resizeObject', payload: { handle: handle.cell, expectVersion: objects.version,
    paintedRowTwips: Math.round(handle.paintedHeight * TWIPS_PER_PX), rowDeltaTwips } };
}
