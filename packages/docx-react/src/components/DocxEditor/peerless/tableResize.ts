import type { PeerlessPropertyIntent } from './inputEvents';
import type { PeerlessFrameGeometry } from './viewStore';

/** Twips per layout pixel. */
const TWIPS_PER_PX = 15;
const MIN_COLUMN_PX = 20;

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
  /** Any cell of the table: the worker resolves the table from it. */
  cell: string;
}

function isBodyStory(story: string): boolean {
  return story === 'body' || story.startsWith('body:');
}

/**
 * Column borders and right edges of the body tables on `pageIndices`, from the
 * cell boxes the worker's frame geometry carries. Empty while that geometry
 * belongs to another frame or document version than the pages shown.
 */
export function tableResizeHandles(geometry: PeerlessFrameGeometry | null | undefined, pageIndices: ReadonlySet<number>): TableResizeHandle[] {
  const objects = geometry?.objects;
  if (!geometry || !objects || objects.version !== geometry.documentVersion ||
    objects.frameEpoch !== geometry.frameEpoch || objects.layoutRevision !== geometry.layoutRevision) return [];
  const fragments = new Map<string, { table: string; pageIndex: number; columns: number; cell: string;
    edges: Map<number, number[]>; top: number; bottom: number }>();
  for (const object of objects.objects) {
    const placement = object.cell;
    if (object.kind !== 'cell' || !placement || !isBodyStory(object.story)) continue;
    for (const rect of object.rects) {
      if (!pageIndices.has(rect.pageIndex) || !(rect.width > 0) || !(rect.height > 0)) continue;
      const key = `${placement.table}@${rect.pageIndex}`;
      let fragment = fragments.get(key);
      if (!fragment) {
        fragment = { table: placement.table, pageIndex: rect.pageIndex, columns: placement.columns, cell: object.handle,
          edges: new Map(), top: rect.y, bottom: rect.y + rect.height };
        fragments.set(key, fragment);
      }
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
        minDelta: Math.min(0, MIN_COLUMN_PX - (x - previous)),
        maxDelta: next === undefined ? Infinity : Math.max(0, next - x - MIN_COLUMN_PX),
      });
    }
  }
  return handles;
}

/** @internal The worker command for dragging `handle` by `deltaPx` layout px, or null for no change. */
export function tableResizeIntent(geometry: PeerlessFrameGeometry | null | undefined, handle: TableResizeHandle, deltaPx: number): PeerlessPropertyIntent | null {
  const current = tableResizeHandles(geometry, new Set([handle.pageIndex]))
    .find((candidate) => candidate.table === handle.table && candidate.columnEdge === handle.columnEdge);
  const deltaTwips = Math.round(Math.min(Math.max(deltaPx, handle.minDelta), handle.maxDelta) * TWIPS_PER_PX);
  if (!current || !geometry?.objects || deltaTwips === 0) return null;
  return { kind: 'resizeObject', payload: { handle: current.cell, expectVersion: geometry.objects.version,
    columnEdge: handle.columnEdge, deltaTwips } };
}
