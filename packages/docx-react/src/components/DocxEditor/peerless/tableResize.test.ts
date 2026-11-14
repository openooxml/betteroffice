import { expect, test } from 'bun:test';
import { tableHandleAt, tableResizeHandles, tableResizeIntent, tableRowHandleAt, tableRowResizeHandles, tableRowResizeIntent } from './tableResize';
import type { PeerlessFrameGeometry } from './viewStore';

type Cell = { story: string; row: number; column: number; colspan?: number; rects: { pageIndex: number; x: number; y: number; width: number; height: number }[] };

function geometry(cells: Cell[], overrides: { version?: string; columns?: number; table?: string; handles?: string; exact?: boolean } = {}): PeerlessFrameGeometry {
  return {
    documentVersion: 'v1', frameEpoch: 3, layoutRevision: 7,
    objects: {
      version: overrides.version ?? 'v1', frameEpoch: 3, layoutRevision: 7, selection: [], unbuiltPages: [],
      objects: cells.map((cell, index) => ({
        kind: 'cell' as const, handle: `${overrides.handles ?? 'cell'}-${index}`, story: cell.story,
        rects: cell.rects.map((rect) => ({ pageId: `p${rect.pageIndex}`, ...rect })),
        cell: { table: overrides.table ?? 'body#0', row: cell.row, column: cell.column, rowspan: 1, colspan: cell.colspan ?? 1, columns: overrides.columns ?? 3,
          exact: overrides.exact ?? true },
      })),
    },
  } as unknown as PeerlessFrameGeometry;
}

const grid = (row: number, y: number, pageIndex = 0): Cell[] => [0, 1, 2].map((column) => ({
  story: `body:t0:r${row}c${column}`, row, column, rects: [{ pageIndex, x: 96 + column * 208, y, width: 208, height: 20 }],
}));

test('column borders and the right edge sit on the cell edges and span the fragment', () => {
  const handles = tableResizeHandles(geometry([...grid(0, 100), ...grid(1, 120)]), new Set([0]));
  expect(handles.map(({ kind, columnEdge, x, top, bottom }) => ({ kind, columnEdge, x, top, bottom }))).toEqual([
    { kind: 'col', columnEdge: 1, x: 304, top: 100, bottom: 140 },
    { kind: 'col', columnEdge: 2, x: 512, top: 100, bottom: 140 },
    { kind: 'right', columnEdge: 3, x: 720, top: 100, bottom: 140 },
  ]);
  expect(handles[0]).toMatchObject({ minDelta: -188, maxDelta: 188 });
  expect(handles[2]).toMatchObject({ minDelta: -188, maxDelta: Infinity });
});

test('a table split across pages gets handles per page fragment and only on the pages asked for', () => {
  const cells = [...grid(0, 900, 0), ...grid(1, 96, 1)];
  expect(tableResizeHandles(geometry(cells), new Set([0, 1])).map(({ pageIndex, top, bottom }) => [pageIndex, top, bottom]))
    .toEqual([[0, 900, 920], [0, 900, 920], [0, 900, 920], [1, 96, 116], [1, 96, 116], [1, 96, 116]]);
  expect(tableResizeHandles(geometry(cells), new Set([1])).every(({ pageIndex }) => pageIndex === 1)).toBe(true);
});

test('a border every row spans has no handle, its neighbours clamp from the nearest borders, and header, footer and note tables get none', () => {
  const merged: Cell[] = [
    { story: 'body:t0:r0c0', row: 0, column: 0, colspan: 2, rects: [{ pageIndex: 0, x: 96, y: 100, width: 416, height: 20 }] },
    { story: 'body:t0:r0c2', row: 0, column: 2, rects: [{ pageIndex: 0, x: 512, y: 100, width: 208, height: 20 }] },
  ];
  expect(tableResizeHandles(geometry(merged), new Set([0])).map(({ columnEdge, minDelta }) => [columnEdge, minDelta])).toEqual([[2, -396], [3, -188]]);
  const header = grid(0, 20).map((cell) => ({ ...cell, story: cell.story.replace('body', 'hf:rId1') }));
  expect(tableResizeHandles(geometry(header), new Set([0]))).toEqual([]);
});

test('geometry of another document version or frame gives no handles', () => {
  expect(tableResizeHandles(geometry(grid(0, 100), { version: 'v0' }), new Set([0]))).toEqual([]);
  expect(tableResizeHandles({ ...geometry(grid(0, 100)), frameEpoch: 4 }, new Set([0]))).toEqual([]);
  expect(tableResizeHandles(undefined, new Set([0]))).toEqual([]);
});

test('a drag becomes one clamped twips delta against the latest geometry', () => {
  const before = geometry(grid(0, 100));
  const [first, , right] = tableResizeHandles(before, new Set([0]));
  expect(tableResizeIntent(before, first!, 60)).toEqual({ kind: 'resizeObject',
    payload: { handle: 'cell-0', expectVersion: 'v1', columnEdge: 1, deltaTwips: 900 } });
  expect(tableResizeIntent(before, first!, 1000)?.payload).toMatchObject({ deltaTwips: 188 * 15 });
  expect(tableResizeIntent(before, right!, -1000)?.payload).toMatchObject({ columnEdge: 3, deltaTwips: -188 * 15 });
  expect(tableResizeIntent(before, first!, 0.01)).toBeNull();
  const later = { ...geometry(grid(0, 140), { version: 'v2' }), documentVersion: 'v2' };
  expect(tableResizeIntent(later, first!, 60)?.payload).toMatchObject({ expectVersion: 'v2', handle: 'cell-0' });
});

test('a drag follows its own table and never one that took its index', () => {
  const before = geometry(grid(0, 100), { table: 'body#1' });
  const [first] = tableResizeHandles(before, new Set([0]));
  expect(tableResizeIntent(geometry(grid(0, 100), { table: 'body#0' }), first!, 60)?.payload).toMatchObject({ handle: 'cell-0', columnEdge: 1 });
  expect(tableResizeIntent(geometry(grid(0, 100), { table: 'body#1', handles: 'other' }), first!, 60)).toBeNull();
  expect(tableResizeIntent(geometry(grid(0, 100), { table: 'body#1', columns: 4 }), first!, 60)).toBeNull();
  expect(tableResizeIntent(geometry(grid(0, 100).map((cell) => ({ ...cell, column: cell.column + 1 })), { table: 'body#1' }), first!, 60)).toBeNull();
  expect(tableResizeIntent(null, first!, 60)).toBeNull();
});

test('estimated cell boxes get no handles, since their edges are not the real ones', () => {
  expect(tableResizeHandles(geometry(grid(0, 100), { exact: false }), new Set([0]))).toEqual([]);
  expect(tableRowResizeHandles(geometry(grid(0, 100), { exact: false }), new Set([0]))).toEqual([]);
});

test('a point finds the nearest border within the tolerance, on its own page and span', () => {
  const handles = tableResizeHandles(geometry([...grid(0, 100), ...grid(1, 120)]), new Set([0]));
  expect(tableHandleAt(handles, 0, 306, 110, 3)?.columnEdge).toBe(1);
  expect(tableHandleAt(handles, 0, 510, 139, 3)?.columnEdge).toBe(2);
  expect(tableHandleAt(handles, 0, 310, 110, 3)).toBeNull();
  expect(tableHandleAt(handles, 0, 304, 99, 3)).toBeNull();
  expect(tableHandleAt(handles, 0, 304, 141, 3)).toBeNull();
  expect(tableHandleAt(handles, 1, 304, 110, 3)).toBeNull();
});

test('row borders sit under each row, bound to a cell of that row, on body tables only', () => {
  const handles = tableRowResizeHandles(geometry([...grid(0, 100), ...grid(1, 120)]), new Set([0]));
  expect(handles.map(({ row, y, left, right, paintedHeight, cell }) => ({ row, y, left, right, paintedHeight, cell }))).toEqual([
    { row: 0, y: 120, left: 96, right: 720, paintedHeight: 20, cell: 'cell-0' },
    { row: 1, y: 140, left: 96, right: 720, paintedHeight: 20, cell: 'cell-3' },
  ]);
  const header = grid(0, 20).map((cell) => ({ ...cell, story: cell.story.replace('body', 'hf:rId1') }));
  expect(tableRowResizeHandles(geometry(header), new Set([0]))).toEqual([]);
  expect(tableRowResizeHandles(geometry(grid(0, 100), { version: 'v0' }), new Set([0]))).toEqual([]);
});

test('a row drag sends the painted height and a clamped delta, and follows its own cell', () => {
  const before = geometry(grid(0, 100));
  const [row] = tableRowResizeHandles(before, new Set([0]));
  expect(tableRowResizeIntent(before, row!, 30)).toEqual({ kind: 'resizeObject',
    payload: { handle: 'cell-0', expectVersion: 'v1', paintedRowTwips: 300, rowDeltaTwips: 450 } });
  expect(tableRowResizeIntent(before, row!, -500)?.payload).toMatchObject({ rowDeltaTwips: -19 * 15 });
  expect(tableRowResizeIntent(before, row!, 0.01)).toBeNull();
  expect(tableRowResizeIntent(geometry(grid(0, 100), { handles: 'other' }), row!, 30)).toBeNull();
  expect(tableRowResizeIntent(null, row!, 30)).toBeNull();
});

test('a point finds the row border near it', () => {
  const handles = tableRowResizeHandles(geometry([...grid(0, 100), ...grid(1, 120)]), new Set([0]));
  expect(tableRowHandleAt(handles, 0, 200, 122, 3)?.row).toBe(0);
  expect(tableRowHandleAt(handles, 0, 200, 124, 3)).toBeNull();
  expect(tableRowHandleAt(handles, 0, 90, 120, 3)).toBeNull();
  expect(tableRowHandleAt(handles, 1, 200, 120, 3)).toBeNull();
});
