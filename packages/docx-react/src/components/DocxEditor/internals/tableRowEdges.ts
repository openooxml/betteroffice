/** A cell's box on one page, in page-local layout px. */
export interface TableCellBox {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

/** The bottom border of one table row, in page-local layout px. */
export interface TableRowEdge<Cell> {
  row: number;
  pageIndex: number;
  left: number;
  right: number;
  y: number;
  /** Height of the row across every page it covers. */
  paintedHeight: number;
  /** A cell that sits in this row alone. */
  cell: Cell;
}

/**
 * Bottom borders of one table's rows that end on `pageIndices`. A row split
 * across pages ends on its last box, and a row only vertically merged cells
 * reach has no cell of its own to resize by, so it has no edge.
 */
export function tableRowEdges<Cell>(
  cells: ReadonlyArray<{ row: number; rowspan: number; boxes: readonly TableCellBox[]; cell: Cell }>,
  pageIndices: ReadonlySet<number>
): TableRowEdge<Cell>[] {
  const extents = new Map<number, { left: number; right: number }>();
  const rows = new Map<number, { pageIndex: number; y: number; paintedHeight: number; cell: Cell }>();
  for (const { row, rowspan, boxes, cell } of cells) {
    for (const box of boxes) {
      const extent = extents.get(box.pageIndex);
      extents.set(box.pageIndex, {
        left: Math.min(extent?.left ?? box.x, box.x),
        right: Math.max(extent?.right ?? box.x + box.width, box.x + box.width),
      });
    }
    if (rowspan !== 1 || boxes.length === 0) continue;
    const last = boxes.reduce((latest, box) =>
      box.pageIndex > latest.pageIndex || (box.pageIndex === latest.pageIndex && box.y > latest.y) ? box : latest);
    const paintedHeight = boxes.reduce((sum, box) => sum + box.height, 0);
    const y = last.y + last.height;
    const seen = rows.get(row);
    if (!seen || last.pageIndex > seen.pageIndex || (last.pageIndex === seen.pageIndex && y > seen.y)) {
      rows.set(row, { pageIndex: last.pageIndex, y, paintedHeight: Math.max(paintedHeight, seen?.paintedHeight ?? 0), cell });
    } else seen.paintedHeight = Math.max(seen.paintedHeight, paintedHeight);
  }
  const edges: TableRowEdge<Cell>[] = [];
  for (const [row, edge] of rows) {
    const extent = extents.get(edge.pageIndex);
    if (extent && pageIndices.has(edge.pageIndex)) edges.push({ row, ...edge, ...extent });
  }
  return edges.sort((a, b) => a.row - b.row);
}
