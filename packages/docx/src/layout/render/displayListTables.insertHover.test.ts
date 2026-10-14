import { expect, test } from 'bun:test';
import type { DisplayList, DisplayPrimitive } from './displayList';
import { detectDisplayListTableInsertHover } from './displayListTables';

function cell(row: number, col: number, docStart: number): DisplayPrimitive {
  return {
    kind: 'rect',
    x: 10 + col * 100,
    y: 10 + row * 20,
    w: 100,
    h: 20,
    fill: '#ffffff',
    cell: { row, col, rowSpan: 1, colSpan: 1 },
    table: { tableId: 'A' },
    docStart,
    docEnd: docStart + 1,
  } as DisplayPrimitive;
}

test('the insert button keeps its offset from the table in its own pixels', () => {
  const list: DisplayList = {
    pages: [
      {
        pageIndex: 0,
        width: 600,
        height: 800,
        primitives: [cell(0, 0, 5), cell(0, 1, 7), cell(1, 0, 9), cell(1, 1, 11)],
      },
    ],
  };
  const hover = (x: number, y: number, buttonZoom?: number) =>
    detectDisplayListTableInsertHover({
      list,
      pageIndex: 0,
      x,
      y,
      canvasRect: { left: 50, top: 30, width: 900, height: 1200 },
      pageSize: { width: 600, height: 800 },
      tableKeyOf: () => 'A',
      cellPmPosOf: (_table, row, col) => 5 + row * 4 + col * 2,
      ...(buttonZoom === undefined ? {} : { buttonZoom }),
    });
  for (const [x, y, type] of [
    [10, 20, 'row'],
    [60, 10, 'column'],
  ] as const) {
    const plain = hover(x, y)!;
    expect(plain.type).toBe(type);
    expect(hover(x, y, 1)).toEqual(plain);
    const zoomed = hover(x, y, 2)!;
    const [offsetX, offsetY] = type === 'row' ? [24, 10] : [10, 24];
    expect(zoomed).toEqual({
      ...plain,
      clientX: plain.clientX - offsetX,
      clientY: plain.clientY - offsetY,
    });
  }
});
