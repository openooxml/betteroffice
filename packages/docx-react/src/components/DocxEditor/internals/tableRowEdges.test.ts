import { expect, test } from 'bun:test';
import { tableRowEdges } from './tableRowEdges';

const box = (pageIndex: number, x: number, y: number, height: number) => ({ pageIndex, x, y, width: 100, height });
const pages = new Set([0, 1]);

test('each row ends at its tallest cell and spans the table', () => {
  expect(tableRowEdges([
    { row: 0, rowspan: 1, boxes: [box(0, 96, 100, 20)], cell: 'a' },
    { row: 0, rowspan: 1, boxes: [box(0, 196, 100, 35)], cell: 'b' },
    { row: 1, rowspan: 1, boxes: [box(0, 96, 135, 20)], cell: 'c' },
  ], pages)).toEqual([
    { row: 0, pageIndex: 0, left: 96, right: 296, y: 135, paintedHeight: 35, cell: 'b' },
    { row: 1, pageIndex: 0, left: 96, right: 296, y: 155, paintedHeight: 20, cell: 'c' },
  ]);
});

test('a split row ends on its last page and counts every page of its height', () => {
  const cells = [{ row: 0, rowspan: 1, boxes: [box(0, 96, 900, 50), box(1, 96, 96, 30)], cell: 'a' }];
  expect(tableRowEdges(cells, pages)).toEqual([{ row: 0, pageIndex: 1, left: 96, right: 196, y: 126, paintedHeight: 80, cell: 'a' }]);
  expect(tableRowEdges(cells, new Set([0]))).toEqual([]);
});

test('a merged cell gives no edge of its own but widens the table', () => {
  expect(tableRowEdges([
    { row: 0, rowspan: 2, boxes: [box(0, 96, 100, 40)], cell: 'merged' },
    { row: 0, rowspan: 1, boxes: [box(0, 196, 100, 20)], cell: 'a' },
    { row: 1, rowspan: 1, boxes: [box(0, 196, 120, 20)], cell: 'b' },
  ], pages).map(({ row, left, cell }) => [row, left, cell])).toEqual([[0, 96, 'a'], [1, 96, 'b']]);
  expect(tableRowEdges([{ row: 0, rowspan: 2, boxes: [box(0, 96, 100, 40)], cell: 'merged' }], pages)).toEqual([]);
});

test('a repeated header row counts once, on the page it really sits on, even for a cell with no content', () => {
  const copy = (pageIndex: number, repeatedHeader: boolean, y: number) => ({ ...box(pageIndex, 96, y, 24), ...(repeatedHeader ? { repeatedHeader } : {}) });
  const cells = [
    { row: 0, rowspan: 1, boxes: [copy(0, false, 100), copy(1, true, 96), copy(2, true, 96)], cell: 'a' },
    { row: 0, rowspan: 1, boxes: [box(0, 196, 100, 24), box(1, 196, 96, 24), box(2, 196, 96, 24)], cell: 'empty' },
  ];
  expect(tableRowEdges(cells, new Set([0, 1, 2]))).toEqual([{ row: 0, pageIndex: 0, left: 96, right: 296, y: 124, paintedHeight: 24, cell: 'a' }]);
  expect(tableRowEdges(cells, new Set([1, 2]))).toEqual([]);
});
