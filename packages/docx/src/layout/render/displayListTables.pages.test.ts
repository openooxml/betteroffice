import { expect, test } from 'bun:test';
import type { DisplayList, DisplayPrimitive } from './displayList';
import {
  deriveDisplayListTableFragments,
  deriveDisplayListTableFragmentsOnPages,
} from './displayListTables';

function cell(tableId: string, row: number, y: number, docStart?: number): DisplayPrimitive {
  return {
    kind: 'rect',
    x: 10,
    y,
    w: 100,
    h: 20,
    fill: '#ffffff',
    cell: { row, col: 0, rowSpan: 1, colSpan: 1 },
    table: { tableId },
    ...(docStart === undefined ? {} : { docStart, docEnd: docStart + 1 }),
  } as DisplayPrimitive;
}

test('fragments on a page window match the whole-list fragments of those pages', () => {
  const page = (pageIndex: number, primitives: DisplayPrimitive[]) => ({
    pageIndex,
    width: 600,
    height: 800,
    primitives,
  });
  // Table A starts on page 0; its rows on pages 1 and 2 carry no position.
  const list: DisplayList = {
    pages: [
      page(0, [cell('A', 0, 10, 5), cell('A', 1, 40, 9)]),
      page(1, [cell('A', 2, 10)]),
      page(2, [cell('A', 3, 10), cell('B', 0, 400, 50)]),
      page(3, [cell('B', 1, 10)]),
    ],
  };
  const tableKeyOf = (position: number | undefined) =>
    position === undefined ? null : position < 40 ? 'A' : 'B';
  const all = deriveDisplayListTableFragments(list, tableKeyOf);
  for (let start = 0; start < 4; start += 1) {
    for (let end = start; end < 4; end += 1) {
      expect(deriveDisplayListTableFragmentsOnPages(list, tableKeyOf, start, end)).toEqual(
        all.filter((fragment) => fragment.pageIndex >= start && fragment.pageIndex <= end)
      );
    }
  }
  expect(deriveDisplayListTableFragmentsOnPages(list, tableKeyOf, 1, 1)).toHaveLength(1);
});

test('a page window without table cells reads no other page', () => {
  let read = false;
  const list = {
    pages: [
      {
        pageIndex: 0,
        width: 600,
        height: 800,
        get primitives() {
          read = true;
          return [cell('A', 0, 10, 5)];
        },
      },
      { pageIndex: 1, width: 600, height: 800, primitives: [] },
    ],
  } as DisplayList;
  expect(deriveDisplayListTableFragmentsOnPages(list, () => 'A', 1, 1)).toEqual([]);
  expect(read).toBe(false);
});
