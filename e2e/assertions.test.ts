import { expect, test } from 'bun:test';
import type { Layout } from '../packages/docx/src/layout/pagination';
import { assertLayoutParity, assertWorkbookMirror } from './assertions';

const layout = (): Layout =>
  ({
    pageSize: { w: 800, h: 1100 },
    pages: [
      {
        number: 1,
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
        size: { w: 800, h: 1100 },
        fragments: [
          {
            kind: 'paragraph',
            blockId: 'p1',
            x: 50,
            y: 50,
            width: 700,
            height: 30,
            fromLine: 0,
            toLine: 1,
          },
        ],
      },
    ],
  } as Layout);

test('layout parity rejects consistent but wrong page sizes and displaced content', () => {
  assertLayoutParity(layout(), layout());
  const wrongSize = layout();
  wrongSize.pageSize.w = 900;
  expect(() => assertLayoutParity(layout(), wrongSize)).toThrow();
  const displaced = layout();
  displaced.pages[0].fragments[0].y += 10;
  expect(() => assertLayoutParity(layout(), displaced)).toThrow();
});

test('cross-SDK parity rejects a corrupted final formula value', () => {
  const handle = {
    cell: () => ({ a1: 'A392', input: '=LEN(A391)', isFormula: true }),
    searchText: () => [
      {
        a1: 'A392',
        row: 391,
        col: 0,
        sheet: 0,
        sheetId: 's',
        sheetName: 'Sheet1',
        text: '11',
      },
    ],
  };
  const cells = [{ row: 391, col: 0, value: 11 }];
  assertWorkbookMirror(
    handle,
    { values: { A392: 11 }, formulas: { A392: 'LEN(A391)' } },
    cells
  );
  expect(() =>
    assertWorkbookMirror(
      handle,
      { values: { A392: 999 }, formulas: { A392: 'LEN(A391)' } },
      cells
    )
  ).toThrow();
  expect(() =>
    assertWorkbookMirror(
      { ...handle, searchText: () => [] },
      { values: { A392: 11 }, formulas: { A392: 'LEN(A391)' } },
      cells
    )
  ).toThrow();
});
