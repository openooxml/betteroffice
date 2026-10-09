import { expect, test } from 'bun:test';
import { anchorPositionsFromPoints } from './displayListAnchors';
import type { DisplayListQueries, DisplayListRect } from './displayListQueries';
import type { YrsSidebarDisplayPoint } from './yrsSidebarProjection';

const point = (position: number): YrsSidebarDisplayPoint => ({ story: 'body', position });
const queries = {
  pageCount: () => 1,
  pageSize: () => ({ width: 600, height: 800 }),
  anchorRect: (position: number) => position === 0 ? null : ({
    pageIndex: 0, x: 0, y: position, width: 1, height: 10,
  }),
  hfAnchorRects: () => [{ pageIndex: 0, x: 0, y: 30, width: 1, height: 10 }],
} as unknown as DisplayListQueries;

test('the first registrable point wins and missing rects can be retried', () => {
  expect(anchorPositionsFromPoints([
    ['first', point(5)], ['first', point(10)],
    ['retry', point(0)], ['retry', point(20)],
    ['null', null],
  ], queries)).toEqual(new Map([['first', 29], ['retry', 44]]));
});

test('a null projected Y skips a point and permits a later point for its key', () => {
  const project = (rect: DisplayListRect) => rect.y === 5 ? null : rect.y * 2;
  expect(anchorPositionsFromPoints([
    ['retry', point(5)], ['retry', point(10)], ['skip', point(5)],
  ], queries, undefined, project)).toEqual(new Map([['retry', 20]]));
});

test('header/footer points use their region rect and unknown regions remain unplaced', () => {
  expect(anchorPositionsFromPoints([
    ['header', { ...point(1), hfRid: 'h' }],
    ['missing', { ...point(1), hfRid: 'unknown' }],
  ], queries, new Map([['h', 'header']]))).toEqual(new Map([['header', 54]]));
});
