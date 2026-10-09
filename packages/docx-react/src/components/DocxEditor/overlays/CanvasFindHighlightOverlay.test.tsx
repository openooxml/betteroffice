import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type {
  DisplayList,
  DisplayListQueries,
  DisplayListRect,
} from '@betteroffice/docx/layout/render';
import {
  CanvasFindHighlightOverlay,
  displayOrder,
  matchesInRange,
  pagePositionIntervals,
} from './CanvasFindHighlightOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

describe('matchesInRange', () => {
  const matches = [
    { displayFrom: 10, displayTo: 14 },
    { displayFrom: 20, displayTo: 60 },
    { displayFrom: 30, displayTo: 33 },
    { displayFrom: 70, displayTo: 75 },
  ];
  const order = displayOrder(matches);

  test('keeps the matches overlapping the range, long ones included', () => {
    expect(matchesInRange(matches, order, 0, 100)).toEqual([0, 1, 2, 3]);
    expect(matchesInRange(matches, order, 14, 30)).toEqual([1]);
    expect(matchesInRange(matches, order, 40, 71)).toEqual([1, 3]);
    expect(matchesInRange(matches, order, 76, 90)).toEqual([]);
  });

  test('orders matches by display position', () => {
    const shuffled = [matches[3], matches[0], matches[2], matches[1]];
    expect(matchesInRange(shuffled, displayOrder(shuffled), 0, 32)).toEqual([1, 3, 2]);
  });
});

test('a page repeating table header rows paints two position intervals', () => {
  const text = (docStart: number, docEnd: number) => ({ kind: 'text', docStart, docEnd });
  const displayList = {
    pages: [
      { pageIndex: 0, primitives: [text(10, 20), text(21, 30), text(31, 40)] },
      { pageIndex: 1, primitives: [text(10, 20), text(900, 950), text(951, 990)] },
    ],
  } as unknown as DisplayList;
  expect(pagePositionIntervals(displayList, { start: 1, end: 1 })).toEqual([
    { from: 10, to: 20 },
    { from: 900, to: 990 },
  ]);
  expect(pagePositionIntervals(displayList, { start: 0, end: 1 })).toEqual([
    { from: 10, to: 40 },
    { from: 900, to: 990 },
  ]);
});

const PAGES = 20;
const PAGE_HEIGHT = 1000;

/** One text run per page at positions `100 * page .. + 50`, and a match on each page. */
function fixture() {
  const displayList = {
    pages: Array.from({ length: PAGES }, (_, pageIndex) => ({
      pageIndex,
      width: 800,
      height: PAGE_HEIGHT,
      primitives: [{ kind: 'text', docStart: pageIndex * 100, docEnd: pageIndex * 100 + 50 }],
    })),
  };
  const queried: number[] = [];
  const queries = {
    displayList,
    pageSize: () => ({ width: 800, height: PAGE_HEIGHT }),
    rangeRects: (from: number): DisplayListRect[] => {
      queried.push(Math.floor(from / 100));
      return [{ pageIndex: Math.floor(from / 100), x: 100, y: 200, width: 30, height: 12 }];
    },
  } as unknown as DisplayListQueries;
  const matches = displayList.pages.map((page) => ({
    displayFrom: page.pageIndex * 100 + 10,
    displayTo: page.pageIndex * 100 + 13,
  }));

  const host = document.createElement('div');
  host.className = 'canvas-pages';
  const column = document.createElement('div');
  host.append(column);
  let scrolled = 0;
  column.getBoundingClientRect = () => new DOMRect(0, -scrolled, 800, PAGES * PAGE_HEIGHT);
  const mountPage = (pageIndex: number) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(pageIndex);
    canvas.getBoundingClientRect = () =>
      new DOMRect(0, 24 + pageIndex * (PAGE_HEIGHT + 16) - scrolled, 800, PAGE_HEIGHT);
    column.append(canvas);
  };
  mountPage(0);
  mountPage(1);
  const target = document.createElement('div');
  document.body.append(host, target);
  return {
    host,
    target,
    queries,
    matches,
    queried,
    mountPage,
    scrollTo: (top: number) => {
      scrolled = top;
    },
  };
}

test('passes the visible page window to range rect queries', () => {
  const { host, target, queries, matches } = fixture();
  const calls: number[][] = [];
  queries.rangeRectsOnPages = (from, to, firstPage, lastPage) => {
    calls.push([from, to, firstPage, lastPage]);
    return [{ pageIndex: Math.floor(from / 100), x: 100, y: 200, width: 30, height: 12 }];
  };
  queries.rangeRects = () => {
    throw new Error('unexpected unrestricted range query');
  };
  const view = render(
    <CanvasFindHighlightOverlay
      matches={matches}
      currentIndex={1}
      overlayTarget={target}
      canvasHostRef={{ current: host }}
      displayListQueries={queries}
      sidebarOpen={false}
      zoom={1}
    />
  );
  expect(calls).toEqual([[10, 13, 0, 1], [110, 113, 0, 1]]);
  expect(target.querySelectorAll('.docx-find-highlight')).toHaveLength(1);
  expect(target.querySelectorAll('.docx-find-highlight-current')).toHaveLength(1);
  view.unmount();
});

test('resolves only the matches on the pages in view, and follows scrolling', async () => {
  const { host, target, queries, matches, queried, mountPage, scrollTo } = fixture();
  const view = render(
    <CanvasFindHighlightOverlay
      matches={matches}
      currentIndex={1}
      overlayTarget={target}
      canvasHostRef={{ current: host }}
      displayListQueries={queries}
      sidebarOpen={false}
      zoom={1}
    />
  );
  expect(queried).toEqual([0, 1]);
  expect(target.querySelectorAll('.docx-find-highlight')).toHaveLength(1);
  expect(target.querySelectorAll('.docx-find-highlight-current')).toHaveLength(1);

  queried.length = 0;
  // page 10 at the top of the window
  scrollTo(24 + 10 * (PAGE_HEIGHT + 16));
  act(() => {
    document.dispatchEvent(new Event('scroll'));
  });
  await waitFor(() => expect(queried).toEqual([9, 10, 11]));
  // those pages are not painted yet
  expect(target.querySelector('[data-testid="canvas-find-highlights"]')).toBeNull();

  queried.length = 0;
  act(() => mountPage(10));
  await waitFor(() => {
    expect(queried).toEqual([9, 10, 11]);
    expect(target.querySelectorAll('.docx-find-highlight')).toHaveLength(1);
  });
  view.unmount();
});
