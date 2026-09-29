import { expect, test } from 'bun:test';
import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
  type DisplayListVisualLine,
} from '@betteroffice/docx/layout/render';
import type { YrsStickyPosition } from '@betteroffice/docx/yrs';
import {
  captureDisplayListScrollAnchor,
  captureDisplayListViewportAnchor,
  restoreDisplayListScrollAnchor,
  restoreDisplayListViewportAnchor,
} from './scrollRestore';

const PAGE_HEIGHT = 800;
const PAGE_GAP = 24;
const SCROLLER_TOP = 60.3;
const SCROLLER_HEIGHT = 200;
const STICKY: YrsStickyPosition = { story: 'body', encoded: Uint8Array.of(1) };

function domRect(top: number, width: number, height: number): DOMRect {
  return {
    x: 0,
    y: top,
    top,
    left: 0,
    right: width,
    bottom: top + height,
    width,
    height,
    toJSON: () => ({}),
  };
}

/** Layout-pixel top of a page-local y, the space `scrollTop` moves through. */
function documentTop(pageIndex: number, pageY: number): number {
  return PAGE_GAP + pageIndex * (PAGE_HEIGHT + PAGE_GAP) + pageY;
}

/** Pages and scroller under an ancestor CSS `zoom`: client rects scale, layout sizes do not. */
function zoomedScene(zoom: number, scrollTop: number, pageCount: number) {
  const state = { scrollTop, pageCount };
  const height = () =>
    PAGE_GAP * 2 + state.pageCount * PAGE_HEIGHT + (state.pageCount - 1) * PAGE_GAP;
  const host = {
    querySelector: () => null,
    classList: { contains: () => false },
    getBoundingClientRect: () =>
      domRect(SCROLLER_TOP - state.scrollTop * zoom, 600 * zoom, height() * zoom),
    offsetWidth: 600,
    clientWidth: 600,
  } as unknown as HTMLElement;
  const scroller = {
    style: { overflowAnchor: 'none', setProperty() {} },
    get scrollTop() {
      return state.scrollTop;
    },
    set scrollTop(next: number) {
      state.scrollTop = Math.min(Math.max(0, next), Math.max(0, height() - SCROLLER_HEIGHT));
    },
    get scrollHeight() {
      return height();
    },
    clientHeight: SCROLLER_HEIGHT,
    currentCSSZoom: zoom,
    getBoundingClientRect: () => domRect(SCROLLER_TOP, 600 * zoom, SCROLLER_HEIGHT * zoom),
  } as unknown as HTMLElement;
  const clientTop = (pageIndex: number, pageY: number) =>
    SCROLLER_TOP + (documentTop(pageIndex, pageY) - state.scrollTop) * zoom;
  return { state, host, scroller, clientTop };
}

function queries(lines: DisplayListVisualLine[], pageCount: number, caret?: DisplayListRect) {
  return {
    pageCount: () => pageCount,
    pageSize: () => ({ width: 600, height: PAGE_HEIGHT }),
    visualLines: () => lines,
    visualLinesOnPage: (pageIndex: number) => lines.filter((line) => line.pageIndex === pageIndex),
    visualLineExtent: (pageIndex: number) => {
      const onPage = lines.filter((line) => line.pageIndex === pageIndex);
      return onPage.length === 0
        ? null
        : {
            top: Math.min(...onPage.map((l) => l.y)),
            bottom: Math.max(...onPage.map((l) => l.y + l.height)),
          };
    },
    anchorRect: () => caret ?? null,
  } as unknown as DisplayListQueries;
}

function line(pageIndex: number, y: number): DisplayListVisualLine {
  return {
    pageIndex,
    x: 0,
    y,
    width: 500,
    height: 16,
    baseline: y + 12,
    from: 1,
    to: 20,
    paraId: 'p',
  };
}

test('a page added above the viewport keeps the anchored line on screen under an ancestor CSS zoom', () => {
  for (const zoom of [0.8, 1.25]) {
    const { state, host, scroller, clientTop } = zoomedScene(zoom, 1_000, 2);
    const anchor = captureDisplayListViewportAnchor(
      queries([line(1, 200)], 2),
      host,
      scroller,
      () => STICKY
    );
    const before = clientTop(1, 200);
    state.pageCount = 3;
    restoreDisplayListViewportAnchor(anchor, queries([line(2, 200)], 3), host, scroller, () => 1);
    expect(clientTop(2, 200)).toBeCloseTo(before);
  }
});

test('the caret line stays pinned under an ancestor CSS zoom', () => {
  for (const zoom of [0.8, 1.25]) {
    const { host, scroller } = zoomedScene(zoom, 700, 2);
    const caret = (y: number): DisplayListRect => ({ pageIndex: 0, x: 0, y, width: 2, height: 16 });
    const anchor = captureDisplayListScrollAnchor(queries([], 2, caret(700)), host, scroller, 42);
    restoreDisplayListScrollAnchor(anchor, queries([], 2, caret(740)), host, scroller);
    expect(scroller.scrollTop).toBeCloseTo(740);
  }
});

test('unzoomed anchors restore exactly as before', () => {
  const scrollTop = 2_800;
  const caret = (y: number): DisplayListRect => ({ pageIndex: 3, x: 0, y, width: 2, height: 16 });
  const pinned = zoomedScene(1, scrollTop, 4);
  const pageTop = resolveDisplayPageClientRect(pinned.host, queries([], 4), 3)!.top;
  const expected = scrollTop + (pageTop + 540.7) - SCROLLER_TOP - (pageTop + 500.3 - SCROLLER_TOP);
  const scrollAnchor = captureDisplayListScrollAnchor(
    queries([], 4, caret(500.3)),
    pinned.host,
    pinned.scroller,
    42
  );
  restoreDisplayListScrollAnchor(
    scrollAnchor,
    queries([], 4, caret(540.7)),
    pinned.host,
    pinned.scroller
  );
  expect(pinned.scroller.scrollTop).toBe(expected);

  const anchored = zoomedScene(1, scrollTop, 4);
  const viewportAnchor = captureDisplayListViewportAnchor(
    queries([line(3, 500.3)], 4),
    anchored.host,
    anchored.scroller,
    () => STICKY
  );
  restoreDisplayListViewportAnchor(
    viewportAnchor,
    queries([line(3, 540.7)], 4),
    anchored.host,
    anchored.scroller,
    () => 1
  );
  expect(anchored.scroller.scrollTop).toBe(expected);
});
