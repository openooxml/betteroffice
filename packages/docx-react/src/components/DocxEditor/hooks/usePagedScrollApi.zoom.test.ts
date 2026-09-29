import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { usePagedScrollApi } from './usePagedScrollApi';

const { cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE = { width: 100, height: 200 };
const GAP = 24;
const SCROLLER_TOP = 40;
const SCROLLER_HEIGHT = 300;

function rect(top: number, height: number, width = 0): DOMRect {
  return {
    top,
    height,
    width,
    left: 0,
    right: width,
    bottom: top + height,
    x: 0,
    y: top,
  } as DOMRect;
}

/** A scroller and pages drawn under an ancestor CSS `zoom`. */
function scene(zoom: number, scrollTop: number) {
  const host = document.createElement('div');
  for (let index = 0; index < 3; index += 1) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    const top = GAP + index * (PAGE.height + GAP);
    canvas.getBoundingClientRect = () =>
      rect(SCROLLER_TOP + (top - scrollTop) * zoom, PAGE.height * zoom, PAGE.width * zoom);
    host.appendChild(canvas);
  }
  const moves: number[] = [];
  const scroller = document.createElement('div');
  Object.defineProperties(scroller, {
    offsetHeight: { value: SCROLLER_HEIGHT },
    clientHeight: { value: SCROLLER_HEIGHT },
    scrollTop: { value: scrollTop },
  });
  scroller.getBoundingClientRect = () => rect(SCROLLER_TOP, SCROLLER_HEIGHT * zoom);
  scroller.scrollTo = ((options: ScrollToOptions) =>
    moves.push(options.top ?? NaN)) as typeof scroller.scrollTo;
  return { host, scroller, moves };
}

const queries = {
  pageCount: () => 3,
  pageSize: () => PAGE,
  pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, ...PAGE }),
} as unknown as DisplayListQueries;

test('scrollToPage centres the page under an ancestor CSS zoom', () => {
  for (const zoom of [0.8, 1, 1.25]) {
    const { host, scroller, moves } = scene(zoom, 100);
    const { result } = renderHook(() =>
      usePagedScrollApi({
        pagesContainerRef: { current: host },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => scroller,
        displayListQueries: queries,
        canvasHostRef: { current: host },
      })
    );
    result.current.scrollToPageImpl(3);
    const pageCentre = GAP + 2 * (PAGE.height + GAP) + PAGE.height / 2;
    expect(moves).toHaveLength(1);
    expect(moves[0]).toBeCloseTo(pageCentre - SCROLLER_HEIGHT / 2);
  }
});
