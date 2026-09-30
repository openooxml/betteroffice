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
const PAGES = 12;
const GAP = 24;
const SCROLLER_TOP = 40.3;
const SCROLLER_BOX = 300.4;
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

/** Pages under `scroller`, drawn at `zoom`, with `scroller` at client top `scrollerTop`. */
function pagesUnder(zoom: number, scrollTop: number, scrollerTop: number) {
  const host = document.createElement('div');
  for (let index = 0; index < PAGES; index += 1) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    const top = GAP + index * (PAGE.height + GAP);
    canvas.getBoundingClientRect = () =>
      rect(scrollerTop + (top - scrollTop) * zoom, PAGE.height * zoom, PAGE.width * zoom);
    host.appendChild(canvas);
  }
  return host;
}

/** A scroller drawn under an ancestor CSS `zoom`, recording where it is asked to scroll. */
function scene(zoom: number, scrollTop: number) {
  const host = pagesUnder(zoom, scrollTop, SCROLLER_TOP);
  const moves: number[] = [];
  const behaviors: Array<ScrollBehavior | undefined> = [];
  const scroller = document.createElement('div');
  Object.defineProperties(scroller, {
    currentCSSZoom: { value: zoom },
    clientHeight: { value: SCROLLER_HEIGHT },
    scrollTop: { value: scrollTop },
  });
  scroller.getBoundingClientRect = () => rect(SCROLLER_TOP, SCROLLER_BOX * zoom);
  scroller.scrollTo = ((options: ScrollToOptions) => (
    moves.push(options.top ?? NaN), behaviors.push(options.behavior)
  )) as typeof scroller.scrollTo;
  return { host, scroller, moves, behaviors };
}

const queries = {
  pageCount: () => PAGES,
  pageSize: () => PAGE,
  pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, ...PAGE }),
} as unknown as DisplayListQueries;

function scrollApi(host: HTMLElement, scroller: HTMLElement) {
  return renderHook(() =>
    usePagedScrollApi({
      pagesContainerRef: { current: host as HTMLDivElement },
      yrsInputRef: { current: null },
      yrsSession: null,
      yrsLocToDisplayPosition: () => null,
      getScrollContainer: () => scroller as HTMLDivElement,
      displayListQueries: queries,
      canvasHostRef: { current: host as HTMLDivElement },
    })
  ).result.current;
}

const pageCentre = (page: number) => GAP + (page - 1) * (PAGE.height + GAP) + PAGE.height / 2;

test('scrollToPage centres the page under an ancestor CSS zoom', () => {
  for (const zoom of [0.8, 1, 1.25]) {
    const { host, scroller, moves } = scene(zoom, 100.5);
    scrollApi(host, scroller).scrollToPageImpl(3);
    expect(moves).toHaveLength(1);
    expect(moves[0]).toBeCloseTo(pageCentre(3) - SCROLLER_HEIGHT / 2, 9);
  }
});

test('scrollToPage centres the page in the window when the root scrolls', () => {
  const root = document.documentElement;
  const scrollTop = 1_000;
  const host = pagesUnder(1, scrollTop, 0);
  const moves: number[] = [];
  const scrollTo = root.scrollTo;
  Object.defineProperty(root, 'scrollTop', { value: scrollTop, configurable: true });
  root.scrollTo = ((options: ScrollToOptions) =>
    moves.push(options.top ?? NaN)) as typeof root.scrollTo;
  try {
    scrollApi(host, root).scrollToPageImpl(9);
    expect(moves[0]).toBeCloseTo(pageCentre(9) - window.innerHeight / 2);
  } finally {
    root.scrollTo = scrollTo;
    delete (root as { scrollTop?: number }).scrollTop;
  }
});

test('scrollToPage animates within two viewports and jumps farther, in layout pixels', () => {
  for (const zoom of [0.8, 1, 1.25]) {
    const { host, scroller, behaviors } = scene(zoom, 0);
    const api = scrollApi(host, scroller);
    for (let page = 1; page <= 4; page += 1) api.scrollToPageImpl(page);
    expect(behaviors).toEqual(['smooth', 'smooth', 'smooth', 'instant']);
  }
});

test('an explicitly unanimated scroll keeps the default behaviour at any distance', () => {
  const { host, scroller, behaviors } = scene(1, 0);
  const far = { pageIndex: PAGES - 1, x: 0, y: 0, width: 10, height: 10 };
  const api = renderHook(() =>
    usePagedScrollApi({
      pagesContainerRef: { current: host as HTMLDivElement },
      yrsInputRef: { current: null },
      yrsSession: null,
      yrsLocToDisplayPosition: () => null,
      getScrollContainer: () => scroller as HTMLDivElement,
      displayListQueries: { ...queries, anchorRect: () => far } as DisplayListQueries,
      canvasHostRef: { current: host as HTMLDivElement },
    })
  ).result.current;
  api.scrollToPositionImpl(1, true);
  api.revealPositionImpl(1);
  expect(behaviors).toEqual(['auto', 'instant']);
});

test('scrollToPage computes an unzoomed target exactly as before', () => {
  const scrollTop = 100.1;
  const { host, scroller, moves } = scene(1, scrollTop);
  scrollApi(host, scroller).scrollToPageImpl(3);
  const clientY = host.children[2]!.getBoundingClientRect().top + PAGE.height / 2;
  expect(moves[0]).toBe(scrollTop + clientY - SCROLLER_TOP - SCROLLER_HEIGHT / 2);
});
