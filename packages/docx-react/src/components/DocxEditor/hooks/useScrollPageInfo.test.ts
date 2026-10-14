import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { PagedEditorRef } from '../PagedEditor';
import { pageAtViewportMiddle, useScrollPageInfo } from './useScrollPageInfo';

const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE_HEIGHT = 200;
const layout = (pages: number, partial = false) =>
  ({
    pageSize: { w: 100, h: PAGE_HEIGHT },
    pages: Array.from({ length: pages }, () => ({ size: { w: 100, h: PAGE_HEIGHT } })),
    ...(partial ? { partial } : {}),
  }) as unknown as Layout;

function rectAt(top: number, height: number): DOMRect {
  return {
    left: 0,
    top,
    width: 100,
    height,
    right: 100,
    bottom: top + height,
    x: 0,
    y: top,
    toJSON() {},
  } as DOMRect;
}

/**
 * A scroller at client y 100, 400 px tall, over five page canvases `scale` px per layout px,
 * 16 px apart below 24 px of padding and `above` px of other content.
 */
function pagesScroller(pages = 5, { scale = 1, above = 0 } = {}) {
  const scroller = document.createElement('div');
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvases = Array.from({ length: pages }, (_, index) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    host.appendChild(canvas);
    return canvas;
  });
  scroller.appendChild(host);
  scroller.getBoundingClientRect = () => rectAt(100, 400);
  const scrollTo = (scrollTop: number) => {
    canvases.forEach((canvas, index) => {
      const top = 100 + (above + 24 + index * (PAGE_HEIGHT + 16)) * scale - scrollTop;
      canvas.getBoundingClientRect = () => rectAt(top, PAGE_HEIGHT * scale);
    });
  };
  scrollTo(0);
  return { scroller, scrollTo };
}

test('the current page is the canvas under the middle of the viewport', () => {
  const { scroller, scrollTo } = pagesScroller();
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(1);
  // the middle at 300 in the scroller: page 2 spans 240..440
  scrollTo(100);
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(2);
  // the middle in the gap below page 2 belongs to page 3
  scrollTo(245);
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(3);
  scrollTo(5000);
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(5);
});

test('zoom and content above the pages count', () => {
  const { scroller, scrollTo } = pagesScroller(40, { scale: 0.713, above: 48 });
  // page 30 starts at (48 + 24 + 30 * 216) * 0.713 in the scroll content
  const pageTop = (48 + 24 + 30 * 216) * 0.713;
  scrollTo(pageTop - 200 + 10);
  expect(pageAtViewportMiddle(scroller, layout(40))).toBe(31);
});

test('only the part of the scroller inside the window counts', () => {
  const { scroller } = pagesScroller();
  scroller.getBoundingClientRect = () => rectAt(-400, 4000);
  expect(window.innerHeight).toBe(768);
  // the visible band is the window's 0..768, its middle at 384: page 2 spans 340..540
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(2);
});

test('without canvases the layout page heights place the page', () => {
  const scroller = document.createElement('div');
  scroller.getBoundingClientRect = () => rectAt(0, 400);
  Object.defineProperty(scroller, 'clientHeight', { value: 400 });
  Object.defineProperty(scroller, 'scrollTop', { value: 300, writable: true });
  expect(pageAtViewportMiddle(scroller, layout(5))).toBe(3);
});

test('the ref reads the page at call time, before any scroll event re-renders', () => {
  const { scroller, scrollTo } = pagesScroller();
  let current: Layout | null = layout(5);
  const editor = { getLayout: () => current } as unknown as PagedEditorRef;
  const hook = renderHook(() =>
    useScrollPageInfo({
      scrollContainerRef: { current: scroller as HTMLDivElement },
      pagedEditorRef: { current: editor },
    })
  );
  expect(hook.result.current.readCurrentPage()).toBe(1);
  scrollTo(100 + 216 * 2);
  expect(hook.result.current.readCurrentPage()).toBe(4);
  expect(hook.result.current.scrollPageInfo.currentPage).toBe(1);
  current = layout(3);
  expect(hook.result.current.readCurrentPage()).toBe(3);
  current = layout(0);
  expect(hook.result.current.readCurrentPage()).toBeNull();
});

test('a partial layout reads the page among its pages but leaves the indicator alone', () => {
  const { scroller, scrollTo } = pagesScroller(3);
  let current: Layout = layout(3, true);
  const editor = { getLayout: () => current } as unknown as PagedEditorRef;
  const hook = renderHook(() =>
    useScrollPageInfo({
      scrollContainerRef: { current: scroller as HTMLDivElement },
      pagedEditorRef: { current: editor },
    })
  );
  scrollTo(100);
  expect(hook.result.current.readCurrentPage()).toBe(2);
  act(() => void scroller.dispatchEvent(new Event('scroll')));
  expect(hook.result.current.scrollPageInfo).toEqual({
    currentPage: 1,
    totalPages: 0,
    visible: false,
  });
  current = layout(3);
  act(() => void scroller.dispatchEvent(new Event('scroll')));
  expect(hook.result.current.scrollPageInfo).toEqual({
    currentPage: 2,
    totalPages: 3,
    visible: true,
  });
  current = layout(2, true);
  scrollTo(5000);
  expect(hook.result.current.readCurrentPage()).toBe(2);
});
