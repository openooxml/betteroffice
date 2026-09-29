import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type {
  DisplayListQueries,
  DisplayListRect,
  DisplayPage,
} from '@betteroffice/docx/layout/render';
import { usePagedScrollApi } from './usePagedScrollApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function pagedDom() {
  const scroller = document.createElement('div');
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  for (const index of [6, 8]) {
    const page = document.createElement('div');
    page.className = 'canvas-page';
    page.dataset.pageIndex = String(index);
    page.getBoundingClientRect = () =>
      ({
        top: index * 1000 - scroller.scrollTop,
        bottom: (index + 1) * 1000 - scroller.scrollTop,
        height: 1000,
        left: 0,
        right: 800,
        width: 800,
      } as DOMRect);
    host.append(page);
  }
  scroller.append(host);
  document.body.append(scroller);
  Object.defineProperty(scroller, 'clientHeight', { value: 400 });
  scroller.getBoundingClientRect = () =>
    ({ top: 0, bottom: 400, height: 400, left: 0, right: 800, width: 800 } as DOMRect);
  const scrolls: number[] = [];
  scroller.scrollTo = ((options: ScrollToOptions) => {
    scrolls.push(options.top ?? 0);
    scroller.scrollTop = options.top ?? 0;
  }) as typeof scroller.scrollTo;
  return { scroller, host, scrolls };
}

function queries(built: boolean | number[], anchor: DisplayListRect): DisplayListQueries {
  const pages: DisplayPage[] = Array.from({ length: 10 }, (_, pageIndex) => {
    const page: DisplayPage = { pageIndex, width: 800, height: 1000, primitives: [] };
    const isBuilt = Array.isArray(built) ? built.includes(pageIndex) : built || pageIndex < 6;
    if (!isBuilt) Object.assign(page, { unbuilt: true, positionSpan: [400, 900] });
    return page;
  });
  return {
    displayList: { pages },
    anchorRect: () => anchor,
    pageSize: () => ({ width: 800, height: 1000 }),
  } as unknown as DisplayListQueries;
}

test('a position on an unbuilt page is scrolled to again once the page is built', async () => {
  const { scroller, host, scrolls } = pagedDom();
  const placeholder = { pageIndex: 6, x: 20, y: 20, width: 0, height: 0 };
  const match = { pageIndex: 6, x: 20, y: 900, width: 0, height: 16 };
  const { result, rerender } = renderHook(
    ({ displayListQueries }) =>
      usePagedScrollApi({
        pagesContainerRef: { current: host },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => scroller,
        displayListQueries,
      }),
    { initialProps: { displayListQueries: queries(false, placeholder) } }
  );
  await act(async () => result.current.scrollToPositionImpl(500));
  expect(scrolls).toHaveLength(1);

  await act(async () => rerender({ displayListQueries: queries(false, placeholder) }));
  expect(scrolls).toHaveLength(1);

  await act(async () => rerender({ displayListQueries: queries(true, match) }));
  expect(scrolls).toHaveLength(2);
  const matchTop = 6000 - scroller.scrollTop + 900;
  expect(matchTop).toBeGreaterThanOrEqual(0);
  expect(matchTop).toBeLessThanOrEqual(400);

  await act(async () => rerender({ displayListQueries: queries(true, match) }));
  expect(scrolls).toHaveLength(2);
});

test('a scroll to an unbuilt page follows its position to the page that holds it', async () => {
  const { scroller, host, scrolls } = pagedDom();
  const guess = { pageIndex: 6, x: 20, y: 20, width: 0, height: 0 };
  const next = { pageIndex: 8, x: 20, y: 20, width: 0, height: 0 };
  const match = { pageIndex: 8, x: 20, y: 500, width: 0, height: 16 };
  const { result, rerender } = renderHook(
    ({ displayListQueries }) =>
      usePagedScrollApi({
        pagesContainerRef: { current: host },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => scroller,
        displayListQueries,
      }),
    { initialProps: { displayListQueries: queries([0, 1, 2, 3, 4, 5], guess) } }
  );
  await act(async () => result.current.scrollToPositionImpl(700));
  await act(async () => rerender({ displayListQueries: queries([0, 1, 2, 3, 4, 5, 6, 7], next) }));
  expect(scrolls).toHaveLength(2);
  await act(async () => rerender({ displayListQueries: queries(true, match) }));
  expect(scrolls).toHaveLength(3);
  const matchTop = 8000 - scroller.scrollTop + 500;
  expect(matchTop).toBeGreaterThanOrEqual(0);
  expect(matchTop).toBeLessThanOrEqual(400);
});

test('a user scroll or a later navigation drops the pending refinement', async () => {
  const { scroller, host, scrolls } = pagedDom();
  const placeholder = { pageIndex: 6, x: 20, y: 20, width: 0, height: 0 };
  const match = { pageIndex: 6, x: 20, y: 900, width: 0, height: 16 };
  const { result, rerender } = renderHook(
    ({ displayListQueries }) =>
      usePagedScrollApi({
        pagesContainerRef: { current: host },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => scroller,
        displayListQueries,
      }),
    { initialProps: { displayListQueries: queries(false, placeholder) } }
  );
  await act(async () => result.current.scrollToPositionImpl(500));
  scroller.dispatchEvent(new Event('wheel'));
  await act(async () => rerender({ displayListQueries: queries(true, match) }));
  expect(scrolls).toHaveLength(1);

  const onlyFirst = {
    ...queries(false, placeholder),
    anchorRect: (position: number) => (position === 500 ? placeholder : null),
  } as unknown as DisplayListQueries;
  await act(async () => rerender({ displayListQueries: onlyFirst }));
  await act(async () => result.current.scrollToPositionImpl(500));
  expect(result.current.revealPositionImpl(600)).toBe('unsupported');
  await act(async () => rerender({ displayListQueries: queries(true, match) }));
  expect(scrolls).toHaveLength(2);
});
