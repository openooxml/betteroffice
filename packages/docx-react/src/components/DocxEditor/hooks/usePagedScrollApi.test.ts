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
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '6';
  host.className = 'canvas-pages';
  host.append(page);
  scroller.append(host);
  document.body.append(scroller);
  Object.defineProperty(scroller, 'clientHeight', { value: 400 });
  scroller.getBoundingClientRect = () =>
    ({ top: 0, bottom: 400, height: 400, left: 0, right: 800, width: 800 }) as DOMRect;
  page.getBoundingClientRect = () =>
    ({
      top: 6000 - scroller.scrollTop,
      bottom: 7000 - scroller.scrollTop,
      height: 1000,
      left: 0,
      right: 800,
      width: 800,
    }) as DOMRect;
  const scrolls: number[] = [];
  scroller.scrollTo = ((options: ScrollToOptions) => {
    scrolls.push(options.top ?? 0);
    scroller.scrollTop = options.top ?? 0;
  }) as typeof scroller.scrollTo;
  return { scroller, host, scrolls };
}

function queries(built: boolean, anchor: DisplayListRect): DisplayListQueries {
  const page: DisplayPage = { pageIndex: 6, width: 800, height: 1000, primitives: [] };
  if (!built) Object.assign(page, { unbuilt: true, positionSpan: [400, 600] });
  const pages = Array.from({ length: 7 }, (_, pageIndex) =>
    pageIndex === 6 ? page : { pageIndex, width: 800, height: 1000, primitives: [] }
  );
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
