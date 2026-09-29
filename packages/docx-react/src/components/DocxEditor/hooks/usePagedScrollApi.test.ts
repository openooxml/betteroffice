import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type {
  DisplayListQueries,
  DisplayListRect,
  DisplayPage,
} from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { usePagedScrollApi } from './usePagedScrollApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const layout = (pages: number, partial = false) =>
  ({ pageSize: { w: 1, h: 1 }, pages: Array(pages).fill({}), ...(partial ? { partial } : {}) }) as Layout;
const queries = (pages: number) =>
  ({ pageCount: () => pages, pageBounds: () => null }) as unknown as DisplayListQueries;

type Props = { layout: Layout; queries: DisplayListQueries; session?: YrsSession };

function scrollApi() {
  const scrolled: number[] = [];
  let page = 0;
  const navigation = { epoch: 0 };
  const hook = renderHook(
    (props: Props) =>
      usePagedScrollApi({
        pagesContainerRef: { current: null },
        yrsInputRef: { current: null },
        yrsSession: props.session ?? null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => null,
        displayListQueries: props.queries,
        layout: props.layout,
        onNavigationIntent: () => scrolled.push(page),
        navigationEpoch: () => navigation.epoch,
      }),
    { initialProps: { layout: layout(7, true), queries: queries(7) } as Props }
  );
  const scrollTo = (target: number) => {
    page = target;
    act(() => hook.result.current.scrollToPageImpl(target));
  };
  return { hook, scrolled, scrollTo, navigation };
}

test('a page past a partial layout waits for the full one', () => {
  const { hook, scrolled, scrollTo } = scrollApi();
  scrollTo(20);
  expect(scrolled).toEqual([]);
  // The full layout lands before its pages are displayed.
  hook.rerender({ layout: layout(29), queries: queries(7) });
  expect(scrolled).toEqual([]);
  hook.rerender({ layout: layout(29), queries: queries(29) });
  expect(scrolled).toEqual([20]);
});

test('a page past the full layout is dropped', () => {
  const { hook, scrolled, scrollTo } = scrollApi();
  scrollTo(40);
  hook.rerender({ layout: layout(29), queries: queries(29) });
  hook.rerender({ layout: layout(45), queries: queries(45) });
  expect(scrolled).toEqual([]);
});

test('a newer navigation or another session drops a waiting page', () => {
  const navigated = scrollApi();
  navigated.scrollTo(20);
  navigated.navigation.epoch += 1;
  navigated.hook.rerender({ layout: layout(29), queries: queries(29) });
  expect(navigated.scrolled).toEqual([]);

  const reopened = scrollApi();
  reopened.scrollTo(20);
  reopened.hook.rerender({ layout: layout(29), queries: queries(29), session: {} as YrsSession });
  expect(reopened.scrolled).toEqual([]);
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

function unbuiltQueries(built: boolean | number[], anchor: DisplayListRect): DisplayListQueries {
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
    { initialProps: { displayListQueries: unbuiltQueries(false, placeholder) } }
  );
  await act(async () => result.current.scrollToPositionImpl(500));
  expect(scrolls).toHaveLength(1);

  await act(async () => rerender({ displayListQueries: unbuiltQueries(false, placeholder) }));
  expect(scrolls).toHaveLength(1);

  await act(async () => rerender({ displayListQueries: unbuiltQueries(true, match) }));
  expect(scrolls).toHaveLength(2);
  const matchTop = 6000 - scroller.scrollTop + 900;
  expect(matchTop).toBeGreaterThanOrEqual(0);
  expect(matchTop).toBeLessThanOrEqual(400);

  await act(async () => rerender({ displayListQueries: unbuiltQueries(true, match) }));
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
    { initialProps: { displayListQueries: unbuiltQueries([0, 1, 2, 3, 4, 5], guess) } }
  );
  await act(async () => result.current.scrollToPositionImpl(700));
  await act(async () =>
    rerender({ displayListQueries: unbuiltQueries([0, 1, 2, 3, 4, 5, 6, 7], next) })
  );
  expect(scrolls).toHaveLength(2);
  await act(async () => rerender({ displayListQueries: unbuiltQueries(true, match) }));
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
    { initialProps: { displayListQueries: unbuiltQueries(false, placeholder) } }
  );
  await act(async () => result.current.scrollToPositionImpl(500));
  scroller.dispatchEvent(new Event('wheel'));
  await act(async () => rerender({ displayListQueries: unbuiltQueries(true, match) }));
  expect(scrolls).toHaveLength(1);

  const onlyFirst = {
    ...unbuiltQueries(false, placeholder),
    anchorRect: (position: number) => (position === 500 ? placeholder : null),
  } as unknown as DisplayListQueries;
  await act(async () => rerender({ displayListQueries: onlyFirst }));
  await act(async () => result.current.scrollToPositionImpl(500));
  expect(result.current.revealPositionImpl(600)).toBe('unsupported');
  await act(async () => rerender({ displayListQueries: unbuiltQueries(true, match) }));
  expect(scrolls).toHaveLength(2);
});
