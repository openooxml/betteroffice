import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
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

function scrollApi() {
  const scrolled: number[] = [];
  let page = 0;
  const hook = renderHook(
    (props: { layout: Layout; queries: DisplayListQueries }) =>
      usePagedScrollApi({
        pagesContainerRef: { current: null },
        yrsInputRef: { current: null },
        yrsSession: null,
        yrsLocToDisplayPosition: () => null,
        getScrollContainer: () => null,
        displayListQueries: props.queries,
        layout: props.layout,
        onNavigationIntent: () => scrolled.push(page),
      }),
    { initialProps: { layout: layout(7, true), queries: queries(7) } }
  );
  const scrollTo = (target: number) => {
    page = target;
    act(() => hook.result.current.scrollToPageImpl(target));
  };
  return { hook, scrolled, scrollTo };
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
