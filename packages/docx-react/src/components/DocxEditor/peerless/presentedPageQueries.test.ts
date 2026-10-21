import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import * as display from '@betteroffice/docx/layout/render';
import { createPresentedPageQueries } from './presentedPageQueries';

const cleanups: Array<() => void> = [];
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });

function setup() {
  const open = mock(() => 1);
  const update = mock((_handle: number, _json: string) => {});
  const close = mock(() => {});
  const engine: display.RustDisplayListQueryEngine = {
    hasDisplayListSession: () => true, hasDisplayListUpdate: () => true,
    openDisplayList: open, updateDisplayList: update, closeDisplayList: close,
    hitTestRegionsByHandle: () => '{"region":"body","pos":12}',
    rangeRectsByHandle: () => '[{"pageIndex":0,"x":47,"y":6,"width":1,"height":18}]',
    hitTestRegionsJson: () => { throw new Error('Unexpected JSON fallback'); },
    rangeRectsJson: () => { throw new Error('Unexpected JSON fallback'); },
    verticalMoveJson: () => 'null',
  };
  const loaded = spyOn(display, 'loadedRustDisplayListQueryEngine').mockReturnValue(engine);
  cleanups.push(() => loaded.mockRestore());
  const cache = createPresentedPageQueries();
  cleanups.push(() => cache.dispose());
  const page: display.DisplayPage = { pageIndex: 17, width: 200, height: 40, primitives: [{
    kind: 'text', text: 'body', x: 0, baselineY: 20, width: 100, font: '12px Arial', color: '#000', docStart: 10, docEnd: 20,
  }] };
  return { cache, page, open, update, close, loaded, engine };
}

test('a presented page warms once and eight later clicks reuse its parsed store', () => {
  const { cache, page, open, update, close } = setup();
  cache.retain([page]);
  const warmed = cache.get(page)!;
  warmed.hitTestRegions(0, 0, 0);
  expect(open).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalledTimes(1);
  expect(JSON.parse(update.mock.calls[0]![1]).replace).toEqual([[0, page]]);
  for (let click = 0; click < 8; click++) {
    const queries = cache.get(page)!;
    expect(queries).toBe(warmed);
    expect(queries.hitTestRegions(0, click, 10)?.pos).toBe(12);
    expect(queries.caretRect(12)?.height).toBe(18);
  }
  expect(open).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalledTimes(1);
  expect(close).not.toHaveBeenCalled();
  cache.retain([]);
  expect(close).toHaveBeenCalledTimes(1);
});

test('only presented pages are retained and a new engine replaces the cached source', () => {
  const { cache, page, close, loaded, engine } = setup();
  cache.retain([page]);
  const first = cache.get(page)!;
  first.hitTestRegions(0, 0, 0);
  cache.retain([{ ...page }]);
  expect(close).toHaveBeenCalledTimes(1);
  expect(cache.get(page)).toBeNull();
  cache.retain([page]);
  const second = cache.get(page)!;
  second.hitTestRegions(0, 0, 0);
  loaded.mockReturnValue({ ...engine });
  expect(cache.get(page)).not.toBe(second);
  expect(close).toHaveBeenCalledTimes(2);
});

test('twelve presented pages share one handle and caret queries stay on the clicked page', () => {
  const { cache, page, open, update, engine } = setup();
  const pages = Array.from({ length: 12 }, (_, pageIndex) => ({ ...page, pageIndex }));
  const ranges = mock((_handle: number, _from: number, _to: number, first: number, _last: number) =>
    JSON.stringify([{ pageIndex: first, x: first * 10, y: 6, width: 7, height: 18 }]));
  engine.hasRangeRectsOnPages = () => true;
  engine.rangeRectsOnPagesJson = () => { throw new Error('Unexpected JSON fallback'); };
  engine.rangeRectsOnPagesByHandle = ranges;
  cache.retain(pages);
  for (const presented of pages) cache.get(presented)!.hitTestRegions(0, 0, 0);
  expect(open).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalledTimes(12);
  for (let click = 0; click < 8; click++) {
    expect(cache.get(pages[11]!)!.caretRect(12)).toEqual({ pageIndex: 0, x: 110, y: 6, width: 0, height: 18 });
  }
  expect(ranges).toHaveBeenCalledWith(1, 12, 13, 11, 11);
  expect(open).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalledTimes(12);
});

test('an owned position shift invalidates the presented page store', () => {
  const { cache, page, close, update } = setup();
  const delta = (frameEpoch: number, operations: display.FramePageOperation[]): display.DecodedFrameDelta => ({
    protocolVersion: display.FRAME_DELTA_VERSION, full: frameEpoch === 1, docEpoch: frameEpoch,
    layoutEpoch: frameEpoch, frameEpoch, baseFrameEpoch: frameEpoch - 1,
    pageCount: 1, bytes: new Uint8Array(), operations,
  });
  page.pageIndex = 0;
  const retained = display.applyFrameDeltaOwned(null, delta(1, [{
    kind: 'upsert', pageId: 1n, pageIndex: 0, fingerprint: 1n,
    primitiveIds: new BigUint64Array([1n]), page,
  }]));
  const presented = retained.pages[0]!.page;
  cache.retain([presented]);
  const before = cache.get(presented)!;
  before.hitTestRegions(0, 0, 0);
  display.applyFrameDeltaOwned(retained, delta(2, [{ kind: 'shift-range', pageIndex: 0, pageId: 1n, count: 1, delta: 2, salt: 3n }]));
  const after = cache.get(presented)!;
  expect(after).not.toBe(before);
  after.hitTestRegions(0, 0, 0);
  expect(close).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalledTimes(2);
  expect(JSON.parse(update.mock.calls[1]![1]).replace[0][1].primitives[0].docStart).toBe(12);
});

for (const pageCount of [1, 2]) {
  for (const direction of ['rtl', 'bidiLevel']) {
    test(`${pageCount} presented pages decline a provisional caret in a ${direction} run or paragraph`, () => {
      const { cache, page, engine } = setup();
      page.primitives = [{ ...page.primitives[0]!, fragmentDocStart: 5, fragmentDocEnd: 25,
        ...(direction === 'rtl' ? { rtl: true } : { bidiLevel: 1 }) }];
      engine.hasRangeRectsOnPages = () => true;
      engine.rangeRectsOnPagesJson = () => { throw new Error('Unexpected JSON fallback'); };
      engine.rangeRectsOnPagesByHandle = () => '[{"pageIndex":0,"x":47,"y":6,"width":7,"height":18}]';
      cache.retain(Array.from({ length: pageCount }, (_, index) => index === 0 ? page : { ...page, pageIndex: index }));
      const queries = cache.get(page)!;
      expect(queries.hitTestRegions(0, 47, 6)?.pos).toBe(12);
      expect(queries.caretRect(12)).toBeNull();
      expect(queries.caretRect(20)).toBeNull();
      expect(queries.caretRect(5)).toBeNull();
      expect(queries.caretRect(30)?.height).toBe(18);
    });
  }
}
