import { describe, expect, test } from 'bun:test';
import { createDisplayListQueries } from './displayListQueries';
import type { DisplayPage } from './displayList';
import type { RustDisplayListQueryEngine } from './rustDisplayList';

function page(pageIndex: number): DisplayPage {
  return { pageIndex, width: 100, height: 100, primitives: [] };
}

function fakeEngine() {
  const calls = { open: 0, update: 0, close: 0, rangeByHandle: 0, rangeJson: 0 };
  let nextHandle = 1;
  const engine: RustDisplayListQueryEngine = {
    hitTestRegionsJson: () => 'null',
    rangeRectsJson: () => {
      calls.rangeJson += 1;
      return '[]';
    },
    hasDisplayListSession: () => true,
    openDisplayList: () => {
      calls.open += 1;
      return nextHandle++;
    },
    closeDisplayList: () => {
      calls.close += 1;
    },
    updateDisplayList: () => {
      calls.update += 1;
    },
    hasDisplayListUpdate: () => true,
    rangeRectsByHandle: () => {
      calls.rangeByHandle += 1;
      return '[]';
    },
  };
  return { engine, calls };
}

describe('createDisplayListQueries handle lifecycle', () => {
  test('opens the session handle lazily, on the first query', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    expect(calls.open).toBe(0);
    queries.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.rangeByHandle).toBe(1);
  });

  test('prime() acquires the handle without a query', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    queries.prime();
    expect(calls.open).toBe(1);
    expect(calls.rangeByHandle).toBe(0);
    queries.prime();
    expect(calls.open).toBe(1);
  });

  test('adoption chains across unqueried generations as one page-delta', () => {
    const { engine, calls } = fakeEngine();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    const third = createDisplayListQueries({ pages: [shared] }, engine, second);
    expect(calls.open).toBe(1);
    expect(calls.update).toBe(0);
    third.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.update).toBe(1);
  });

  test('superseded generations fall back to JSON-arg queries, never reopening', () => {
    const { engine, calls } = fakeEngine();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    second.rangeRects(0, 1);
    expect(calls.update).toBe(1);
    first.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.rangeJson).toBe(1);
  });
});
