import { expect, spyOn, test } from 'bun:test';
import type { DisplayPage } from '../layout/render/displayList';
import { createDisplayListQueries, type DisplayListQueries } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import {
  RESIDENT_QUERY_STORE_PAGES,
  releaseResidentDisplayListQueries,
  residentDisplayListQueries,
} from './residentDisplayListQueries';

function textPage(pageIndex: number): DisplayPage {
  return {
    pageIndex,
    width: 816,
    height: 1056,
    primitives: [{
      kind: 'text', text: `Page ${pageIndex}`, x: 100, baselineY: 100, width: 80, font: '400 16px Arial',
      color: '#000000', docStart: pageIndex * 20 + 1, docEnd: pageIndex * 20 + 8,
    } as DisplayPage['primitives'][number]],
  };
}

async function countedStore() {
  const engine = await loadRustDisplayListQueryEngine();
  const original = { open: engine.openDisplayList, update: engine.updateDisplayList, close: engine.closeDisplayList };
  const calls = { open: 0, close: 0 };
  const parsed = new Map<number, Set<number>>();
  engine.openDisplayList = (json: string) => {
    calls.open += 1;
    const handle = original.open!.call(engine, json);
    parsed.set(handle, new Set((JSON.parse(json) as { pages: DisplayPage[] }).pages
      .flatMap((page, index) => page.primitives.length ? [index] : [])));
    return handle;
  };
  engine.updateDisplayList = (handle: number, json: string) => {
    original.update!.call(engine, handle, json);
    const update = JSON.parse(json) as { keep?: boolean; replace?: [number, DisplayPage][]; reuse?: [number, number][] };
    const before = parsed.get(handle) ?? new Set<number>();
    const after = update.keep ? before : new Set((update.reuse ?? []).flatMap(([next, from]) => before.has(from) ? [next] : []));
    for (const [index, page] of update.replace ?? []) {
      if (page.primitives.length) after.add(index);
      else after.delete(index);
    }
    parsed.set(handle, after);
  };
  engine.closeDisplayList = (handle: number) => {
    calls.close += 1;
    parsed.delete(handle);
    original.close!.call(engine, handle);
  };
  const restore = () => {
    engine.openDisplayList = original.open;
    engine.updateDisplayList = original.update;
    engine.closeDisplayList = original.close;
  };
  return { calls, parsed, restore };
}

test('a session keeps one query engine per display list, builds the next from it and closes it on release', async () => {
  const store = await countedStore();
  try {
    const owner = {};
    const pages = [0, 1, 2].map(textPage);
    const list = { pages };
    const first = residentDisplayListQueries(owner, list);
    first.rangeRects(2, 5);
    expect(residentDisplayListQueries(owner, list)).toBe(first);
    const next = residentDisplayListQueries(owner, { pages: [...pages] });
    expect(next).not.toBe(first);
    expect(next.rangeRects(2, 5)).toEqual(first.rangeRects(2, 5));
    expect(store.calls.open).toBe(1);
    expect([...store.parsed.values()].map((indices) => [...indices])).toEqual([[0]]);
    releaseResidentDisplayListQueries(owner);
    expect(store.calls.close).toBe(1);
    expect(store.parsed.size).toBe(0);
  } finally { store.restore(); }
});

test('a session parses at most its page cap of a long document', async () => {
  const store = await countedStore();
  try {
    const owner = {};
    const pages = Array.from({ length: RESIDENT_QUERY_STORE_PAGES * 3 }, (_, index) => textPage(index));
    let list = { pages };
    for (let index = 0; index < pages.length; index += 1) {
      list = { pages: [...pages] };
      expect(residentDisplayListQueries(owner, list).rangeRects(index * 20 + 2, index * 20 + 5)).toHaveLength(1);
      const held = [...store.parsed.values()];
      expect(held).toHaveLength(1);
      expect(held[0]!.size).toBeLessThanOrEqual(RESIDENT_QUERY_STORE_PAGES);
      expect(held[0]!.has(index)).toBe(true);
    }
    releaseResidentDisplayListQueries(owner);
    expect(store.parsed.size).toBe(0);
  } finally { store.restore(); }
});

function malformedPage(pageIndex: number): DisplayPage {
  const page = textPage(pageIndex);
  return { ...page, primitives: page.primitives.map((primitive) => ({ ...primitive, table: { tableId: 't0' } }) as DisplayPage['primitives'][number]) };
}

test('a session closes a query engine whose read failed and answers the next query from a fresh one', async () => {
  await loadRustDisplayListQueryEngine();
  const built: { queries: DisplayListQueries; donor: DisplayListQueries | null | undefined; disposed: boolean }[] = [];
  const create: typeof createDisplayListQueries = (...args) => {
    const queries = createDisplayListQueries(...args);
    const entry = { queries, donor: args[2], disposed: false };
    const dispose = queries.dispose;
    queries.dispose = () => {
      entry.disposed = true;
      dispose();
    };
    built.push(entry);
    return queries;
  };
  const owner = {};
  const list = { pages: [textPage(0), malformedPage(1)] };
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const failed = residentDisplayListQueries(owner, list, create);
    expect(failed.rangeRects(22, 25)).toEqual([]);
    expect(failed.sourceState().status).toBe('error');
    const again = residentDisplayListQueries(owner, list, create);
    expect(again).not.toBe(failed);
    expect(built.map((entry) => [entry.donor ?? null, entry.disposed])).toEqual([[null, true], [null, false]]);
    expect(again.rangeRects(22, 25)).toEqual([]);
    expect(again.sourceState().status).toBe('error');
    const healthy = residentDisplayListQueries(owner, list, create);
    expect(healthy).not.toBe(again);
    expect(built.map((entry) => [entry.donor ?? null, entry.disposed])).toEqual([[null, true], [null, true], [null, false]]);
    expect(healthy.rangeRects(2, 5)).toHaveLength(1);
    expect(healthy.sourceState().status).toBe('ready');
    expect(residentDisplayListQueries(owner, list, create)).toBe(healthy);
    expect(built).toHaveLength(3);
  } finally {
    warn.mockRestore();
    releaseResidentDisplayListQueries(owner);
  }
});

test('a session answers a hit on a page of its list after eight later lists opened in the same store', async () => {
  await loadRustDisplayListQueryEngine();
  const owner = {};
  const list = { pages: [textPage(0), textPage(1)] };
  const others = Array.from({ length: 8 }, () => ({}));
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const first = residentDisplayListQueries(owner, list);
    expect(first.rangeRects(2, 5)).toHaveLength(1);
    for (const other of others) {
      expect(residentDisplayListQueries(other, { pages: [textPage(0)] }).rangeRects(2, 5)).toHaveLength(1);
    }
    const hit = first.hitTestRegions(1, 120, 95);
    expect(hit?.region).toBe('body');
    expect(hit?.pos).toBeGreaterThanOrEqual(21);
    expect(hit?.pos).toBeLessThanOrEqual(28);
    expect(first.sourceState().status).toBe('ready');
    expect(residentDisplayListQueries(owner, list)).toBe(first);
  } finally {
    warn.mockRestore();
    for (const other of others) releaseResidentDisplayListQueries(other);
    releaseResidentDisplayListQueries(owner);
  }
});
