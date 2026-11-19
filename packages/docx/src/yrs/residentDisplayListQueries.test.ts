import { expect, test } from 'bun:test';
import type { DisplayPage } from '../layout/render/displayList';
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
