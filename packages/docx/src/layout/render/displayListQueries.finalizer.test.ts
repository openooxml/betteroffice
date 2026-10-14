import { describe, expect, test } from 'bun:test';
import type { DisplayPage } from './displayList';
import type { RustDisplayListQueryEngine } from './rustDisplayList';

const registered = new Set<object>();
const NativeRegistry = globalThis.FinalizationRegistry;

class CountingRegistry<T> extends NativeRegistry<T> {
  override register(target: WeakKey, held: T, token?: WeakKey): void {
    super.register(target, held, token);
    if (token) registered.add(token);
  }

  override unregister(token: WeakKey): boolean {
    registered.delete(token);
    return super.unregister(token);
  }
}

globalThis.FinalizationRegistry = CountingRegistry as typeof FinalizationRegistry;
const { createDisplayListQueries, endDisplayListQueriesLine } = await import('./displayListQueries');
globalThis.FinalizationRegistry = NativeRegistry;

function page(pageIndex: number): DisplayPage {
  return { pageIndex, width: 100, height: 100, primitives: [] };
}

function fakeEngine() {
  const open = new Set<number>();
  let nextHandle = 1;
  const engine: RustDisplayListQueryEngine = {
    hitTestRegionsJson: () => 'null',
    verticalMoveJson: () => 'null',
    rangeRectsJson: () => '[]',
    hasDisplayListSession: () => true,
    openDisplayList: () => {
      open.add(nextHandle);
      return nextHandle++;
    },
    closeDisplayList: (handle: number) => {
      open.delete(handle);
    },
    updateDisplayList: () => {},
    hasDisplayListUpdate: () => true,
    rangeRectsByHandle: () => '[]',
  };
  return { engine, open };
}

describe('display-list handle finalization', () => {
  test('registers only a facade that holds a handle', () => {
    const { engine } = fakeEngine();
    registered.clear();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    expect(registered.size).toBe(0);

    second.rangeRects(0, 1);
    expect(registered.size).toBe(1);

    second.dispose();
    first.dispose();
    expect(registered.size).toBe(0);
  });

  test('an adopted handle moves its registration to the adopter', () => {
    const { engine, open } = fakeEngine();
    registered.clear();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.prime();
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    second.rangeRects(0, 1);
    expect(open.size).toBe(1);
    expect(registered.size).toBe(1);

    first.dispose();
    expect(open.size).toBe(1);
    second.dispose();
    expect(open.size).toBe(0);
    expect(registered.size).toBe(0);
  });

  test('a handle dropped after a failed by-handle query is no longer registered', () => {
    const { engine, open } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      throw new Error('stale handle');
    };
    registered.clear();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    const warn = console.warn;
    console.warn = () => {};
    try {
      queries.rangeRects(0, 1);
    } finally {
      console.warn = warn;
    }
    expect(open.size).toBe(0);
    expect(registered.size).toBe(0);
  });

  test('ending a line closes and unregisters every handle its facades hold', () => {
    const { engine, open } = fakeEngine();
    registered.clear();
    const line = {};
    const first = createDisplayListQueries({ pages: [page(0)] }, engine, null, line);
    first.prime();
    const second = createDisplayListQueries({ pages: [page(0)] }, engine, null, line);
    second.rangeRects(0, 1);
    expect(open.size).toBe(2);
    expect(registered.size).toBe(2);

    endDisplayListQueriesLine(line);
    expect(open.size).toBe(0);
    expect(registered.size).toBe(0);
    expect(second.rangeRects(0, 1)).toEqual([]);
    expect(open.size).toBe(0);
  });
});
