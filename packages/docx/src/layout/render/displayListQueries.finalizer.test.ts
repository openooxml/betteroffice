import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import type { DisplayPage } from './displayList';
import { createDisplayListQueries, endDisplayListQueriesLine } from './displayListQueries';
import type { RustDisplayListQueryEngine } from './rustDisplayList';

// Tokens of live finalizer registrations whose target is a query facade.
const registered = new Set<WeakKey>();
const registry = FinalizationRegistry.prototype;
const { register, unregister } = registry;

beforeAll(() => {
  registry.register = function (this: FinalizationRegistry<unknown>, target, held, token) {
    const facade = target as { rangeRects?: unknown; dispose?: unknown };
    if (token && typeof facade.rangeRects === 'function' && typeof facade.dispose === 'function') {
      registered.add(token);
    }
    register.call(this, target, held, token);
  };
  registry.unregister = function (this: FinalizationRegistry<unknown>, token) {
    registered.delete(token);
    return unregister.call(this, token);
  };
});

afterAll(() => {
  registry.register = register;
  registry.unregister = unregister;
});

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
    expect(registered.size).toBe(1);
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

  test('a query method kept past its collected facade opens no handle', async () => {
    const { engine, open } = fakeEngine();
    const line = {};
    const { prime, facade } = (() => {
      const queries = createDisplayListQueries({ pages: [page(0)] }, engine, null, line);
      return { prime: queries.prime, facade: new WeakRef(queries) };
    })();
    for (let i = 0; i < 20 && facade.deref(); i++) {
      await new Promise<void>((resolve) =>
        setTimeout(() => {
          Bun.gc(true);
          resolve();
        }, 0)
      );
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    expect(facade.deref()).toBeUndefined();

    prime();
    endDisplayListQueriesLine(line);
    expect(open.size).toBe(0);
  });
});
