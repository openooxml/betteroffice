import { describe, expect, test } from 'bun:test';
import {
  createDisplayListQueries,
  endDisplayListQueriesLine,
  isDisplayListQuerySourceDead,
  onDisplayListQuerySourceFailure,
  type DisplayListQueries,
  type DisplayListRegionHit,
} from './displayListQueries';
import type { DisplayPage } from './displayList';
import type { RustDisplayListQueryEngine } from './rustDisplayList';

function page(pageIndex: number): DisplayPage {
  return { pageIndex, width: 100, height: 100, primitives: [] };
}

function fakeEngine() {
  const calls = {
    open: 0,
    update: 0,
    close: 0,
    rangeByHandle: 0,
    rangeJson: 0,
    verticalByHandle: 0,
  };
  let nextHandle = 1;
  const engine: RustDisplayListQueryEngine = {
    hitTestRegionsJson: () => 'null',
    verticalMoveJson: () => 'null',
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
    verticalMoveByHandle: () => {
      calls.verticalByHandle += 1;
      return '{"position":2,"goalX":24}';
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

  test('a superseded generation answers from the live one, never reopening or serialising', () => {
    const { engine, calls } = fakeEngine();
    const line = {};
    // the store answers with the generation it holds, as the live layout would
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      return `[{"pageIndex":0,"x":${calls.update},"y":0,"width":1,"height":1}]`;
    };
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine, null, line);
    first.rangeRects(0, 1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first, line);
    // still holding the handle, the old generation answers itself
    expect(first.rangeRects(0, 1)[0]?.x).toBe(0);
    second.rangeRects(0, 1);
    const third = createDisplayListQueries({ pages: [shared, page(1)] }, engine, second, line);
    expect(third.rangeRects(0, 1)[0]?.x).toBe(2);

    expect(first.rangeRects(0, 1)[0]?.x).toBe(2);
    expect(second.caretRect(0)?.x).toBe(2);
    expect(calls.open).toBe(1);
    expect(calls.rangeJson).toBe(0);
  });

  const rect = (x: number) => ({ pageIndex: 1, x, y: 0, width: 1, height: 1 });
  const hit = (pos: number): DisplayListRegionHit => ({ region: 'body', pos, target: 'text' });

  test('a superseded generation without a line queries its own list after handle adoption', () => {
    const { engine, calls } = fakeEngine();
    const queriedHandles: number[] = [];
    const jsonQueries: string[] = [];
    let hitJson = 0;
    engine.rangeRectsByHandle = (handle) => {
      calls.rangeByHandle += 1;
      queriedHandles.push(handle);
      return JSON.stringify([rect(300)]);
    };
    engine.hitTestRegionsByHandle = (handle) => {
      queriedHandles.push(handle);
      return JSON.stringify(hit(300));
    };
    engine.rangeRectsJson = (json) => {
      calls.rangeJson += 1;
      jsonQueries.push(json);
      return JSON.stringify([rect(JSON.parse(json).pages[1].width)]);
    };
    engine.hitTestRegionsJson = (json) => {
      hitJson += 1;
      jsonQueries.push(json);
      return JSON.stringify(hit(JSON.parse(json).pages[1].width));
    };
    const shared = page(0);
    const firstList = { pages: [shared, { ...page(1), width: 200 }] };
    const first = createDisplayListQueries(firstList, engine);
    first.prime();
    const second = createDisplayListQueries(
      { pages: [shared, { ...page(1), width: 300 }] },
      engine,
      first
    );
    expect(second.rangeRects(0, 1)).toEqual([rect(300)]);
    expect(calls.update).toBe(1);
    expect(second.hitTestRegions(1, 1, 1)).toEqual(hit(300));

    expect(first.rangeRects(0, 1)).toEqual([rect(200)]);
    expect(first.hitTestRegions(1, 1, 1)).toEqual(hit(200));
    expect(queriedHandles).toEqual([1, 1]);
    expect(jsonQueries).toEqual([JSON.stringify(firstList), JSON.stringify(firstList)]);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(1);
    expect(hitJson).toBe(1);
    expect(calls.open).toBe(1);
  });

  test('a document line rejoins after a gap in the previous facade chain', () => {
    const { engine, calls } = fakeEngine();
    const queriedHandles: number[] = [];
    const jsonQueries: string[] = [];
    engine.rangeRectsByHandle = (handle) => {
      queriedHandles.push(handle);
      return JSON.stringify([rect(handle)]);
    };
    engine.hitTestRegionsByHandle = (handle) => {
      queriedHandles.push(handle);
      return JSON.stringify(hit(handle));
    };
    engine.rangeRectsJson = () => {
      calls.rangeJson += 1;
      jsonQueries.push('rangeRects');
      return '[]';
    };
    engine.hitTestRegionsJson = () => {
      jsonQueries.push('hitTestRegions');
      return 'null';
    };
    const shared = page(0);
    const line = {};
    const first = createDisplayListQueries({ pages: [shared, page(1)] }, engine, null, line);
    first.prime();
    const second = createDisplayListQueries({ pages: [shared, page(1)] }, engine, first, line);
    expect(second.rangeRects(0, 1)).toEqual([rect(1)]);
    expect(calls.update).toBe(1);
    expect(second.hitTestRegions(1, 1, 1)).toEqual(hit(1));
    const third = createDisplayListQueries(
      { pages: [shared, { ...page(1), width: 300 }] },
      engine,
      null,
      line
    );
    third.prime();

    expect(first.rangeRects(0, 1)).toEqual([rect(2)]);
    expect(first.hitTestRegions(1, 1, 1)).toEqual(hit(2));
    expect(second.rangeRects(0, 1)).toEqual([rect(1)]);
    expect(queriedHandles).toEqual([1, 1, 2, 2, 1]);
    expect(jsonQueries).toEqual([]);
    expect(calls.rangeJson).toBe(0);
    expect(calls.open).toBe(2);
  });

  test('ending a document line silences superseded facades until the line rejoins', () => {
    const { engine, calls } = fakeEngine();
    const queriedHandles: number[] = [];
    const jsonQueries: string[] = [];
    engine.rangeRectsByHandle = (handle) => {
      queriedHandles.push(handle);
      return JSON.stringify([rect(handle)]);
    };
    engine.hitTestRegionsByHandle = (handle) => {
      queriedHandles.push(handle);
      return JSON.stringify(hit(handle));
    };
    engine.rangeRectsJson = () => {
      calls.rangeJson += 1;
      jsonQueries.push('rangeRects');
      return '[]';
    };
    engine.hitTestRegionsJson = () => {
      jsonQueries.push('hitTestRegions');
      return 'null';
    };
    const shared = page(0);
    const line = {};
    const first = createDisplayListQueries({ pages: [shared, page(1)] }, engine, null, line);
    first.prime();
    const second = createDisplayListQueries({ pages: [shared, page(1)] }, engine, first, line);
    second.prime();
    const third = createDisplayListQueries({ pages: [shared, page(1)] }, engine, second, line);
    expect(third.rangeRects(0, 1)).toEqual([rect(1)]);
    expect(calls.open).toBe(1);
    expect(calls.update).toBe(2);

    endDisplayListQueriesLine(line);
    for (const queries of [first, second]) {
      expect(queries.rangeRects(0, 1)).toEqual([]);
      expect(queries.hitTestRegions(1, 1, 1)).toBeNull();
    }
    expect(queriedHandles).toEqual([1]);
    expect(jsonQueries).toEqual([]);
    expect(calls.open).toBe(1);

    const rejoined = createDisplayListQueries({ pages: [shared, page(1)] }, engine, null, line);
    rejoined.prime();
    for (const queries of [first, second]) {
      expect(queries.rangeRects(0, 1)).toEqual([rect(2)]);
      expect(queries.hitTestRegions(1, 1, 1)).toEqual(hit(2));
    }
    expect(queriedHandles).toEqual([1, 2, 2, 2, 2]);
    expect(jsonQueries).toEqual([]);
    expect(calls.rangeJson).toBe(0);
    expect(calls.open).toBe(2);
  });

  const lineQueries = [
    {
      name: 'rangeRects',
      read: (queries: DisplayListQueries) => queries.rangeRects(0, 1),
      a: [rect(200)],
      b: [rect(300)],
      empty: [],
    },
    {
      name: 'anchorRect',
      read: (queries: DisplayListQueries) => queries.anchorRect(0),
      a: rect(200),
      b: rect(300),
      empty: null,
    },
    {
      name: 'hitTestRegions',
      read: (queries: DisplayListQueries) => queries.hitTestRegions(1, 1, 1),
      a: hit(200),
      b: hit(300),
      empty: null,
    },
  ];

  for (const query of lineQueries) {
    test(`a document line split ends ${query.name} forwarding for every superseded generation`, () => {
      const { engine, calls } = fakeEngine();
      const queriedHandles: number[] = [];
      const adoptedHandles: number[] = [];
      const jsonQueries: string[] = [];
      let storedWidth = 100;
      engine.updateDisplayList = (handle, json) => {
        calls.update += 1;
        const update = JSON.parse(json) as {
          reuse?: Array<[number, number]>;
          replace: Array<[number, DisplayPage]>;
        };
        if (update.reuse) adoptedHandles.push(handle);
        const changed = update.replace.find(([index]) => index === 1);
        if (changed) storedWidth = changed[1].width;
      };
      engine.rangeRectsByHandle = (handle) => {
        calls.rangeByHandle += 1;
        queriedHandles.push(handle);
        return JSON.stringify([rect(storedWidth)]);
      };
      engine.hitTestRegionsByHandle = (handle) => {
        queriedHandles.push(handle);
        return JSON.stringify(hit(storedWidth));
      };
      engine.rangeRectsJson = () => {
        calls.rangeJson += 1;
        jsonQueries.push('rangeRects');
        return '[]';
      };
      engine.hitTestRegionsJson = () => {
        jsonQueries.push('hitTestRegions');
        return 'null';
      };
      const read = (queries: DisplayListQueries) => {
        const handleStart = queriedHandles.length;
        const jsonStart = jsonQueries.length;
        const answer = query.read(queries);
        return {
          answer,
          handles: queriedHandles.slice(handleStart),
          json: jsonQueries.slice(jsonStart),
        };
      };
      const shared = page(0);
      const lineA = {};
      const lineB = {};
      const first = createDisplayListQueries({ pages: [shared, page(1)] }, engine, null, lineA);
      first.prime();
      const second = createDisplayListQueries(
        { pages: [shared, { ...page(1), width: 200 }] },
        engine,
        first,
        lineA
      );
      const currentA = read(second);
      const forwardedA = read(first);
      const replacement = createDisplayListQueries(
        { pages: [shared, { ...page(1), width: 300 }] },
        engine,
        second,
        lineB
      );
      const currentB = read(replacement);
      const staleA1 = read(second);
      const staleA0 = read(first);

      expect({
        currentA,
        forwardedA,
        currentB,
        staleA1,
        staleA0,
        opens: calls.open,
        adoptedHandles,
      }).toEqual({
        currentA: { answer: query.a, handles: [1], json: [] },
        forwardedA: { answer: query.a, handles: [1], json: [] },
        currentB: { answer: query.b, handles: [1], json: [] },
        staleA1: { answer: query.empty, handles: [], json: [] },
        staleA0: { answer: query.empty, handles: [], json: [] },
        opens: 1,
        adoptedHandles: [1, 1],
      });
    });
  }

  test('a superseded generation reads page metadata from the live layout too', () => {
    const { engine, calls } = fakeEngine();
    const shared = page(0);
    const line = {};
    const first = createDisplayListQueries({ pages: [shared] }, engine, null, line);
    first.rangeRects(0, 1);
    const landscape = { ...page(1), width: 140 };
    const second = createDisplayListQueries({ pages: [shared, landscape] }, engine, first, line);
    expect(first.pageCount()).toBe(1);
    second.rangeRects(0, 1);
    expect(calls.update).toBe(1);
    expect(first.pageCount()).toBe(2);
    expect(first.pageSize(1)).toEqual({ width: 140, height: 100 });
    expect(first.displayList).toBe(second.displayList);
  });

  test('a superseded generation whose successor is gone answers nothing', () => {
    const { engine, calls } = fakeEngine();
    const line = {};
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      return '[{"pageIndex":0,"x":1,"y":0,"width":1,"height":1}]';
    };
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine, null, line);
    first.rangeRects(0, 1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first, line);
    expect(second.rangeRects(0, 1)).toHaveLength(1);
    expect(calls.update).toBe(1);
    expect(first.rangeRects(0, 1)).toHaveLength(1);
    second.dispose();
    expect(first.rangeRects(0, 1)).toEqual([]);
    expect(first.hitTestRegions(0, 1, 1)).toBeNull();
    expect(first.anchorRect(0)).toBeNull();
    expect(calls.rangeJson).toBe(0);
  });

  test('routes vertical movement through the retained handle', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    expect(queries.verticalMove(1, 'down')).toEqual({ position: 2, goalX: 24 });
    expect(calls.open).toBe(1);
    expect(calls.verticalByHandle).toBe(1);
  });
});

function wasmTrap(): Error {
  const trap = new Error('unreachable executed');
  trap.name = 'RuntimeError';
  return trap;
}

describe('createDisplayListQueries wasm trap containment', () => {
  test('stops querying an instance whose wasm trapped', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
    expect(queries.sourceState().status).toBe('error');

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(queries.verticalMove(1, 'down')).toBeNull();
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
    expect(calls.verticalByHandle).toBe(0);
    expect(isDisplayListQuerySourceDead(engine)).toBe(true);
  });

  test('a returned Err still falls back to the JSON-arg path', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw new Error('unknown display-list handle 7');
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(1);
    expect(isDisplayListQuerySourceDead(engine)).toBe(false);
  });

  test('a later build over the dead instance never queries it', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    const openAfterTrap = calls.open;

    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    expect(second.rangeRects(0, 1)).toEqual([]);
    expect(calls.open).toBe(openAfterTrap);
    expect(calls.update).toBe(0);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
  });

  test('notifies failure listeners once so the host can rebuild', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const failures: Error[] = [];
    const unsubscribe = onDisplayListQuerySourceFailure((error) => failures.push(error));
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    queries.rangeRects(0, 1);
    queries.rangeRects(0, 1);
    unsubscribe();

    expect(failures).toHaveLength(1);
    expect(failures[0].name).toBe('RuntimeError');
  });

  test('a resident trap stops resident queries', () => {
    let residentCalls = 0;
    const resident = {
      displayHitTestRegionsJson: () => 'null',
      displayVerticalMoveJson: () => {
        residentCalls += 1;
        throw wasmTrap();
      },
      displayRangeRectsJson: () => {
        residentCalls += 1;
        return '[]';
      },
      displayRangeRectsRegionJson: () => '[]',
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, resident);

    expect(queries.verticalMove(1, 'down')).toBeNull();
    expect(residentCalls).toBe(1);
    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(residentCalls).toBe(1);
    expect(isDisplayListQuerySourceDead(resident)).toBe(true);
  });
});

describe('createDisplayListQueries lazy store pages', () => {
  function textPage(pageIndex: number, from: number, to: number): DisplayPage {
    return {
      pageIndex,
      width: 100,
      height: 100,
      primitives: [
        {
          kind: 'text',
          text: 'x',
          x: 0,
          baselineY: 10,
          width: 10,
          font: '10px serif',
          color: '#000',
          docStart: from,
          docEnd: to,
        } as DisplayPage['primitives'][number],
      ],
    };
  }

  function recordingEngine() {
    const opened: string[] = [];
    const updates: Array<{
      total: number;
      keep?: boolean;
      reuse?: number[][];
      replace: Array<[number, DisplayPage]>;
    }> = [];
    const { engine, calls } = fakeEngine();
    engine.openDisplayList = (json: string) => {
      calls.open += 1;
      opened.push(json);
      return 1;
    };
    engine.updateDisplayList = (_handle: number, json: string) => {
      calls.update += 1;
      updates.push(JSON.parse(json));
    };
    engine.hitTestRegionsByHandle = () => 'null';
    engine.rangeRectsRegionByHandle = () => '[]';
    engine.rangeRectsRegionJson = () => '[]';
    engine.hasRangeRectsRegion = () => true;
    return { engine, calls, opened, updates };
  }

  const list = () => ({ pages: [textPage(0, 1, 10), textPage(1, 11, 20), textPage(2, 21, 30)] });

  test('opens the store with page sizes only and parses the pages a range touches', () => {
    const { engine, opened, updates } = recordingEngine();
    const queries = createDisplayListQueries(list(), engine);
    queries.rangeRects(14, 15);
    expect(JSON.parse(opened[0]).pages.map((p: DisplayPage) => p.primitives.length)).toEqual([
      0, 0, 0,
    ]);
    expect(updates).toHaveLength(1);
    expect(updates[0].replace.map(([index]) => index)).toEqual([1]);
    expect(updates[0]).toMatchObject({ total: 3, keep: true });
    expect(updates[0].reuse).toBeUndefined();
    expect(updates[0].replace[0][1].primitives).toHaveLength(1);

    queries.rangeRects(15, 16);
    expect(updates).toHaveLength(1);
  });

  test('loads the hit page, a move’s neighbours, and every page for region queries', () => {
    const { engine, updates } = recordingEngine();
    const queries = createDisplayListQueries(list(), engine);
    queries.hitTestRegions(2, 5, 5);
    expect(updates.at(-1)!.replace.map(([index]) => index)).toEqual([2]);
    queries.verticalMove(5, 'down');
    expect(updates.at(-1)!.replace.map(([index]) => index)).toEqual([0, 1]);
    queries.hfRangeRects('header', 'rId1', 0, 1);
    expect(updates).toHaveLength(2);
  });

  test('a range spanning pages loads each of them', () => {
    const { engine, updates } = recordingEngine();
    const queries = createDisplayListQueries(list(), engine);
    queries.rangeRects(5, 25);
    expect(updates[0].replace.map(([index]) => index)).toEqual([0, 1, 2]);
  });
});

describe('createDisplayListQueries page load failures', () => {
  test('a trap while loading pages stops querying the instance', () => {
    const { engine, calls } = fakeEngine();
    engine.updateDisplayList = () => {
      throw wasmTrap();
    };
    const textPage: DisplayPage = {
      ...page(0),
      primitives: [
        { kind: 'rect', x: 0, y: 0, w: 1, h: 1, docStart: 1, docEnd: 2 } as DisplayPage['primitives'][number],
      ],
    };
    const queries = createDisplayListQueries({ pages: [textPage] }, engine);
    expect(queries.rangeRects(1, 2)).toEqual([]);
    expect(calls.rangeJson).toBe(0);
    expect(calls.rangeByHandle).toBe(0);
    expect(queries.sourceState().status).toBe('error');
  });

  test('an ordinary page load failure closes the handle and falls back to JSON', () => {
    const { engine, calls } = fakeEngine();
    engine.updateDisplayList = () => {
      throw new Error('update rejected');
    };
    const textPage: DisplayPage = {
      ...page(0),
      primitives: [
        { kind: 'rect', x: 0, y: 0, w: 1, h: 1, docStart: 1, docEnd: 2 } as DisplayPage['primitives'][number],
      ],
    };
    const warn = console.warn;
    console.warn = () => {};
    try {
      const queries = createDisplayListQueries({ pages: [textPage] }, engine);
      queries.rangeRects(1, 2);
      expect(calls.close).toBe(1);
      expect(calls.rangeJson).toBe(1);
      expect(calls.rangeByHandle).toBe(0);
    } finally {
      console.warn = warn;
    }
  });
});

describe('createDisplayListQueries unbuilt pages', () => {
  test('a position on an unbuilt page anchors to its content box', () => {
    const { engine } = fakeEngine();
    const unbuilt: DisplayPage = {
      pageIndex: 1,
      width: 100,
      height: 100,
      primitives: [],
      unbuilt: true,
      positionSpan: [40, 80],
      contentBounds: { x: 10, y: 12, width: 80, height: 70 },
    };
    const queries = createDisplayListQueries({ pages: [page(0), unbuilt] }, engine);
    expect(queries.caretRect(50)).toEqual({ pageIndex: 1, x: 10, y: 12, width: 0, height: 0 });
    expect(queries.anchorRect(50)).toEqual({ pageIndex: 1, x: 10, y: 12, width: 0, height: 0 });
    expect(queries.caretRect(90)).toBeNull();
  });

  test('the first position of an unbuilt page anchors to it, not to the page painted before it', () => {
    const { engine } = fakeEngine();
    engine.rangeRectsByHandle = (_handle, from, to) =>
      from === 39 && to === 40
        ? JSON.stringify([{ pageIndex: 0, x: 20, y: 30, width: 5, height: 10 }])
        : '[]';
    const unbuilt = (pageIndex: number, positionSpan: [number, number]): DisplayPage => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [],
      unbuilt: true,
      positionSpan,
      contentBounds: { x: 10, y: 12 + pageIndex, width: 80, height: 70 },
    });
    const queries = createDisplayListQueries(
      { pages: [page(0), unbuilt(1, [40, 80]), unbuilt(2, [80, 120])] },
      engine
    );
    expect(queries.caretRect(40)).toEqual({ pageIndex: 1, x: 10, y: 13, width: 0, height: 0 });
    expect(queries.caretRect(80)).toEqual({ pageIndex: 2, x: 10, y: 14, width: 0, height: 0 });
  });

  test('a position in a row split across unbuilt pages picks the page by its share of the row', () => {
    const { engine } = fakeEngine();
    const unbuilt = (pageIndex: number, positionSpan: [number, number]): DisplayPage => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [],
      unbuilt: true,
      positionSpan,
    });
    const queries = createDisplayListQueries(
      { pages: [page(0), unbuilt(1, [50, 400]), unbuilt(2, [100, 400]), unbuilt(3, [100, 450])] },
      engine
    );
    expect(queries.caretRect(120)?.pageIndex).toBe(1);
    expect(queries.caretRect(250)?.pageIndex).toBe(2);
    expect(queries.caretRect(390)?.pageIndex).toBe(3);
    expect(queries.caretRect(60)?.pageIndex).toBe(1);
    expect(queries.caretRect(420)?.pageIndex).toBe(3);
  });
});

describe('visual lines', () => {
  const text = (
    x: number,
    baselineY: number,
    docStart: number,
    identity: { paraId?: string; blockKey?: string } = {}
  ) => ({
    kind: 'text' as const,
    text: 'x',
    x,
    baselineY,
    width: 10,
    font: '400 10px Calibri',
    color: '#000000',
    docStart,
    docEnd: docStart + 1,
    ...identity,
  });

  test('group each page on its own, the way the whole list does', () => {
    const lines: DisplayPage = {
      pageIndex: 0,
      width: 600,
      height: 800,
      primitives: [
        text(0, 20, 1, { paraId: 'a' }),
        text(300, 20, 50, { paraId: 'b' }),
        text(20, 21, 2, { paraId: 'a' }),
        text(0, 40, 3, { paraId: 'a' }),
        text(310, 20.5, 51, { paraId: 'b' }),
        text(0, 60, 70),
        text(0, 60, 71),
        text(0, 80, 80, { blockKey: 'k' }),
      ],
    };
    const second: DisplayPage = {
      ...lines,
      pageIndex: 1,
      primitives: [text(0, 20, 90, { paraId: 'c' })],
    };
    const queries = createDisplayListQueries({ pages: [lines, second] }, fakeEngine().engine);

    const first = queries.visualLinesOnPage(0);
    expect(
      first.map(({ paraId, blockId, from, to, x, width }) => [paraId, blockId, from, to, x, width])
    ).toEqual([
      ['a', undefined, 1, 3, 0, 30],
      ['b', undefined, 50, 52, 300, 20],
      ['a', undefined, 3, 4, 0, 10],
      [undefined, undefined, 70, 71, 0, 10],
      [undefined, undefined, 71, 72, 0, 10],
      [undefined, 'k', 80, 81, 0, 10],
    ]);
    expect(queries.visualLines()).toEqual([...first, ...queries.visualLinesOnPage(1)]);
    expect(queries.visualLinesOnPage(2)).toEqual([]);
    expect(queries.visualLineExtent(0)).toEqual({
      top: Math.min(...first.map((line) => line.y)),
      bottom: Math.max(...first.map((line) => line.y + line.height)),
    });
    const blank: DisplayPage = { ...lines, pageIndex: 2, primitives: [] };
    expect(
      createDisplayListQueries({ pages: [blank] }, fakeEngine().engine).visualLineExtent(0)
    ).toBeNull();
  });
});

describe('createDisplayListQueries page loads in the real store', () => {
  test('pages loaded one query at a time answer as a store holding every page', async () => {
    const { preloadLayoutWasm } = await import('../../wasm/layout');
    const { loadRustDisplayListQueryEngine } = await import('./rustDisplayList');
    await preloadLayoutWasm();
    const engine = await loadRustDisplayListQueryEngine();
    const textPage = (pageIndex: number): DisplayPage => ({
      pageIndex,
      width: 816,
      height: 1056,
      primitives: [0, 1, 2].map(
        (line) =>
          ({
            kind: 'text',
            text: `Page ${pageIndex} line ${line}`,
            x: 100,
            baselineY: 100 + line * 20,
            width: 200,
            font: '400 16px Arial',
            color: '#000000',
            docStart: pageIndex * 100 + line * 20 + 1,
            docEnd: pageIndex * 100 + line * 20 + 15,
          }) as DisplayPage['primitives'][number]
      ),
    });
    const list = { pages: Array.from({ length: 12 }, (_, index) => textPage(index)) };
    const calls = { byHandle: 0, json: 0, updates: 0 };
    const counted = {
      ...engine,
      rangeRectsByHandle: (handle: number, from: number, to: number) => {
        calls.byHandle += 1;
        return engine.rangeRectsByHandle!(handle, from, to);
      },
      rangeRectsJson: (json: string, from: number, to: number) => {
        calls.json += 1;
        return engine.rangeRectsJson(json, from, to);
      },
      updateDisplayList: (handle: number, update: string) => {
        calls.updates += 1;
        engine.updateDisplayList!(handle, update);
      },
    };
    const lazy = createDisplayListQueries(list, counted);
    const order = [7, 2, 11, 2, 0, 5];
    const answers = order.map((index) => lazy.rangeRects(index * 100 + 3, index * 100 + 50));
    expect(calls).toEqual({ byHandle: 6, json: 0, updates: 5 });
    const eager = createDisplayListQueries(list, engine);
    eager.rangeRects(1, 1200);
    expect(order.map((index) => eager.rangeRects(index * 100 + 3, index * 100 + 50))).toEqual(
      answers
    );
    expect(answers.every((rects) => rects.length > 0)).toBe(true);
  });
});
