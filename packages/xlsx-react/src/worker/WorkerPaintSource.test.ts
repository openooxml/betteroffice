import type { DisplayList, MergedRange, Viewport, WorkbookFrame } from '@betteroffice/xlsx';
import { describe, expect, test } from 'bun:test';
import {
  WorkerPaintSource,
  type WorkerPaintRequest,
  type WorkerPaintResult,
} from './WorkerPaintSource';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function request(): WorkerPaintRequest {
  return {
    generation: 1, sheet: 0, navigation: 0, width: 800, height: 600, zoom: 1, dpr: 1,
    viewport: { x: 0, y: 0, width: 800, height: 600 },
  };
}

function displayList(text: string, viewport: Viewport): DisplayList {
  return {
    width: viewport.width, height: viewport.height,
    commands: [{ op: 'text', text, x: 8, y: 18, fontSize: 11, color: '#000000' }],
    grid: { startRow: 0, startCol: 0, rowOffsets: [0, 24, 48], colOffsets: [0, 96, 192] },
  };
}

function frame(request: WorkerPaintRequest, sequence = 0): WorkbookFrame {
  return {
    sheet: request.sheet, viewport: { ...request.viewport }, sequence, epoch: 1,
    displayList: displayList(`worker ${sequence}`, request.viewport), version: `v${sequence}`,
    mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
  };
}

function harness(onPublish?: (painted: WorkerPaintResult) => void) {
  const state = { request: request(), sentSequence: 0, current: true, available: true };
  const callbacks = new Map<number, () => void>();
  const requests: {
    request: WorkerPaintRequest;
    reply: ReturnType<typeof deferred<WorkbookFrame>>;
  }[] = [];
  const paints: WorkerPaintResult[] = [];
  const errors: unknown[] = [];
  let id = 0;
  let surface: {
    pixels: DisplayList;
    geometry: DisplayList['grid'];
    accessibility: DisplayList['commands'];
    mergedRanges: readonly MergedRange[];
    pluginLayout: { geometry: DisplayList['grid']; version: string };
  } | null = null;
  const source = new WorkerPaintSource({
    generation: 1,
    capture: () => state.available ? state.request : null,
    requestFrame: (request) => {
      const reply = deferred<WorkbookFrame>();
      requests.push({ request, reply });
      return reply.promise;
    },
    sentSequence: () => state.sentSequence,
    publish: (painted) => {
      paints.push(painted);
      surface = {
        pixels: painted.displayList, geometry: painted.geometry,
        accessibility: painted.displayList.commands, mergedRanges: painted.mergedRanges,
        pluginLayout: { geometry: painted.geometry, version: painted.version },
      };
      onPublish?.(painted);
    },
    onError: (error) => { errors.push(error); },
    isCurrent: () => state.current,
    scheduler: {
      request: (callback) => { callbacks.set(++id, callback); return id; },
      cancel: (id) => { callbacks.delete(id); },
    },
  });
  const tick = () => {
    const scheduled = [...callbacks.values()];
    callbacks.clear();
    for (const callback of scheduled) callback();
  };
  const respond = async (index: number, frame: WorkbookFrame) => {
    const reply = requests[index].reply;
    reply.resolve(frame);
    await reply.promise;
  };
  return { source, state, callbacks, requests, paints, errors, tick, respond, surface: () => surface };
}

describe('WorkerPaintSource', () => {
  for (const result of ['accepted', 'failed'] as const) {
    test(`coalesces scheduled paints until RAF resumes (${result})`, async () => {
      const { source, requests, paints, errors, callbacks, tick, respond } = harness();
      source.schedule();
      expect(requests).toHaveLength(0);
      expect(callbacks.size).toBe(1);
      tick();
      expect(requests).toHaveLength(1);
      if (result === 'failed') {
        const failure = new Error('Worker paint failed');
        requests[0].reply.reject(failure);
        await requests[0].reply.promise.catch(() => {});
        expect(errors).toEqual([failure]);
        expect(paints).toEqual([]);
      } else {
        await respond(0, frame(requests[0].request));
        expect(errors).toEqual([]);
        expect(paints).toHaveLength(1);
      }
    });
  }

  test('coalesces edit paints into one latest sequence request', async () => {
    const { source, state, requests, callbacks, tick, respond } = harness();
    state.sentSequence = 1;
    source.schedule();
    state.sentSequence = 2;
    source.schedule();
    expect(callbacks.size).toBe(1);
    tick();
    expect(requests).toHaveLength(1);
    await respond(0, frame(requests[0].request, 2));
    expect(source.painted?.sequence).toBe(2);
  });

  for (const result of ['accepted', 'failed'] as const) {
    test(`drops stale viewport failures and processes the latest frame (${result})`, async () => {
      const { source, state, requests, errors, callbacks, tick, respond } = harness();
      source.schedule();
      tick();
      state.request.viewport.x = 100;
      source.schedule();
      requests[0].reply.reject(new Error('Old viewport failed'));
      await requests[0].reply.promise.catch(() => {});
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(1);
      tick();
      if (result === 'failed') {
        const failure = new Error('Current viewport failed');
        requests[1].reply.reject(failure);
        await requests[1].reply.promise.catch(() => {});
        expect(errors).toEqual([failure]);
      } else {
        await respond(1, frame(requests[1].request));
        expect(source.painted?.request.viewport.x).toBe(100);
        expect(errors).toEqual([]);
      }
    });
  }

  test('cancels every coalesced paint on retirement without issuing a frame', () => {
    const { source, requests, callbacks, errors, tick } = harness();
    source.schedule();
    source.schedule();
    source.dispose();
    expect(callbacks.size).toBe(0);
    tick();
    expect(requests).toEqual([]);
    expect(errors).toEqual([]);
  });

  test('cancels a repaint without a local change', () => {
    const { source, requests, callbacks, tick } = harness();
    source.schedule();
    expect(callbacks.size).toBe(1);
    source.dispose();
    tick();
    expect(requests).toEqual([]);
  });

  test('drops an in-flight frame when a dependent edit advances its sequence', async () => {
    const { source, state, requests, paints, errors, tick, respond } = harness();
    source.schedule();
    tick();
    state.sentSequence = 1;
    source.schedule();
    await respond(0, frame(requests[0].request));
    expect(paints).toHaveLength(0);
    expect(errors).toEqual([]);
    tick();
    await respond(1, frame(requests[1].request, 1));
    expect(paints).toHaveLength(1);
  });

  test('fails the session on a real canvas publication error', async () => {
    const failure = new Error('Canvas publication failed');
    const { source, requests, errors, callbacks, tick, respond } = harness(() => { throw failure; });
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    expect(errors).toEqual([failure]);
    expect(source.painted).toBeNull();
    expect(callbacks.size).toBe(0);
  });

  test('discards an obsolete viewport error without failing the session', async () => {
    const { source, state, requests, errors, tick } = harness();
    source.schedule();
    tick();
    state.request.viewport.x = 100;
    source.schedule();
    requests[0].reply.reject(new Error('Old viewport failed'));
    await requests[0].reply.promise.catch(() => {});
    expect(errors).toEqual([]);
    expect(source.painted).toBeNull();
    expect(requests).toHaveLength(1);
  });

  test('disposes a source before its blocked worker frame settles', async () => {
    const { source, requests, errors, tick, respond } = harness();
    source.schedule();
    tick();
    source.dispose();
    expect(errors).toEqual([]);
    await respond(0, frame(requests[0].request));
    expect(source.painted).toBeNull();
  });

  test('keeps discarding obsolete viewport errors as scrolling advances', async () => {
    const { source, state, requests, errors, tick, respond } = harness();
    source.schedule();
    for (let index = 0; index < 6; index++) {
      tick();
      state.request.viewport.x += 100;
      source.schedule();
      requests[index].reply.reject(new Error('Old viewport failed'));
      await requests[index].reply.promise.catch(() => {});
    }
    expect(errors).toEqual([]);
    tick();
    await respond(6, frame(requests[6].request));
    expect(source.painted?.request.viewport.x).toBe(600);
  });

  for (const reply of ['stale', 'superseded'] as const) {
    test(`retries ${reply} replies without a budget while the edit sequence advances`, async () => {
      const { source, state, requests, errors, callbacks, tick, respond } = harness();
      const superseded = new Error('Frame superseded');
      superseded.name = 'SessionSuperseded';
      source.schedule();
      for (let attempt = 0; attempt < 6; attempt++) {
        tick();
        state.sentSequence += 1;
        if (reply === 'stale') await respond(attempt, frame(requests[attempt].request, attempt));
        else {
          requests[attempt].reply.reject(superseded);
          await requests[attempt].reply.promise.catch(() => {});
        }
        expect(errors).toEqual([]);
        expect(callbacks.size).toBe(1);
      }
      tick();
      await respond(6, frame(requests[6].request, 6));
      expect(source.painted?.sequence).toBe(6);
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(0);
    });

    test(`drops ${reply} replies without retrying an unchanged request`, async () => {
      const { source, state, requests, errors, callbacks, tick, respond } = harness();
      state.sentSequence = 1;
      source.schedule();
      tick();
      if (reply === 'stale') await respond(0, frame(requests[0].request, 0));
      else {
        const error = new Error('Frame superseded');
        error.name = 'SessionSuperseded';
        requests[0].reply.reject(error);
        await requests[0].reply.promise.catch(() => {});
      }
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(0);
      tick();
      expect(requests).toHaveLength(1);
    });
  }

  test('retains the adopted worker frame until a matching edit sequence arrives', async () => {
    const { source, state, requests, paints, errors, callbacks, tick, respond, surface } = harness();
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    const committed = source.painted;
    source.schedule();
    tick();
    state.sentSequence = 1;
    const committedSurface = surface();

    await respond(1, frame(requests[1].request, 0));
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(paints).toHaveLength(1);
    expect(callbacks.size).toBe(1);
    expect(errors).toEqual([]);

    tick();
    await respond(2, frame(requests[2].request, 1));
    expect(paints.map((painted) => painted.source)).toEqual(['worker', 'worker']);
    expect(source.painted?.sequence).toBe(1);
    expect(callbacks.size).toBe(0);
  });

  test('reads the sent sequence at acceptance time without another paint reason', async () => {
    const { source, state, requests, paints, callbacks, tick, respond } = harness();
    source.schedule();
    tick();
    state.sentSequence = 2;

    await respond(0, frame(requests[0].request, 1));
    expect(paints).toEqual([]);
    expect(source.painted).toBeNull();
    expect(callbacks.size).toBe(1);

    tick();
    await respond(1, frame(requests[1].request, 2));
    expect(source.painted?.sequence).toBe(2);
    expect(source.painted?.source).toBe('worker');
  });

  for (const sequence of [2, 3]) {
    test(`accepts sequence ${sequence} at or above the current sent sequence`, async () => {
      const { source, state, requests, paints, tick, respond } = harness();
      state.sentSequence = 2;
      source.schedule();
      tick();
      await respond(0, frame(requests[0].request, sequence));
      expect(paints).toHaveLength(1);
      expect(source.painted?.sequence).toBe(sequence);
    });
  }

  test('waits for worker adoption and rejects an older in-flight frame even at the same sequence', async () => {
    const { source, state, requests, paints, callbacks, tick, respond, surface } = harness();
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    const committed = source.painted;
    const committedSurface = surface();
    expect(committed).not.toBeNull();
    if (!committed) throw new Error('Missing adopted worker frame');
    source.schedule();
    tick();
    source.schedule();
    state.sentSequence = 2;

    expect(paints).toEqual([committed]);
    expect(committed?.source).toBe('worker');
    expect(committed?.sequence).toBe(0);
    expect(source.painted).toBe(committed);
    expect(callbacks.size).toBe(1);

    await respond(1, frame(requests[1].request, 2));
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(paints).toHaveLength(1);
    expect(callbacks.size).toBe(1);
    tick();
    const worker = frame(requests[2].request, 2);
    if (!worker.mergedRanges) throw new Error('Missing worker merged ranges');
    await respond(2, worker);
    expect(paints.map((painted) => painted.source)).toEqual(['worker', 'worker']);
    expect(source.painted?.displayList).toBe(worker.displayList);
    expect(source.painted?.geometry).toBe(worker.displayList.grid);
    expect(source.painted?.mergedRanges).toBe(worker.mergedRanges);
    expect(source.painted?.version).toBe(worker.version);
    expect(source.painted?.sequence).toBe(2);
    expect(surface()?.pixels).toBe(worker.displayList);
    expect(callbacks.size).toBe(0);
  });

  test('coalesces many viewport reasons into one in-flight request and one follow-up', async () => {
    const { source, state, requests, paints, callbacks, tick, respond } = harness();
    for (let reason = 0; reason < 10; reason += 1) source.schedule();
    expect(callbacks.size).toBe(1);
    expect(requests).toHaveLength(0);
    tick();
    expect(requests).toHaveLength(1);

    for (const x of [100, 200, 300, 400]) {
      state.request = { ...state.request, viewport: { ...state.request.viewport, x } };
      for (let reason = 0; reason < 10; reason += 1) source.schedule();
      expect(callbacks.size).toBe(1);
      tick();
      expect(requests).toHaveLength(1);
    }

    await respond(0, frame(requests[0].request));
    expect(paints).toEqual([]);
    expect(requests).toHaveLength(2);
    expect(requests[1].request.viewport.x).toBe(400);
    await respond(1, frame(requests[1].request));
    expect(paints).toHaveLength(1);
    expect(source.painted?.request.viewport.x).toBe(400);
    expect(requests).toHaveLength(2);
    expect(callbacks.size).toBe(0);
  });

  const mismatches: [string, Partial<WorkerPaintRequest>][] = [
    ['generation', { generation: 2 }],
    ['sheet', { sheet: 1 }],
    ['navigation', { navigation: 1 }],
    ['zoom', { zoom: 2 }],
    ['DPR', { dpr: 2 }],
    ['surface width', { width: 801 }],
    ['surface height', { height: 601 }],
    ['viewport x', { viewport: { ...request().viewport, x: 100 } }],
    ['viewport y', { viewport: { ...request().viewport, y: 100 } }],
    ['viewport width', { viewport: { ...request().viewport, width: 400 } }],
    ['viewport height', { viewport: { ...request().viewport, height: 300 } }],
  ];

  for (const [name, change] of mismatches) {
    test(`rejects a frame when ${name} changes before the next scheduled request`, async () => {
      const { source, state, requests, paints, errors, callbacks, tick, respond, surface } = harness();
      source.schedule();
      tick();
      await respond(0, frame(requests[0].request));
      const committed = source.painted;
      const committedSurface = surface();
      source.schedule();
      tick();
      state.request = { ...state.request, ...change };

      await respond(1, frame(requests[1].request));
      expect(source.painted).toBe(committed);
      expect(surface()).toBe(committedSurface);
      expect(paints).toHaveLength(1);
      expect(errors).toEqual([]);

      if (name === 'generation') {
        expect(callbacks.size).toBe(0);
        source.schedule();
        tick();
        expect(requests).toHaveLength(2);
      } else {
        expect(callbacks.size).toBe(1);
        tick();
        expect(requests[2].request).toEqual(state.request);
        await respond(2, frame(requests[2].request));
        expect(source.painted?.request).toEqual(state.request);
        expect(paints).toHaveLength(2);
      }
      expect(paints.every((painted) => painted.source === 'worker')).toBe(true);
    });
  }

  for (const mismatch of ['sheet', 'viewport'] as const) {
    test(`drops a worker reply with an unexpected ${mismatch} until the viewport moves`, async () => {
      const { source, state, requests, paints, errors, callbacks, tick, respond, surface } = harness();
      source.schedule();
      tick();
      await respond(0, frame(requests[0].request));
      const committed = source.painted;
      const committedSurface = surface();
      source.schedule();
      tick();
      const rejected = frame(requests[1].request);
      if (mismatch === 'sheet') rejected.sheet = 1;
      else rejected.viewport.x = 100;

      await respond(1, rejected);
      expect(source.painted).toBe(committed);
      expect(surface()).toBe(committedSurface);
      expect(paints).toHaveLength(1);
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(0);
      state.request.viewport.x = 200;
      source.schedule();
      tick();
      await respond(2, frame(requests[2].request));
      expect(paints).toHaveLength(2);
      expect(source.painted?.source).toBe('worker');
      expect(paints.every((painted) => painted.source === 'worker')).toBe(true);
    });
  }

  test('snapshots mutable capture values so they cannot make an old frame match', async () => {
    const { source, state, requests, paints, tick, respond } = harness();
    source.schedule();
    tick();
    state.request.viewport.x = 100;
    state.request.zoom = 2;
    expect(requests[0].request.viewport.x).toBe(0);
    expect(requests[0].request.zoom).toBe(1);

    await respond(0, frame(requests[0].request));
    expect(paints).toEqual([]);
    tick();
    await respond(1, frame(requests[1].request));
    expect(source.painted?.request).toEqual(state.request);
    state.request.viewport.x = 200;
    expect(source.painted?.request.viewport.x).toBe(100);
  });

  test('publishes only worker frames with matching geometry, merges and version', async () => {
    const { source, state, requests, paints, tick, respond } = harness();
    state.sentSequence = 1;
    source.schedule();
    tick();
    expect(source.painted).toBeNull();
    expect(paints).toEqual([]);
    const worker = frame(requests[0].request, 1);
    await respond(0, worker);
    if (!worker.mergedRanges) throw new Error('Missing worker merged ranges');
    expect(source.painted).toEqual({
      source: 'worker', request: state.request, sequence: 1, displayList: worker.displayList,
      geometry: worker.displayList.grid, mergedRanges: worker.mergedRanges, version: worker.version,
    } satisfies WorkerPaintResult);
    expect(source.painted?.displayList).toBe(worker.displayList);
    expect(source.painted?.geometry).toBe(worker.displayList.grid);
    expect(source.painted?.mergedRanges).toBe(worker.mergedRanges);
    expect(paints.map((painted) => painted.source)).toEqual(['worker']);
  });

  test('normalizes missing worker merges without inventing grid geometry', async () => {
    const { source, requests, tick, respond } = harness();
    source.schedule();
    tick();
    const worker = frame(requests[0].request);
    delete worker.mergedRanges;
    delete worker.displayList.grid;
    await respond(0, worker);
    expect(source.painted?.mergedRanges).toEqual([]);
    expect(source.painted?.geometry).toBeUndefined();
  });

  test('retains worker pixels when a new edit is scheduled inside worker publication', async () => {
    let schedule = () => {};
    const { source, state, requests, paints, tick, respond, surface } = harness((painted) => {
      if (painted.sequence === 0) schedule();
    });
    schedule = () => {
      state.sentSequence = 1;
      source.schedule();
    };
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    expect(paints.map((painted) => painted.source)).toEqual(['worker']);
    expect(source.painted).toBe(paints[0]);
    expect(surface()?.pixels).toBe(source.painted?.displayList);
    tick();
    await respond(1, frame(requests[1].request, 1));
    expect(paints.map((painted) => painted.source)).toEqual(['worker', 'worker']);
    expect(source.painted).toBe(paints[1]);
    expect(surface()?.pixels).toBe(source.painted?.displayList);
  });

  test('rejects an older worker viewport without cancelling current work', async () => {
    const { source, state, requests, paints, callbacks, tick, respond } = harness();
    source.schedule();
    tick();
    state.request.viewport.x = 100;
    source.schedule();
    await respond(0, frame(requests[0].request));
    expect(source.painted).toBeNull();
    expect(paints).toEqual([]);
    expect(callbacks.size).toBe(1);
    tick();
    await respond(1, frame(requests[1].request));
    expect(source.painted?.request.viewport.x).toBe(100);
    expect(paints.map((painted) => painted.source)).toEqual(['worker']);
  });

  for (const lifecycle of ['disposed', 'replaced'] as const) {
    test(`ignores pending worker frames after the source is ${lifecycle}`, async () => {
      const { source, state, requests, paints, errors, callbacks, tick, respond } = harness();
      source.schedule();
      tick();
      source.schedule();
      if (lifecycle === 'disposed') source.dispose();
      else state.current = false;

      await respond(0, frame(requests[0].request));
      source.schedule();
      expect(source.painted).toBeNull();
      tick();
      expect(requests).toHaveLength(1);
      expect(paints).toEqual([]);
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(0);
    });
  }

  test('retries superseded requests after the viewport moves without reporting a failure', async () => {
    const { source, state, requests, errors, callbacks, tick, respond } = harness();
    source.schedule();
    tick();
    state.request.viewport.x = 100;
    const error = new Error('Frame superseded');
    error.name = 'SessionSuperseded';
    requests[0].reply.reject(error);
    await requests[0].reply.promise.catch(() => {});
    expect(errors).toEqual([]);
    expect(callbacks.size).toBe(1);
    tick();
    await respond(1, frame(requests[1].request));
    expect(source.painted?.source).toBe('worker');
  });

  test('reports a worker failure once and retains the adopted worker paint without peer recovery publication', async () => {
    const { source, requests, paints, errors, callbacks, tick, respond, surface } = harness();
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    const committed = source.painted;
    const committedSurface = surface();
    source.schedule();
    tick();
    const error = new Error('Worker stopped');
    requests[1].reply.reject(error);
    await requests[1].reply.promise.catch(() => {});
    source.fail(new Error('Already failed'));
    source.schedule();
    tick();
    expect(errors).toEqual([error]);
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(callbacks.size).toBe(0);
    expect(requests).toHaveLength(2);
    expect(source.painted?.version).toBe('v0');
    expect(paints.map((painted) => painted.source)).toEqual(['worker']);
    expect(paints).toHaveLength(1);
    expect(errors).toHaveLength(1);
  });
});
