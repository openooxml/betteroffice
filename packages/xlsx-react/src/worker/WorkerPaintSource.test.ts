import type { DisplayList, MergedRange, Viewport, WorkbookFrame } from '@betteroffice/xlsx';
import { describe, expect, test } from 'bun:test';
import {
  WorkerPaintSource,
  type WorkerPaintRequest,
  type WorkerPaintResult,
  type WorkerPeerPaint,
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

function peerPaint(request: WorkerPaintRequest, version = 'peer'): WorkerPeerPaint {
  return {
    displayList: displayList(version, request.viewport), version,
    mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } }],
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
  test('rejects a stale sequence after a peer commit and accepts the worker once it catches up', async () => {
    const { source, state, requests, paints, errors, callbacks, tick, respond, surface } = harness();
    source.schedule();
    tick();
    state.sentSequence = 1;
    const committed = source.commit(state.request, peerPaint(state.request));
    const committedSurface = surface();

    await respond(0, frame(requests[0].request, 0));
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(paints).toHaveLength(1);
    expect(callbacks.size).toBe(1);
    expect(errors).toEqual([]);

    tick();
    await respond(1, frame(requests[1].request, 1));
    expect(paints.map((painted) => painted.source)).toEqual(['peer', 'worker']);
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

  test('publishes a commit synchronously and invalidates an older in-flight frame even at the same sequence', async () => {
    const { source, state, requests, paints, callbacks, tick, respond, surface } = harness();
    source.schedule();
    tick();
    source.schedule();
    state.sentSequence = 2;
    const paint = peerPaint(state.request);
    const committed = source.commit(state.request, paint);
    if (!committed) throw new Error('Missing peer paint');
    const committedSurface = surface();

    expect(paints).toEqual([committed]);
    expect(committed?.source).toBe('peer');
    expect(committed?.displayList).toBe(paint.displayList);
    expect(committed?.geometry).toBe(paint.displayList.grid);
    expect(committed?.mergedRanges).toBe(paint.mergedRanges);
    expect(committed?.version).toBe(paint.version);
    expect(committed?.sequence).toBe(2);
    expect(source.painted).toBe(committed);
    expect(callbacks.size).toBe(0);

    await respond(0, frame(requests[0].request, 2));
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(paints).toHaveLength(1);
    expect(callbacks.size).toBe(1);
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
      const committed = source.commit(state.request, peerPaint(state.request));
      const committedSurface = surface();
      source.schedule();
      tick();
      state.request = { ...state.request, ...change };

      await respond(0, frame(requests[0].request));
      expect(source.painted).toBe(committed);
      expect(surface()).toBe(committedSurface);
      expect(paints).toHaveLength(1);
      expect(errors).toEqual([]);

      if (name === 'generation') {
        expect(callbacks.size).toBe(0);
        source.schedule();
        tick();
        expect(requests).toHaveLength(1);
      } else {
        expect(callbacks.size).toBe(1);
        tick();
        expect(requests[1].request).toEqual(state.request);
        await respond(1, frame(requests[1].request));
        expect(source.painted?.request).toEqual(state.request);
        expect(paints).toHaveLength(2);
      }
    });
  }

  for (const mismatch of ['sheet', 'viewport'] as const) {
    test(`rejects a worker reply with an unexpected ${mismatch} and retries`, async () => {
      const { source, state, requests, paints, errors, callbacks, tick, respond, surface } = harness();
      const committed = source.commit(state.request, peerPaint(state.request));
      const committedSurface = surface();
      source.schedule();
      tick();
      const rejected = frame(requests[0].request);
      if (mismatch === 'sheet') rejected.sheet = 1;
      else rejected.viewport.x = 100;

      await respond(0, rejected);
      expect(source.painted).toBe(committed);
      expect(surface()).toBe(committedSurface);
      expect(paints).toHaveLength(1);
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(1);
      tick();
      await respond(1, frame(requests[1].request));
      expect(paints).toHaveLength(2);
      expect(source.painted?.source).toBe('worker');
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

  test('records source attribution and publishes matching geometry, merges and version for both sources', async () => {
    const { source, state, requests, paints, tick, respond } = harness();
    state.sentSequence = 1;
    const peer = peerPaint(state.request);
    const committed = source.commit(state.request, peer);
    expect(committed).toEqual({
      ...peer, source: 'peer', request: state.request, sequence: 1, geometry: peer.displayList.grid,
    });

    source.schedule();
    tick();
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
    expect(paints.map((painted) => painted.source)).toEqual(['peer', 'worker']);
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

  test('keeps a synchronous peer commit made inside the worker publication callback', async () => {
    let commit = () => {};
    const { source, state, requests, paints, tick, respond, surface } = harness((painted) => {
      if (painted.source === 'worker') commit();
    });
    commit = () => {
      state.sentSequence = 1;
      source.commit(state.request, peerPaint(state.request));
    };
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request));
    expect(paints.map((painted) => painted.source)).toEqual(['worker', 'peer']);
    expect(source.painted).toBe(paints[1]);
    expect(surface()?.pixels).toBe(source.painted?.displayList);
  });

  test('rejects a peer publication captured for an older viewport without cancelling current work', async () => {
    const { source, state, requests, paints, callbacks, tick, respond } = harness();
    const old = request();
    state.request.viewport.x = 100;
    source.schedule();
    expect(source.commit(old, peerPaint(old))).toBeNull();
    expect(paints).toEqual([]);
    expect(callbacks.size).toBe(1);
    tick();
    await respond(0, frame(requests[0].request));
    expect(source.painted?.request.viewport.x).toBe(100);
  });

  for (const lifecycle of ['disposed', 'replaced'] as const) {
    test(`ignores pending worker frames and peer publications after the source is ${lifecycle}`, async () => {
      const { source, state, requests, paints, errors, callbacks, tick, respond } = harness();
      source.schedule();
      tick();
      source.schedule();
      if (lifecycle === 'disposed') source.dispose();
      else state.current = false;

      await respond(0, frame(requests[0].request));
      source.schedule();
      expect(source.commit(state.request, peerPaint(state.request))).toBeNull();
      tick();
      expect(requests).toHaveLength(1);
      expect(paints).toEqual([]);
      expect(errors).toEqual([]);
      expect(callbacks.size).toBe(0);
    });
  }

  test('retries superseded requests without reporting a failure', async () => {
    const { source, requests, errors, callbacks, tick, respond } = harness();
    source.schedule();
    tick();
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

  test('reports a worker failure once, retains the paint and permits peer recovery publication', async () => {
    const { source, state, requests, paints, errors, callbacks, tick, surface } = harness();
    const committed = source.commit(state.request, peerPaint(state.request));
    const committedSurface = surface();
    source.schedule();
    tick();
    const error = new Error('Worker stopped');
    requests[0].reply.reject(error);
    await requests[0].reply.promise.catch(() => {});
    source.fail(new Error('Already failed'));
    source.schedule();
    tick();
    expect(errors).toEqual([error]);
    expect(source.painted).toBe(committed);
    expect(surface()).toBe(committedSurface);
    expect(callbacks.size).toBe(0);
    expect(requests).toHaveLength(1);

    const recovered = source.commit(state.request, peerPaint(state.request, 'recovery'));
    expect(recovered?.source).toBe('peer');
    expect(source.painted?.version).toBe('recovery');
    expect(paints).toHaveLength(2);
    expect(errors).toHaveLength(1);
  });
});
