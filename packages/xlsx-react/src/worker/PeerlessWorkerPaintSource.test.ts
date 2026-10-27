import type { DisplayList, MergedRange, Viewport, WorkbookFrame } from '@betteroffice/xlsx';
import type { WorkbookCommandFrame } from '@betteroffice/xlsx/worker';
import { describe, expect, test } from 'bun:test';
import {
  PeerlessWorkerPaintSource as WorkerPaintSource,
  type WorkerPaintRequest,
  type WorkerPaintResult,
} from './PeerlessWorkerPaintSource';

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

function replyFrame(request: WorkerPaintRequest, coversThroughSeq = 1): WorkbookCommandFrame {
  return { ...frame(request, coversThroughSeq), session: 'book', incarnation: 'worker',
    coversThroughSeq, requestGeneration: request.requestGeneration ?? request.generation, zoom: request.zoom, dpr: request.dpr };
}

function harness(onPublish?: (painted: WorkerPaintResult) => void, onInstalled?: (painted: WorkerPaintResult) => void,
  onCovered?: (coversThroughSeq: number) => void) {
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
    commandIdentity: { session: 'book', incarnation: 'worker' },
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
    onInstalled,
    onCovered,
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
  test('coalesces a burst of pipe frames and presents only the newest on the next animation frame', () => {
    const installed: WorkerPaintResult[] = [];
    const covered: number[] = [];
    const { source, state, paints, requests, callbacks, tick, surface } = harness(undefined,
      (painted) => installed.push(painted), (seq) => covered.push(seq));
    const transactions: number[] = [];
    source.registerPaintTransaction((painted, publish) => { transactions.push(painted.coversThroughSeq!); publish(); });
    for (let seq = 1; seq <= 201; seq++) {
      state.sentSequence = seq;
      const request = source.captureCommandFrame()!;
      expect(source.queuePipeFrame(replyFrame(request, seq), request)).toBe(true);
    }
    expect(paints).toEqual([]);
    expect(installed).toEqual([]);
    expect(covered).toEqual([]);
    expect(transactions).toEqual([]);
    expect(source.coversThroughSeq).toBe(0);
    expect(surface()).toBeNull();
    expect(callbacks.size).toBe(1);
    tick();
    expect(paints).toHaveLength(1);
    expect(installed).toEqual(paints);
    expect(transactions).toEqual([201]);
    expect(source.painted).toMatchObject({ version: 'v201', coversThroughSeq: 201 });
    expect(source.coversThroughSeq).toBe(201);
    expect(surface()!.pixels).toBe(paints[0].displayList);
    expect(requests).toEqual([]);
    expect(callbacks.size).toBe(0);
  });

  test('keeps the newest pending pipe frame when an older ack arrives later', () => {
    const { source, paints, tick } = harness();
    const request = source.captureCommandFrame()!;
    expect(source.queuePipeFrame(replyFrame(request, 2), request)).toBe(true);
    expect(source.queuePipeFrame(replyFrame(request, 1), request)).toBe(true);
    tick();
    expect(paints.map((paint) => paint.coversThroughSeq)).toEqual([2]);
  });

  test('keeps the newest pending pipe frame when equal coverage arrives with an older sequence', () => {
    const { source, paints, callbacks, tick } = harness();
    const request = source.captureCommandFrame()!;
    const newest = { ...replyFrame(request, 2), sequence: 102, version: 'newest' };
    const older = { ...replyFrame(request, 2), sequence: 101, version: 'older' };
    expect(source.queuePipeFrame(newest, request)).toBe(true);
    expect(source.queuePipeFrame(older, request)).toBe(true);
    expect(paints).toEqual([]);
    expect(callbacks.size).toBe(1);
    tick();
    expect(paints).toHaveLength(1);
    expect(source.painted).toMatchObject({ sequence: 102, version: 'newest', coversThroughSeq: 2 });
    expect(callbacks.size).toBe(0);
  });

  test('unchanged-version coverage cannot bypass a pending pipe paint', () => {
    const covered: number[] = [];
    const { source, paints, tick } = harness(undefined, undefined, (seq) => covered.push(seq));
    const request = source.captureCommandFrame()!;
    expect(source.install(replyFrame(request, 1), request)).toBe(true);
    expect(source.queuePipeFrame(replyFrame(request, 2), request)).toBe(true);
    expect(source.advanceCoverage('v1', 3)).toBe(false);
    expect(source.coversThroughSeq).toBe(1);
    expect(covered).toEqual([]);
    tick();
    expect(paints.map((paint) => paint.coversThroughSeq)).toEqual([1, 2]);
  });

  test('a failed queued pipe paint does not advance presentation', () => {
    const installed: WorkerPaintResult[] = [];
    const { source, paints, errors, tick } = harness(undefined, (painted) => installed.push(painted));
    source.registerPaintTransaction(() => { throw new Error('Paint failed'); });
    const request = source.captureCommandFrame()!;
    expect(source.queuePipeFrame(replyFrame(request), request)).toBe(true);
    tick();
    expect(paints).toEqual([]);
    expect(installed).toEqual([]);
    expect(source.coversThroughSeq).toBe(0);
    expect(errors).toHaveLength(1);
  });

  test('installs local command frames immediately and cancels covered pending pipe frames', () => {
    const { source, paints, callbacks, tick } = harness();
    const request = source.captureCommandFrame()!;
    expect(source.queuePipeFrame(replyFrame(request, 1), request)).toBe(true);
    expect(source.install(replyFrame(request, 2), request)).toBe(true);
    expect(paints.map((paint) => paint.coversThroughSeq)).toEqual([2]);
    expect(source.coversThroughSeq).toBe(2);
    expect(callbacks.size).toBe(0);
    tick();
    expect(paints).toHaveLength(1);
  });

  test('rejects a queued pipe frame after viewport navigation and requests the current viewport', () => {
    const { source, state, paints, requests, tick } = harness();
    const request = source.captureCommandFrame()!;
    expect(source.queuePipeFrame(replyFrame(request), request)).toBe(true);
    state.request = { ...state.request, navigation: 1, viewport: { ...state.request.viewport, y: 24 } };
    tick();
    expect(paints).toEqual([]);
    expect(source.coversThroughSeq).toBe(0);
    tick();
    expect(requests).toHaveLength(1);
    expect(requests[0].request.viewport.y).toBe(24);
  });

  test('viewport invalidation replaces a queued pipe install with a frame request on the next animation', () => {
    const { source, state, paints, requests, callbacks, tick } = harness();
    const request = source.captureCommandFrame()!;
    expect(source.queuePipeFrame(replyFrame(request), request)).toBe(true);
    state.request = { ...state.request, viewport: { ...state.request.viewport, y: 24 } };
    source.schedule();
    expect(callbacks.size).toBe(1);
    tick();
    expect(paints).toEqual([]);
    expect(source.coversThroughSeq).toBe(0);
    expect(requests).toHaveLength(1);
    expect(requests[0].request.viewport.y).toBe(24);
  });

  for (const cancel of ['dispose', 'fail'] as const) {
    test(`${cancel} cancels a pending pipe frame without presenting it`, () => {
      const { source, paints, callbacks, tick } = harness();
      const request = source.captureCommandFrame()!;
      expect(source.queuePipeFrame(replyFrame(request), request)).toBe(true);
      if (cancel === 'dispose') source.dispose();
      else source.fail(new Error('Session failed'));
      expect(callbacks.size).toBe(0);
      tick();
      expect(paints).toEqual([]);
      expect(source.coversThroughSeq).toBe(0);
    });
  }

  test('advances unchanged-version coverage without installing or publishing a frame', () => {
    const covered: number[] = [];
    const installed: WorkerPaintResult[] = [];
    const { source, state, paints, callbacks, requests } = harness(undefined,
      (painted) => installed.push(painted), (seq) => covered.push(seq));
    expect(source.advanceCoverage('v1', 1)).toBe(false);
    const request = source.captureCommandFrame()!;
    expect(source.install(replyFrame(request, 1), request)).toBe(true);
    const painted = source.painted;
    state.request = { ...state.request, navigation: 1 };
    expect(source.advanceCoverage('v1', 2)).toBe(true);
    expect(source.advanceCoverage('v1', 2)).toBe(true);
    expect(source.coversThroughSeq).toBe(2);
    expect(source.painted).toBe(painted);
    expect(paints).toHaveLength(1);
    expect(installed).toHaveLength(1);
    expect(covered).toEqual([2]);
    expect(callbacks.size).toBe(0);
    expect(requests).toEqual([]);
    expect(source.advanceCoverage('v2', 3)).toBe(false);
    expect(source.advanceCoverage('v1', 1)).toBe(false);
    state.request = { ...state.request, viewport: { ...state.request.viewport, x: 100 } };
    expect(source.advanceCoverage('v1', 3)).toBe(false);
    expect(source.coversThroughSeq).toBe(2);
  });

  test('coverage cannot bypass a pending visual invalidation', () => {
    const { source, state, paints } = harness();
    const request = source.captureCommandFrame()!;
    expect(source.install(replyFrame(request, 1), request)).toBe(true);
    source.schedule();
    expect(source.advanceCoverage('v1', 2)).toBe(false);
    expect(source.coversThroughSeq).toBe(1);
    expect(paints).toHaveLength(1);
    state.current = false;
    expect(source.advanceCoverage('v1', 2)).toBe(false);
  });

  test('presentation follows the entire successful paint transaction', () => {
    const events: string[] = [];
    const { source, state } = harness(() => events.push('canvas'), () => events.push('presented'));
    source.registerPaintTransaction((_painted, publish) => { publish(); events.push('echo'); });
    expect(source.install(replyFrame(state.request), state.request)).toBe(true);
    expect(events).toEqual(['canvas', 'echo', 'presented']);
    const failed = harness(() => events.push('failed canvas'), () => events.push('failed presentation'));
    failed.source.registerPaintTransaction((_painted, publish) => { publish(); throw new Error('Echo paint failed'); });
    expect(failed.source.install(replyFrame(failed.state.request), failed.state.request)).toBe(false);
    expect(failed.source.coversThroughSeq).toBe(0);
    expect(events).not.toContain('failed presentation');
  });

  test('earlier command frame cannot satisfy a later request', () => {
    const { source, state, paints } = harness();
    state.sentSequence = 1;
    const earlier = source.captureCommandFrame()!;
    state.sentSequence = 2;
    const later = source.captureCommandFrame()!;
    expect(later.coversThroughSeq).toBe(2);
    expect(source.install({ ...replyFrame(later, 1), sequence: 200 }, later)).toBe(false);
    expect(source.painted).toBeNull();
    expect(source.install({ ...replyFrame(earlier, 1), sequence: 201 }, earlier)).toBe(true);
    expect(source.coversThroughSeq).toBe(1);
    expect(source.install({ ...replyFrame(later, 2), sequence: 202 }, later)).toBe(true);
    expect(paints.map((paint) => paint.coversThroughSeq)).toEqual([1, 2]);
  });

  test('covering install retires echo in the same transaction', () => {
    const events: string[] = [];
    let echo = [1, 3];
    const { source, state, requests, callbacks } = harness((painted) => {
      events.push(`publish:${painted.coversThroughSeq}:${echo.join(',')}`);
    });
    source.registerPaintTransaction((painted, publish) => {
      events.push('begin');
      echo = echo.filter((seq) => seq > (painted.coversThroughSeq ?? 0));
      publish();
      events.push(`redraw:${echo.join(',')}`);
    });
    expect(source.install(replyFrame(state.request, 2), state.request)).toBe(true);
    expect(events).toEqual(['begin', 'publish:2:3', 'redraw:3']);
    expect(source.coversThroughSeq).toBe(2);
    expect(requests).toEqual([]);
    expect(callbacks.size).toBe(0);
  });

  test('frame sequence and command coverage are tracked separately', async () => {
    const { source, requests, tick, respond } = harness();
    source.schedule();
    tick();
    await respond(0, frame(requests[0].request, 100));
    const captured = source.captureCommandFrame()!;
    expect(source.install({ ...replyFrame(captured, 2), sequence: 101 }, captured)).toBe(true);
    expect(source.painted?.sequence).toBe(101);
    expect(source.painted?.coversThroughSeq).toBe(2);
    expect(source.install({ ...replyFrame(captured, 1), sequence: 500 }, captured)).toBe(false);
    expect(source.install({ ...replyFrame(captured, 500), sequence: 102 }, captured)).toBe(true);
    source.schedule();
    tick();
    await respond(1, frame(requests[1].request, 103));
    expect(source.painted?.sequence).toBe(103);
    expect(source.painted?.coversThroughSeq).toBeUndefined();
    expect(source.coversThroughSeq).toBe(500);
    source.schedule();
    tick();
    await respond(2, frame(requests[2].request, 102));
    expect(source.painted?.sequence).toBe(103);
    const latest = source.captureCommandFrame()!;
    expect(source.install({ ...replyFrame(latest, 499), sequence: 600 }, latest)).toBe(false);
  });

  test('covering install invalidates an older viewport failure', async () => {
    const { source, state, requests, paints, errors, tick } = harness();
    source.schedule();
    tick();
    const captured = source.captureCommandFrame()!;
    expect(source.install({ ...replyFrame(captured, 2), sequence: 1 }, captured)).toBe(true);
    requests[0].reply.reject(new Error('Obsolete viewport failed'));
    await requests[0].reply.promise.catch(() => {});
    expect(errors).toEqual([]);
    expect(source.coversThroughSeq).toBe(2);
    expect(paints).toHaveLength(1);
    state.sentSequence = 1;
    expect(source.captureCommandFrame()).not.toBeNull();
  });

  test('unregisters the paint transaction', () => {
    const { source, state } = harness();
    let calls = 0;
    const unregister = source.registerPaintTransaction((_painted, publish) => { calls++; publish(); });
    expect(source.install(replyFrame(state.request, 1), state.request)).toBe(true);
    unregister();
    expect(source.install(replyFrame(state.request, 2), state.request)).toBe(true);
    expect(calls).toBe(1);
  });

  test('initial frame installs without scheduled frame RPC', () => {
    const { source, state, paints, requests, callbacks } = harness();
    const initial = replyFrame({ ...state.request, requestGeneration: 7 }, 0);
    expect(source.installInitial(initial)).toBe(true);
    expect(source.painted?.displayList).toBe(initial.displayList);
    expect(paints).toHaveLength(1);
    expect(requests).toEqual([]);
    expect(callbacks.size).toBe(0);
  });

  test('resized initial frame is rejected', () => {
    for (const change of [{ width: 700 }, { zoom: 2 }, { dpr: 2 }]) {
      const { source, state, paints, requests, callbacks } = harness();
      const initial = replyFrame({ ...state.request, requestGeneration: 7 }, 0);
      state.request = { ...state.request, ...change,
        viewport: { ...state.request.viewport, width: (change.width ?? 800) / (change.zoom ?? 1) } };
      expect(source.installInitial(initial)).toBe(false);
      expect(paints).toEqual([]);
      expect(requests).toEqual([]);
      expect(callbacks.size).toBe(0);
    }
  });

  test('initial frame rejects an obsolete measured request generation', () => {
    const { source, state, paints } = harness();
    const initial = replyFrame({ ...state.request, requestGeneration: 7 }, 0);
    state.request.requestGeneration = 8;
    expect(source.installInitial(initial)).toBe(false);
    expect(paints).toEqual([]);
  });

  test('installs reply without RAF or frame RPC', () => {
    const { source, state, paints, requests, callbacks } = harness();
    state.sentSequence = 2;
    const captured = source.captureCommandFrame();
    expect(captured).toEqual({ ...state.request, requestGeneration: 0, coversThroughSeq: 2 });
    const reply = replyFrame(captured!, 2);
    expect(source.install(reply, captured!)).toBe(true);
    expect(paints).toHaveLength(1);
    expect(source.painted?.displayList).toBe(reply.displayList);
    expect(source.painted?.sequence).toBe(2);
    expect(source.painted?.coversThroughSeq).toBe(2);
    expect(requests).toEqual([]);
    expect(callbacks.size).toBe(0);
  });

  test('rejects stale coverage and changed viewport', () => {
    const { source, state, paints, requests } = harness();
    const captured = source.captureCommandFrame()!;
    expect(source.install(replyFrame(captured, 2), captured)).toBe(true);
    const installed = source.painted;
    expect(source.install(replyFrame(captured, 1), captured)).toBe(false);
    state.request = { ...state.request, viewport: { ...state.request.viewport, x: 100 } };
    expect(source.install(replyFrame(captured, 3), captured)).toBe(false);
    expect(source.painted).toBe(installed);
    expect(paints).toHaveLength(1);
    expect(requests).toEqual([]);
  });

  test('rejects an outdated request generation after the viewport returns', () => {
    const { source, state, paints } = harness();
    const captured = source.captureCommandFrame()!;
    const reply = replyFrame(captured);
    state.request.viewport.x = 100;
    source.schedule();
    state.request.viewport.x = 0;
    source.schedule();
    expect(source.install(reply, captured)).toBe(false);
    const current = source.captureCommandFrame()!;
    expect(source.install(replyFrame(current), current)).toBe(true);
    expect(paints).toHaveLength(1);
  });

  for (const [name, change] of [
    ['session', { session: 'other' }], ['incarnation', { incarnation: 'other' }],
    ['request generation', { requestGeneration: 0 }], ['zoom', { zoom: 2 }],
    ['DPR', { dpr: 2 }], ['sheet', { sheet: 1 }],
  ] as const) {
    test(`rejects reply with wrong ${name}`, () => {
      const { source, state, paints } = harness();
      expect(source.install({ ...replyFrame(state.request), ...change }, state.request)).toBe(false);
      expect(paints).toEqual([]);
    });
  }

  test('bare ack publishes no paint', () => {
    const { source, state, paints, requests, callbacks } = harness();
    expect(source.install(undefined, state.request)).toBe(false);
    expect(paints).toEqual([]);
    expect(requests).toEqual([]);
    expect(callbacks.size).toBe(0);
  });

  test('paint failure does not advance presentation', () => {
    const failure = new Error('Paint failed');
    const { source, state, errors } = harness(() => { throw failure; });
    expect(source.install(replyFrame(state.request, 3), state.request)).toBe(false);
    expect(source.painted).toBeNull();
    expect(source.coversThroughSeq).toBe(0);
    expect(errors).toEqual([failure]);
  });

  test('a scheduled viewport frame cannot replace newer reply coverage', async () => {
    const { source, state, requests, paints, tick, respond } = harness();
    source.schedule();
    tick();
    expect(source.install(replyFrame(state.request, 2), state.request)).toBe(true);
    await respond(0, frame(requests[0].request, 100));
    expect(source.painted?.sequence).toBe(2);
    expect(paints).toHaveLength(1);
  });

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
