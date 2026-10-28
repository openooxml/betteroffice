import { expect, test } from 'bun:test';
import { createRenderPresenter, type RenderPresenterOptions } from './presenter';
import { createBitmapOwnership } from './ownership';
import type { ConsumedMessage, RenderFrame, ResyncMessage } from './types';

function bitmap() { return { width: 2, height: 2, closed: 0, close() { this.closed += 1; } }; }
type Bitmap = ReturnType<typeof bitmap>;
function frame(epoch: number, extra: Partial<RenderFrame<Bitmap>> = {}): RenderFrame<Bitmap> {
  return { session: 'session', incarnation: 1, epoch, coversThroughSeq: epoch, viewportGeneration: 1,
    scale: 1, retire: [], progressive: false, geometryRevision: {},
    patches: [{ surface: { kind: 'page', id: '1' }, rect: { x: 0, y: 0, width: 2, height: 2 },
      bitmap: bitmap(), full: true }], ...extra };
}
function host(hooks: Partial<Pick<RenderPresenterOptions<Bitmap, number>, 'validate' | 'present'>> = {}) {
  const turns = new Map<number, () => void>(), cancelled: number[] = [];
  const consumed: ConsumedMessage[] = [], resync: ResyncMessage[] = [], errors: unknown[] = [];
  const presented: number[] = [], ownership = createBitmapOwnership<Bitmap>();
  let handle = 0, stalls = 0;
  const presenter = createRenderPresenter({
    ownership,
    requestFrame(callback) { const id = handle++; turns.set(id, callback); return id; },
    cancelFrame(id: number) { cancelled.push(id); turns.delete(id); },
    validate: hooks.validate ?? (() => 'ok'),
    present(value) { const result = hooks.present?.(value); presented.push(value.epoch); return result; },
    consumed(message) { consumed.push(message); },
    resync(message) { resync.push(message); },
    onError(error) { errors.push(error); },
    onStall() { stalls += 1; },
  });
  return { presenter, consumed, resync, errors, presented, ownership, turns, cancelled,
    get stalls() { return stalls; },
    tick() {
      expect(turns.size).toBe(1);
      const [id, callback] = [...turns][0]!;
      turns.delete(id);
      callback();
    },
  };
}

test('30 arrivals in 10 callbacks present only the latest, consume once and close every bitmap', () => {
  const h = host(), frames: RenderFrame<Bitmap>[] = [];
  for (let turn = 0; turn < 10; turn += 1) {
    for (let offset = 1; offset <= 3; offset += 1) {
      const next = frame(turn * 3 + offset);
      frames.push(next);
      h.presenter.receive(next);
    }
    h.tick();
    expect(h.presenter.counters.presented).toBe(turn + 1);
  }
  expect(h.presented).toEqual([3, 6, 9, 12, 15, 18, 21, 24, 27, 30]);
  expect(h.presenter.counters).toEqual({ received: 30, presented: 10, superseded: 20, discarded: 0 });
  expect(h.consumed.length).toBe(30);
  expect(new Set(h.consumed.map(({ epoch }) => epoch)).size).toBe(30);
  for (const message of h.consumed) expect(message.status).toBe(message.epoch % 3 === 0 ? 'presented' : 'superseded');
  for (const next of frames) expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('stale validation discards and closes patches and plates without resync', () => {
  const h = host({ validate: () => 'stale' }), next = frame(1);
  next.plates = [{ ...next.patches[0]!, bitmap: bitmap() }];
  h.presenter.receive(next);
  h.tick();
  expect(h.consumed[0]?.status).toBe('discarded');
  expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(next.plates[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
  expect(h.resync).toEqual([]);
});

test('adapter validation mismatches discard the frame and request its patched surfaces', () => {
  const h = host({ validate: () => 'mismatch' }), next = frame(1);
  h.presenter.receive(next); h.tick();
  expect(h.consumed[0]?.status).toBe('discarded');
  expect(h.resync[0]?.surface).toEqual(next.patches[0]!.surface);
  expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('base mismatches request twice, stall once and re-arm on invalidation', () => {
  const h = host(), frames: RenderFrame<Bitmap>[] = [];
  h.presenter.receive(frame(1)); h.tick();
  function mismatch(epoch: number) {
    const next = frame(epoch, { baseEpoch: 0 });
    next.patches[0]!.full = false;
    frames.push(next);
    h.presenter.receive(next); h.tick();
  }
  for (let epoch = 2; epoch <= 5; epoch += 1) mismatch(epoch);
  expect(h.resync).toEqual(Array.from({ length: 2 }, () => ({ surface: { kind: 'page', id: '1' }, reason: 'base-mismatch' })));
  expect(h.stalls).toBe(1);
  h.presenter.invalidate();
  mismatch(6);
  expect(h.resync.length).toBe(3);
  h.presenter.receive(frame(7)); h.tick();
  mismatch(8);
  expect(h.resync.length).toBe(4);
  expect(h.presenter.counters.discarded).toBe(6);
  for (const next of frames) expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('presented frame epoch and surface membership reset across incarnations', () => {
  const epochs: (number | undefined)[] = [], surfaces: string[][] = [];
  const h = host({ validate(_frame, presented, epoch) { epochs.push(epoch); surfaces.push([...presented]); return 'ok'; } });
  h.presenter.receive(frame(7)); h.tick();
  const partial = frame(8, { baseEpoch: 7 }); partial.patches[0]!.full = false;
  h.presenter.receive(partial); h.tick();
  h.presenter.receive(frame(6)); h.tick();
  h.presenter.receive(frame(1, { incarnation: 2 })); h.tick();
  h.presenter.receive(frame(99)); h.tick();
  expect(h.presented).toEqual([7, 8, 1]);
  expect(epochs).toEqual([undefined, 7, 8, undefined, undefined]);
  expect(surfaces).toEqual([[], ['page:1'], ['page:1'], [], []]);
  expect(h.presenter.counters.discarded).toBe(2);
  expect(h.ownership.openCount).toBe(0);
});

test('alternating surfaces use the last presented frame base and reject missing surfaces or older bases', () => {
  const h = host(), first = frame(1);
  first.patches[0]!.surface.id = 'A';
  first.patches.push({ ...first.patches[0]!, surface: { kind: 'page', id: 'B' }, bitmap: bitmap() });
  const frames = [first];
  h.presenter.receive(first); h.tick();
  for (const [epoch, baseEpoch, id] of [[2, 1, 'A'], [3, 2, 'B'], [4, 3, 'C'], [5, 1, 'B']] as const) {
    const next = frame(epoch, { baseEpoch });
    next.patches[0]!.surface.id = id; next.patches[0]!.full = false;
    frames.push(next);
    h.presenter.receive(next); h.tick();
  }
  expect(h.presented).toEqual([1, 2, 3]);
  expect(h.consumed.map(({ status }) => status)).toEqual(['presented', 'presented', 'presented', 'discarded', 'discarded']);
  expect(h.resync.map(({ surface }) => surface?.id)).toEqual(['C', 'B']);
  for (const next of frames) for (const patch of next.patches) expect(patch.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('throwing presentation discards, closes all bitmaps, requests all surfaces and reports the error', () => {
  const error = new Error('install failed'), h = host({ present() { throw error; } }), next = frame(1);
  next.plates = [{ ...next.patches[0]!, bitmap: bitmap() }];
  h.presenter.receive(next);
  expect(() => h.tick()).not.toThrow();
  expect(h.consumed.map(({ status }) => status)).toEqual(['discarded']);
  expect(h.resync).toEqual([{ reason: 'present-failed' }]);
  expect(h.errors).toEqual([error]);
  expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(next.plates[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('a cast thenable presentation is discarded synchronously without chaining its result', () => {
  type AcceptsDeferred = ((frame: RenderFrame<Bitmap>) => Promise<void>) extends RenderPresenterOptions<Bitmap, number>['present'] ? true : false;
  const acceptsDeferred: AcceptsDeferred = false;
  expect(acceptsDeferred).toBe(false);
  let chained = 0;
  const present = (() => ({ then() { chained += 1; } })) as unknown as RenderPresenterOptions<Bitmap, number>['present'];
  const h = host({ present }), next = frame(1);
  next.plates = [{ ...next.patches[0]!, bitmap: bitmap() }];
  h.presenter.receive(next); h.tick();
  expect(h.presenter.counters.presented).toBe(0);
  expect(h.consumed.map(({ status }) => status)).toEqual(['discarded']);
  expect(h.resync).toEqual([{ reason: 'present-failed' }]);
  expect(h.errors[0]).toBeInstanceOf(Error);
  expect((h.errors[0] as Error).message).toBe('present must be synchronous');
  expect(chained).toBe(0);
  expect(next.patches[0]!.bitmap.closed).toBe(1);
  expect(next.plates[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('disposal cancels handle zero, discards waiting frames and rejects later arrivals', () => {
  const h = host(), waiting = frame(1), later = frame(2);
  waiting.plates = [{ ...waiting.patches[0]!, bitmap: bitmap() }];
  h.presenter.receive(waiting);
  h.presenter.dispose(); h.presenter.dispose();
  h.presenter.receive(later);
  expect(h.cancelled).toEqual([0]);
  expect(h.turns.size).toBe(0);
  expect(h.consumed.map(({ status }) => status)).toEqual(['discarded', 'discarded']);
  expect(waiting.patches[0]!.bitmap.closed).toBe(1);
  expect(waiting.plates[0]!.bitmap.closed).toBe(1);
  expect(later.patches[0]!.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('presentation, acknowledgement and counters complete inside the callback; caller releases plates', () => {
  const h = host(), next = frame(1), plate = { ...next.patches[0]!, bitmap: bitmap() };
  next.plates = [plate];
  h.presenter.receive(next);
  expect(h.presented).toEqual([]);
  h.tick();
  expect(h.presented).toEqual([1]);
  expect(h.presenter.counters.presented).toBe(1);
  expect(h.consumed[0]?.status).toBe('presented');
  expect(h.ownership.openCount).toBe(1);
  expect(plate.bitmap.closed).toBe(0);
  h.ownership.release(plate.bitmap, plate);
  expect(plate.bitmap.closed).toBe(1);
  expect(h.ownership.openCount).toBe(0);
});

test('retirement removes surface membership and shared patch/plate resources close once', () => {
  const h = host(), next = frame(1), patch = next.patches[0]!;
  next.plates = [patch];
  h.presenter.receive(next); h.tick();
  h.ownership.release(patch.bitmap, patch);
  expect(patch.bitmap.closed).toBe(1);
  h.presenter.receive(frame(2, { patches: [], retire: [patch.surface] })); h.tick();
  const delta = frame(3, { baseEpoch: 2 }); delta.patches[0]!.full = false;
  h.presenter.receive(delta); h.tick();
  expect(h.resync[0]?.surface).toEqual(patch.surface);
  expect(h.ownership.openCount).toBe(0);
});
