import { describe, expect, test } from 'bun:test';
import type { GlyphCache, RetainedFrame } from '@betteroffice/docx/layout/render';
import {
  CanvasReplayState,
  presentCanvasReplay,
  type CanvasRasterEnvironment,
} from './canvasReplay';

const environment: CanvasRasterEnvironment = { dpr: 1, zoom: 1 };
const ids = Array.from({ length: 8 }, (_, index) => BigInt(index + 1));

function frame(epoch: number, damaged = ids): RetainedFrame {
  const pages = ids.map((pageId, pageIndex) => ({
    pageId,
    pageIndex,
    fingerprint: 0n,
    primitiveIds: new BigUint64Array(),
    page: { pageIndex, width: 100, height: 100, primitives: [] },
  }));
  return {
    protocolVersion: 1,
    docEpoch: 1,
    layoutEpoch: epoch,
    frameEpoch: epoch,
    pages,
    damagedPageIds: new Set(damaged),
    removedPageIds: new Set(),
    displayList: { pages: pages.map(({ page }) => page) },
  };
}

function canvas(): HTMLCanvasElement {
  return { width: 100, height: 100 } as HTMLCanvasElement;
}

function pending() {
  let resolve!: () => void;
  let reject!: (error: Error) => void;
  const promise = new Promise<void>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

describe('CanvasReplayState', () => {
  test('scrolling rasters only entering pages after an initial damaged frame', () => {
    const state = new CanvasReplayState();
    const initial = frame(1);
    const surfaces = ids.map(canvas);
    const replay = (start: number, end: number) => {
      state.updateFrame(initial);
      const rastered: bigint[] = [];
      for (let index = start; index <= end; index += 1) {
        const presentation = state.prepare(surfaces[index], ids[index], environment);
        if (presentation) {
          rastered.push(ids[index]);
          state.didPresent(presentation);
        }
      }
      return rastered;
    };
    expect(replay(0, 6)).toEqual(ids.slice(0, 7));
    state.release(surfaces[0]);
    expect(replay(1, 7)).toEqual([8n]);
    expect(replay(1, 7)).toEqual([]);
    expect(replay(0, 6)).toEqual([1n]);
    expect(initial.damagedPageIds.size).toBe(8);
  });

  test('new damage paints once and survives an interrupted replay into a clean frame', () => {
    const state = new CanvasReplayState();
    const surface = canvas();
    state.updateFrame(frame(1));
    state.didPresent(state.prepare(surface, 1n, environment)!);
    state.updateFrame(frame(2, [1n]));
    expect(state.prepare(surface, 1n, environment)).not.toBeNull();
    state.updateFrame(frame(3, []));
    const retry = state.prepare(surface, 1n, environment);
    expect(retry).not.toBeNull();
    state.didPresent(retry!);
    expect(state.prepare(surface, 1n, environment)).toBeNull();
    state.updateFrame(frame(4, [2n]));
    expect(state.prepare(surface, 1n, environment)).toBeNull();
    state.updateFrame(frame(5, [1n]));
    expect(state.prepare(surface, 1n, environment)).not.toBeNull();
  });

  test('environment changes and remounted canvases require a successful new presentation', () => {
    const state = new CanvasReplayState();
    const surface = canvas();
    state.updateFrame(frame(1));
    state.didPresent(state.prepare(surface, 1n, environment)!);
    const environments: CanvasRasterEnvironment[] = [
      { ...environment, dpr: 2 },
      { ...environment, zoom: 2 },
      { ...environment, glyphCache: {} as GlyphCache },
      { ...environment, resolveImage: async () => null },
    ];
    for (const changed of environments) {
      expect(state.prepare(surface, 1n, changed)).not.toBeNull();
      const retry = state.prepare(surface, 1n, changed);
      expect(retry).not.toBeNull();
      state.didPresent(retry!);
      expect(state.prepare(surface, 1n, changed)).toBeNull();
    }
    expect(state.prepare(canvas(), 1n, environments.at(-1)!)).not.toBeNull();
  });

  test('skipped frames and document changes cannot hide unobserved damage', () => {
    const state = new CanvasReplayState();
    const surface = canvas();
    state.updateFrame(frame(1));
    state.didPresent(state.prepare(surface, 1n, environment)!);
    state.updateFrame(frame(3, []));
    const skipped = state.prepare(surface, 1n, environment);
    expect(skipped).not.toBeNull();
    state.didPresent(skipped!);
    state.updateFrame({ ...frame(4, []), docEpoch: 2 });
    expect(state.prepare(surface, 1n, environment)).not.toBeNull();
  });

  test('legacy lists still repaint and reused canvases do not inherit another page', () => {
    const state = new CanvasReplayState();
    const surface = canvas();
    state.updateFrame(frame(1));
    state.didPresent(state.prepare(surface, 1n, environment)!);
    expect(state.prepare(surface, 2n, environment)).not.toBeNull();
    state.updateFrame(null);
    state.didPresent(state.prepare(surface, undefined, environment)!);
    expect(state.prepare(surface, undefined, environment)).not.toBeNull();
  });
});

describe('presentCanvasReplay', () => {
  test('presents the complete batch before releasing only its scratch buffers', async () => {
    const slow = pending();
    const surface = canvas();
    const buffers = [canvas(), canvas()];
    let presentations = 0;
    const replay = presentCanvasReplay(
      buffers.map((buffer, index) => ({
        buffer,
        ready: index === 0 ? Promise.resolve() : slow.promise,
        present() {
          expect(buffer.width).toBe(100);
          surface.width = buffer.width;
          presentations += 1;
        },
      })),
      () => true
    );
    await Promise.resolve();
    expect(presentations).toBe(0);
    expect(buffers.map(({ width }) => width)).toEqual([100, 100]);
    slow.resolve();
    await replay;
    expect(presentations).toBe(2);
    expect(buffers.map(({ width, height }) => [width, height])).toEqual([
      [0, 0],
      [0, 0],
    ]);
    expect(surface.width).toBe(100);
  });

  test('stale or unmounted work releases scratch without marking a presentation', async () => {
    const state = new CanvasReplayState();
    const surface = canvas();
    const buffer = canvas();
    const slow = pending();
    state.updateFrame(frame(1));
    const presentation = state.prepare(surface, 1n, environment)!;
    let current = true;
    const replay = presentCanvasReplay(
      [
        {
          buffer,
          ready: slow.promise,
          present: () => state.didPresent(presentation),
        },
      ],
      () => current
    );
    current = false;
    expect(buffer.width).toBe(100);
    slow.resolve();
    await replay;
    expect(buffer.width).toBe(0);
    expect(state.prepare(surface, 1n, environment)).not.toBeNull();
  });

  test('raster failure waits for in-flight writers before releasing the whole batch', async () => {
    const failed = pending();
    const slow = pending();
    const buffers = [canvas(), canvas()];
    let presentations = 0;
    const replay = presentCanvasReplay(
      buffers.map((buffer, index) => ({
        buffer,
        ready: index === 0 ? failed.promise : slow.promise,
        present() {
          presentations += 1;
        },
      })),
      () => true
    );
    const rejection = replay.catch((error: Error) => error);
    failed.reject(new Error('raster failed'));
    await Promise.resolve();
    expect(buffers.map(({ width }) => width)).toEqual([100, 100]);
    slow.resolve();
    expect(await rejection).toEqual(new Error('raster failed'));
    expect(presentations).toBe(0);
    expect(buffers.map(({ width }) => width)).toEqual([0, 0]);
  });

  test('presentation failure also releases every scratch buffer', async () => {
    const buffers = [canvas(), canvas()];
    await expect(
      presentCanvasReplay(
        buffers.map((buffer) => ({
          buffer,
          ready: Promise.resolve(),
          present() {
            throw new Error('presentation failed');
          },
        })),
        () => true
      )
    ).rejects.toThrow('presentation failed');
    expect(buffers.map(({ width }) => width)).toEqual([0, 0]);
  });
});
