import { beforeAll, describe, expect, test } from 'bun:test';
import { resolve } from 'node:path';
import type { DecodedFrameDelta, FramePageOperation } from '../layout/render/frameDelta';
import type { DisplayPage } from '../layout/render/displayList';
import type { YrsResidentCaretRect } from './index';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

let startWorker: (scope: unknown, canvas: unknown, harness: unknown) => void;

beforeAll(async () => {
  const frameDelta = resolve(import.meta.dir, '../layout/render/frameDelta.ts');
  const modules: Record<string, string> = {
    './residentEngineSession':
      'export const createResidentEngineSession = async () => testHarness.session;',
    '../layout/render/glyphCache': 'export class GlyphCache {}',
    '../layout/render/frameDelta': `
      export { applyFrameDeltaOwned } from ${JSON.stringify(frameDelta)};
      export const decodeFrameDelta = () => testHarness.delta;
    `,
    '../layout/render/canvasBackend': `
      export const rasterizeDisplayPageToBackBuffer = (...args) => testHarness.rasterize(...args);
      export const presentOffscreenPageBackBuffer = (...args) => testHarness.present(...args);
      export const presentOffscreenPageBackBufferWithCaret = (...args) => testHarness.presentCaret(...args);
    `,
  };
  const result = await Bun.build({
    entrypoints: [resolve(import.meta.dir, 'residentEngineWorker.ts')],
    target: 'bun',
    format: 'iife',
    plugins: [
      {
        name: 'isolated-worker-dependencies',
        setup(build) {
          build.onResolve({ filter: /.*/ }, ({ path, importer }) =>
            importer.endsWith('/residentEngineWorker.ts') && path in modules
              ? { path, namespace: 'worker-test' }
              : undefined
          );
          build.onLoad({ filter: /.*/, namespace: 'worker-test' }, ({ path }) => ({
            contents: modules[path],
            loader: 'js',
          }));
        },
      },
    ],
  });
  if (!result.success) throw new AggregateError(result.logs, 'Worker test bundle failed');
  startWorker = new Function(
    'self',
    'OffscreenCanvas',
    'testHarness',
    await result.outputs[0].text()
  ) as typeof startWorker;
});

class Surface {
  pixels: string | null = null;
  constructor(public width = 1, public height = 1) {}
}

function worker() {
  let nextId = 0;
  let frameEpoch = 0;
  const replies = new Map<number, (reply: ResidentEngineWorkerResponse) => void>();
  const surfaces = new Map<string, Surface>();
  const scope = {
    onmessage: (_event: { data: ResidentEngineWorkerRequest }) => {},
    postMessage(reply: ResidentEngineWorkerResponse) {
      replies.get(reply.id)!(reply);
      replies.delete(reply.id);
    },
  };
  const harness = {
    delta: null as DecodedFrameDelta | null,
    caret: null as YrsResidentCaretRect | null,
    rasterized: [] as number[],
    presented: [] as number[],
    failRaster: null as number | null,
    failPresent: null as number | null,
    session: {
      loadState() {},
      clearFonts() {},
      layoutDocumentJson() {},
      onUpdate() {
        return () => {};
      },
      buildDisplayListFrame() {
        return new Uint8Array([frameEpoch]);
      },
      residentCaretSnapshot() {
        return { frameEpoch, caretRect: harness.caret };
      },
      selection() {
        return null;
      },
      encodeStateVector() {
        return new Uint8Array([1]);
      },
      destroy() {},
    },
    async rasterize(
      buffer: Surface,
      page: DisplayPage,
      _options: unknown,
      dpr: number,
      zoom: number
    ) {
      const id = page.pageIndex + 1;
      harness.rasterized.push(id);
      if (harness.failRaster === id) {
        harness.failRaster = null;
        throw new Error('raster failed');
      }
      buffer.width = page.width * dpr * zoom;
      buffer.height = page.height * dpr * zoom;
      buffer.pixels = `${id}:${page.width}`;
    },
    present(canvas: Surface, buffer: Surface) {
      if (!buffer.pixels) throw new Error('presented a detached buffer');
      const id = Number(buffer.pixels.split(':')[0]);
      if (harness.failPresent === id) {
        harness.failPresent = null;
        throw new Error('present failed');
      }
      harness.presented.push(id);
      canvas.pixels = buffer.pixels;
      canvas.width = buffer.width;
      canvas.height = buffer.height;
      buffer.pixels = null;
    },
    presentCaret(canvas: Surface, buffer: Surface, _stage: Surface, caret: { color: string }) {
      if (!buffer.pixels) throw new Error('caret used a detached buffer');
      harness.presented.push(Number(buffer.pixels.split(':')[0]));
      canvas.pixels = `${buffer.pixels}|caret:${caret.color}`;
      canvas.width = buffer.width;
      canvas.height = buffer.height;
    },
  };
  startWorker(scope, Surface, harness);
  function send(request: ResidentEngineWorkerRequestWithoutId) {
    const id = ++nextId;
    return new Promise<ResidentEngineWorkerResponse>((resolve) => {
      replies.set(id, resolve);
      scope.onmessage({
        data: { ...request, id } as ResidentEngineWorkerRequest,
      });
    });
  }
  function delta(upserts: number[], full = false, width = 100, pageCount = 3) {
    const baseFrameEpoch = frameEpoch++;
    const operations: FramePageOperation[] = upserts.map((id) => ({
      kind: 'upsert',
      pageId: BigInt(id),
      pageIndex: id - 1,
      fingerprint: BigInt(frameEpoch),
      primitiveIds: new BigUint64Array(),
      page: { pageIndex: id - 1, width, height: 100, primitives: [] },
    }));
    harness.delta = {
      protocolVersion: 1,
      full,
      frameEpoch,
      baseFrameEpoch,
      docEpoch: frameEpoch,
      layoutEpoch: frameEpoch,
      pageCount,
      operations,
      bytes: new Uint8Array(),
    };
  }
  return {
    harness,
    surfaces,
    send,
    resetCalls() {
      harness.rasterized = [];
      harness.presented = [];
    },
    async bootstrap(pageCount = 3) {
      delta(Array.from({ length: pageCount }, (_, index) => index + 1), true, 100, pageCount);
      return send({
        type: 'bootstrap',
        expectedFrameEpoch: 0,
        extras: '',
        snapshot: {
          clientId: 1,
          state: new Uint8Array(),
          fontsRevision: 0,
          fonts: [],
          renderInputs: [],
          measureInputs: [],
          layoutInput: '',
          layoutWithRegions: false,
          layoutRevision: 1,
          selection: null,
        },
      });
    },
    build(upserts: number[], width = 100, caret: YrsResidentCaretRect | null = null) {
      delta(upserts, false, width);
      harness.caret = caret;
      return send({
        type: 'buildFrame',
        extras: '',
        expectedFrameEpoch: frameEpoch - 1,
        paintCaret: !!caret,
      });
    },
    attach(active: number[], zoom = 1, color = '#000') {
      const pages = active
        .filter((id) => !surfaces.has(String(id)))
        .map((id) => {
          const canvas = new Surface();
          surfaces.set(String(id), canvas);
          return {
            pageId: String(id),
            canvas: canvas as unknown as OffscreenCanvas,
          };
        });
      return send({
        type: 'attachCanvases',
        pages,
        activePageIds: active.map(String),
        devicePixelRatio: 1,
        zoom,
        caretStyle: { color, width: 2 },
      });
    },
  };
}

function caret(page: number): YrsResidentCaretRect {
  return { pageIndex: page - 1, pageId: String(page), x: 5, y: 6, height: 12 };
}

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

describe('resident worker page damage', () => {
  test('paints each page once while a three-page window crosses twelve pages', async () => {
    const w = worker();
    await w.bootstrap(12);
    for (let first = 1; first <= 10; first++) {
      expect((await w.attach([first, first + 1, first + 2])).ok).toBe(true);
    }
    expect(w.harness.rasterized).toEqual(Array.from({ length: 12 }, (_, index) => index + 1));
  });

  test('replays only entering pages after successful presentation', async () => {
    const w = worker();
    expect((await w.bootstrap()).ok).toBe(true);
    expect((await w.attach([1, 2])).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([1, 2]);
    w.resetCalls();
    expect((await w.attach([2, 3])).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([3]);
    expect(w.surfaces.get('1')!.width).toBe(0);
    w.resetCalls();
    await w.attach([2, 3]);
    expect(w.harness.rasterized).toEqual([]);
    await w.attach([1, 2]);
    expect(w.harness.rasterized).toEqual([1]);
  });

  test('retains failed frame damage across a newer clean frame', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([1, 2]);
    w.resetCalls();
    w.harness.failRaster = 1;
    expect((await w.build([1], 120)).ok).toBe(false);
    expect(w.surfaces.get('1')!.pixels).toBe('1:100');
    w.resetCalls();
    expect((await w.build([])).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([1]);
    expect(w.surfaces.get('1')!.pixels).toBe('1:120');
    w.resetCalls();
    await w.build([]);
    expect(w.harness.rasterized).toEqual([]);
  });

  test('finishes rejected-batch writers before a queued frame can reuse their buffers', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([1, 2]);
    const oldStarted = deferred();
    const releaseOld = deferred();
    const oldFinished = deferred();
    const retryStarted = deferred();
    const releaseRetry = deferred();
    const rasterize = w.harness.rasterize;
    let retryWriting = false;
    w.harness.rasterize = async (...args) => {
      const page = args[1];
      if (page.pageIndex === 1 && page.width === 120) {
        oldStarted.resolve();
        await releaseOld.promise;
        await rasterize(...args);
        oldFinished.resolve();
        return;
      }
      if (page.pageIndex === 0 && page.width === 140) {
        retryWriting = true;
        retryStarted.resolve();
        await releaseRetry.promise;
      }
      await rasterize(...args);
    };
    w.harness.failRaster = 1;
    let failedReplyArrived = false;
    const failed = w.build([1, 2], 120).then((reply) => {
      failedReplyArrived = true;
      return reply;
    });
    await oldStarted.promise;
    const retry = w.build([1, 2], 140);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const beforeRelease = { failedReplyArrived, retryWriting };
    releaseOld.resolve();
    await oldFinished.promise;
    await retryStarted.promise;
    releaseRetry.resolve();
    expect((await failed).ok).toBe(false);
    expect((await retry).ok).toBe(true);
    expect(w.surfaces.get('2')!.pixels).toBe('2:140');
    expect(beforeRelease).toEqual({ failedReplyArrived: false, retryWriting: false });
  });

  test('retries every unpresented page after a failed zoom change', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([1, 2]);
    await w.build([]);
    w.harness.failRaster = 2;
    expect((await w.attach([1, 2], 2)).ok).toBe(false);
    w.resetCalls();
    expect((await w.attach([1, 2], 2)).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([1, 2]);
    expect(w.surfaces.get('1')!.width).toBe(200);
    expect(w.surfaces.get('2')!.width).toBe(200);
  });

  test('consumes only pages successfully presented before a presentation failure', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([1, 2]);
    w.harness.failPresent = 2;
    expect((await w.build([1, 2], 130)).ok).toBe(false);
    expect(w.surfaces.get('1')!.pixels).toBe('1:130');
    expect(w.surfaces.get('2')!.pixels).toBe('2:100');
    w.resetCalls();
    expect((await w.build([])).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([2]);
    expect(w.harness.presented).toEqual([2]);
    expect(w.surfaces.get('2')!.pixels).toBe('2:130');
  });

  test('reuses clean caret buffers and rerasterizes detached buffers when gaining a caret', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([1, 2]);
    await w.build([1], 120, caret(1));
    w.resetCalls();
    await w.attach([1, 2], 1, '#f00');
    expect(w.harness.rasterized).toEqual([]);
    expect(w.surfaces.get('1')!.pixels).toBe('1:120|caret:#f00');
    await w.send({ type: 'eraseCaret' });
    expect(w.harness.rasterized).toEqual([]);
    expect(w.surfaces.get('1')!.pixels).toBe('1:120');
    w.resetCalls();
    await w.build([], 100, caret(1));
    expect(w.harness.rasterized).toEqual([1]);
    w.resetCalls();
    await w.build([], 100, caret(2));
    expect(w.harness.rasterized).toEqual([2]);
    expect(w.surfaces.get('1')!.pixels).toBe('1:120');
    expect(w.surfaces.get('2')!.pixels).toBe('2:100|caret:#f00');
  });
});
