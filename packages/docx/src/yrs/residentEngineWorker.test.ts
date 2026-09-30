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
  const answered: number[] = [];
  const surfaces = new Map<string, Surface>();
  const scope = {
    onmessage: (_event: { data: ResidentEngineWorkerRequest }) => {},
    postMessage(reply: ResidentEngineWorkerResponse) {
      answered.push(reply.id);
      replies.get(reply.id)?.(reply);
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
      setPartialDocument() {},
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
    answered,
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

describe('resident worker layout ownership', () => {
  test('returns the layout it ran and completes the frame extras from it', async () => {
    const w = worker();
    const extras: string[] = [];
    const layoutJson = JSON.stringify({
      layout: { pages: [] },
      headersFooters: { parts: [] },
      notesConverged: true,
    });
    let epoch = 0;
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsRetainedJson: () => layoutJson,
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: (input: string) => {
        extras.push(input);
        epoch += 1;
        w.harness.delta = {
          protocolVersion: 1,
          full: true,
          frameEpoch: epoch,
          baseFrameEpoch: 0,
          docEpoch: epoch,
          layoutEpoch: epoch,
          pageCount: 0,
          operations: [],
          bytes: new Uint8Array(),
        };
        return new Uint8Array([0]);
      },
    });
    const snapshot = {
      clientId: 1,
      state: new Uint8Array(),
      fontsRevision: 0,
      fonts: [],
      renderInputs: [],
      measureInputs: [],
      layoutInput: '{}',
      layoutWithRegions: true,
      layoutRevision: 1,
      selection: null,
    };
    const reply = await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: 'unused',
      snapshot,
      layoutExtras: JSON.stringify({ resolvedCommentIds: [4], fontChains: { 'a|0|0': [1] } }),
    });
    expect(reply.ok && reply.layoutJson).toBe(layoutJson);
    expect(extras).toEqual([
      '{"headersFooters":{"parts":[]},"fontChains":{"a|0|0":[1]},"resolvedCommentIds":[4]}',
    ]);

    const plain = await w.send({
      type: 'sync',
      expectedFrameEpoch: 0,
      extras: 'given',
      paintCaret: false,
      snapshot,
    });
    expect(plain.ok && plain.layoutJson).toBeUndefined();
    expect(extras.at(-1)).toBe('given');
  });

  test('marks a replica as a preview or not before it lays it out', async () => {
    const w = worker();
    const calls: string[] = [];
    Object.assign(w.harness.session, {
      loadState: () => calls.push('load'),
      setPartialDocument: (partial: boolean) => calls.push(`partial:${partial}`),
      layoutDocumentWithRegionsRetainedJson: () => {
        calls.push('layout');
        return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
      },
    });
    const snapshot = {
      clientId: 1,
      state: new Uint8Array(),
      fontsRevision: 0,
      fonts: [],
      renderInputs: [],
      measureInputs: [],
      layoutInput: '{}',
      layoutWithRegions: true,
      layoutRevision: 1,
      selection: null,
    };
    await w.send({ type: 'bootstrap', expectedFrameEpoch: 0, extras: '{}', snapshot });
    expect(calls).toEqual(['load', 'partial:false', 'layout']);
    calls.length = 0;
    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '{}',
      snapshot: { ...snapshot, partialDocument: true },
    });
    expect(calls).toEqual(['load', 'partial:true', 'layout']);
    calls.length = 0;
    // A complete document synced into the same session is no longer a preview's.
    await w.send({ type: 'sync', expectedFrameEpoch: 0, extras: '{}', paintCaret: false, snapshot });
    expect(calls).toEqual(['load', 'partial:false', 'layout']);
  });

  test('finishes a provisional layout on request and before other work', async () => {
    const w = worker();
    const extras: string[] = [];
    const calls: string[] = [];
    let epoch = 0;
    const provisional = '{"layout":{"pages":[1]},"notesConverged":true,"provisional":true}';
    const full = '{"layout":{"pages":[1,2]},"notesConverged":true}';
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsPrefixRetainedJson: (_input: string, pages: number) => {
        calls.push(`prefix:${pages}`);
        return provisional;
      },
      layoutDocumentWithRegionsRetainedJson: () => {
        calls.push('full');
        return full;
      },
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: (input: string) => {
        extras.push(input);
        epoch += 1;
        w.harness.delta = {
          protocolVersion: 1,
          full: true,
          frameEpoch: epoch,
          baseFrameEpoch: 0,
          docEpoch: epoch,
          layoutEpoch: epoch,
          pageCount: 0,
          operations: [],
          bytes: new Uint8Array(),
        };
        return new Uint8Array([0]);
      },
    });
    const snapshot = {
      clientId: 1,
      state: new Uint8Array(),
      fontsRevision: 0,
      fonts: [],
      renderInputs: [],
      measureInputs: [],
      layoutInput: '{}',
      layoutWithRegions: true,
      layoutRevision: 1,
      selection: null,
    };
    const bootstrap = await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '',
      snapshot,
      layoutExtras: '{}',
      provisionalPages: 3,
    });
    expect(bootstrap.ok && bootstrap.layoutJson).toBe(provisional);
    expect(bootstrap.ok && bootstrap.layoutProvisional).toBe(true);
    expect(calls).toEqual(['prefix:3']);

    const completed = await w.send({ type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(completed.ok && completed.layoutProvisional).toBeUndefined();
    expect(calls).toEqual(['prefix:3', 'full']);
    expect(extras).toHaveLength(2);

    const again = await w.send({ type: 'completeLayout', expectedFrameEpoch: 2, paintCaret: false });
    expect(again.ok && again.frame).toBeUndefined();

    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '',
      snapshot,
      layoutExtras: '{}',
      provisionalPages: 3,
    });
    await w.send({ type: 'buildFrame', extras: 'given', expectedFrameEpoch: 3, paintCaret: false });
    expect(calls.slice(2)).toEqual(['prefix:3', 'full']);
    const late = await w.send({ type: 'completeLayout', expectedFrameEpoch: 4, paintCaret: false });
    expect(late.ok && late.layoutJson).toBe(full);
    expect(calls).toHaveLength(4);
  });
});

describe('sliced layout completion', () => {
  const provisional = '{"layout":{"pages":[1]},"notesConverged":true,"provisional":true}';
  const full = '{"layout":{"pages":[1,2]},"notesConverged":true}';
  const snapshot = {
    clientId: 1,
    state: new Uint8Array(),
    fontsRevision: 0,
    fonts: [],
    renderInputs: [],
    measureInputs: [],
    layoutInput: '{}',
    layoutWithRegions: true,
    layoutRevision: 1,
    selection: null,
  };

  function steppedWorker(bodyBlocks = 10) {
    const w = worker();
    const calls: string[] = [];
    let epoch = 0;
    let measured = 0;
    let begun = false;
    let changed = false;
    const onResume: Array<() => void> = [];
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsPrefixRetainedJson: () => provisional,
      layoutDocumentWithRegionsRetainedJson: () => {
        calls.push('whole');
        return full;
      },
      beginRegionLayout: () => {
        calls.push('begin');
        begun = true;
        changed = false;
        measured = 0;
        return { measuredBlocks: 0, bodyBlocks };
      },
      resumeRegionLayout: (blocks: number) => {
        calls.push(`resume:${Math.min(blocks, bodyBlocks)}`);
        onResume.shift()?.();
        if (!begun || changed) {
          begun = false;
          throw new Error('the document or its fonts changed since the region layout began');
        }
        measured = Math.min(bodyBlocks, measured + blocks);
        if (measured < bodyBlocks) return { measuredBlocks: measured, bodyBlocks };
        begun = false;
        return { measuredBlocks: measured, bodyBlocks, layoutJson: full };
      },
      applyUpdate: () => {
        calls.push('update');
        changed = true;
        return null;
      },
      buildDisplayPagesFrame: () => {
        calls.push('pages');
        return new Uint8Array([0]);
      },
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: () => {
        epoch += 1;
        // One page, as a provisional frame shows its prefix.
        w.harness.delta = {
          protocolVersion: 1,
          full: true,
          frameEpoch: epoch,
          baseFrameEpoch: 0,
          docEpoch: epoch,
          layoutEpoch: epoch,
          pageCount: 1,
          operations: [
            {
              kind: 'upsert',
              pageId: 1n,
              pageIndex: 0,
              fingerprint: BigInt(epoch),
              primitiveIds: new BigUint64Array(),
              page: { pageIndex: 0, width: 100, height: 100, primitives: [] },
            },
          ],
          bytes: new Uint8Array(),
        };
        return new Uint8Array([0]);
      },
    });
    const bootstrap = () =>
      w.send({
        type: 'bootstrap',
        expectedFrameEpoch: 0,
        extras: '',
        snapshot,
        layoutExtras: '{}',
        provisionalPages: 3,
      });
    return { w, calls, onResume, bootstrap };
  }

  test('measures the rest in steps, running requests that arrive meanwhile between them', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker();
    await bootstrap();
    const order: string[] = [];
    onResume.push(() => {
      void w
        .send({ type: 'buildPages', pages: [0], expectedFrameEpoch: 1, paintCaret: false })
        .then(() => order.push('pages'));
    });
    const completed = await w
      .send({ type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false, sliceBlocks: 3 })
      .then((reply) => {
        order.push('complete');
        return reply;
      });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(order).toEqual(['pages', 'complete']);
    expect(calls[0]).toBe('begin');
    expect(calls[1]).toBe('resume:3');
    expect(calls[2]).toBe('pages');
    expect(calls.slice(3).every((call) => call.startsWith('resume:'))).toBe(true);
    expect(calls).not.toContain('whole');
  });

  test('an update in between begins the pass again on the new state', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker();
    await bootstrap();
    onResume.push(() => {
      void w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null });
    });
    const completed = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 4,
    });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(calls.filter((call) => call === 'begin')).toHaveLength(2);
    expect(calls.indexOf('update')).toBeLessThan(calls.lastIndexOf('begin'));
    expect(calls).not.toContain('whole');
  });

  test('a frame request finishes the pass first and answers it first', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker(100);
    await bootstrap();
    const order: string[] = [];
    let frame: Promise<unknown> | undefined;
    onResume.push(() => {
      frame = w
        .send({ type: 'buildFrame', extras: 'given', expectedFrameEpoch: 2, paintCaret: false })
        .then(() => order.push('frame'));
    });
    const completed = await w
      .send({ type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false, sliceBlocks: 2 })
      .then((reply) => {
        order.push('complete');
        return reply;
      });
    await frame;
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(order).toEqual(['complete', 'frame']);
    expect(calls).toEqual(['begin', 'resume:2', 'resume:100']);
  });

  test('a completion that fails when a frame request finishes it is answered once, as is the frame request', async () => {
    const { w, onResume, bootstrap } = steppedWorker(100);
    await bootstrap();
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsRetainedJson: () => {
        throw new Error('layout failed');
      },
    });
    let frame: Promise<ResidentEngineWorkerResponse> | undefined;
    onResume.push(() => {
      (w.harness.session as { resumeRegionLayout?: unknown }).resumeRegionLayout = () => {
        throw new Error('layout failed');
      };
      frame = w.send({ type: 'buildFrame', extras: 'given', expectedFrameEpoch: 2, paintCaret: false });
    });
    const completion = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 2,
    });
    const framed = await frame!;
    expect(completion.ok).toBe(false);
    expect(framed.ok).toBe(false);
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(new Set(w.answered).size).toBe(w.answered.length);
  });

  test('a step queued behind a request that traps finishing the pass leaves the completion answered once', async () => {
    const { w, onResume, bootstrap } = steppedWorker(100);
    const build = w.harness.session.buildDisplayListFrame;
    w.harness.session.buildDisplayListFrame = () => {
      const bytes = build();
      w.harness.delta = {
        ...w.harness.delta!,
        pageCount: 1,
        operations: [
          {
            kind: 'upsert',
            pageId: 1n,
            pageIndex: 0,
            fingerprint: BigInt(w.harness.delta!.frameEpoch),
            primitiveIds: new BigUint64Array(),
            page: { pageIndex: 0, width: 100, height: 100, primitives: [] },
          },
        ],
      };
      return bytes;
    };
    await bootstrap();
    const raster = deferred();
    const rasterize = w.harness.rasterize;
    let frame: Promise<ResidentEngineWorkerResponse> | undefined;
    onResume.push(() => {
      w.harness.rasterize = async (...args) => {
        await raster.promise;
        return rasterize(...args);
      };
      (w.harness.session as { resumeRegionLayout?: unknown }).resumeRegionLayout = () => {
        throw new WebAssembly.RuntimeError('unreachable');
      };
      // The attach holds the queue while the next step is queued behind the frame request.
      void w.attach([1]);
      frame = w.send({ type: 'buildFrame', extras: 'given', expectedFrameEpoch: 2, paintCaret: false });
      setTimeout(() => raster.resolve(), 20);
    });
    const completion = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 2,
    });
    const framed = await frame!;
    expect(completion.ok).toBe(false);
    expect(framed.ok).toBe(false);
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(new Set(w.answered).size).toBe(w.answered.length);
  });

  test("a superseded completion's queued step never runs the next completion", async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker(12);
    await bootstrap();
    let second: Promise<ResidentEngineWorkerResponse> | undefined;
    onResume.push(() => {
      void bootstrap();
      second = w.send({
        type: 'completeLayout',
        expectedFrameEpoch: 1,
        paintCaret: false,
        sliceBlocks: 4,
      });
    });
    const first = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 4,
    });
    const next = await second!;
    expect(first.ok && first.frame).toBeUndefined();
    expect(next.ok && next.layoutJson).toBe(full);
    expect(calls.filter((call) => call === 'begin')).toHaveLength(2);
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(new Set(w.answered).size).toBe(w.answered.length);
  });

  test('a new snapshot supersedes the pass', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker();
    await bootstrap();
    onResume.push(() => {
      void w.send({
        type: 'sync',
        expectedFrameEpoch: 1,
        extras: '{}',
        paintCaret: false,
        snapshot: { ...snapshot, layoutWithRegions: false },
      });
    });
    const superseded = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 2,
    });
    expect(superseded.ok && superseded.frame).toBeUndefined();
    expect(calls).toEqual(['begin', 'resume:2']);
  });

  test('without a slice size the rest is laid out in one step', async () => {
    const { w, calls, bootstrap } = steppedWorker();
    await bootstrap();
    const completed = await w.send({ type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(calls).toEqual(['whole']);
  });
});
