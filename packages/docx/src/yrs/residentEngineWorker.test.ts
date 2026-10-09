import { afterEach, beforeAll, beforeEach, describe, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type { ResidentSaveRecord } from './residentSave';
import { syntheticDocx } from './__fixtures__/previewChain';
import { isLayoutMetaV1, type LayoutMetaV1 } from './layoutMeta';
import type { Layout } from '../layout/pagination';
import { proposalRevisionPreview } from './proposals';
import { createYrsSession } from './index';
import { readSidebar, readOutlineHeadings } from './sidebarReads';
import { sidebarDocx } from './__fixtures__/sidebarDocx';
import { readResidentSearch } from './residentSearch';
import { findBodyMatches } from './findMatches';
import { createYrsInputPositionMap } from './inputPositionMap';
import { createYrsPositionProjection, yrsLocToProjectedDisplayPosition } from './yrsPositionProjection';
import { preloadEditWasm } from './wasm/index';
import { PeerMetadataError } from './peerMetadata';
import type { DecodedFrameDelta, FramePageOperation } from '../layout/render/frameDelta';
import type { DisplayPage } from '../layout/render/displayList';
import type {
  DocxEditRequest,
  DocxContentControlsResult,
  DocxParagraphAnchor,
  DocxProposalInput,
  YrsResidentCaretRect,
  YrsResidentWorkerSnapshot,
} from './index';
import {
  RESIDENT_HOST_MODULE_WAIT_MS,
  type ResidentEngineWorkerHostModule,
  type ResidentEngineWorkerRequest,
  type ResidentEngineWorkerRequestWithoutId,
  type ResidentEngineWorkerResponse,
  type ResidentDocumentRead,
  type ResidentProposalOperation,
} from './residentEngineWorkerProtocol';

let startWorker: (scope: unknown, canvas: unknown, harness: unknown, clock: unknown) => void;

beforeAll(async () => {
  const frameDelta = resolve(import.meta.dir, '../layout/render/frameDelta.ts');
  const modules: Record<string, string> = {
    './residentEngineSession':
      `export const createResidentEngineSession = async (heapLimitBytes) => {
        await testHarness.preload();
        testHarness.sessionsCreated += 1;
        (testHarness.heapLimits ??= []).push(heapLimitBytes);
        return testHarness.session;
      };`,
    './wasm/index': `
      export const preloadEditWasm = () => testHarness.preload();
      export const preloadEditWasmFrom = (source) => testHarness.preloadFrom(source);
    `,
    '../layout/render/glyphCache':
      'export class GlyphCache { constructor(options) { testHarness.glyphs = options.provider; testHarness.glyphCacheCreations += 1; } }',
    '../wasm/loadWasmAsset': 'export const wasmModuleMemories = () => testHarness.memories;',
    '../layout/render/frameDelta': `
      export { applyFrameDeltaOwned, retainedFramePageById } from ${JSON.stringify(frameDelta)};
      export const decodeFrameDelta = () => testHarness.delta;
    `,
    '../layout/render/canvasBackend': `
      export const rasterizeDisplayPageToBackBuffer = (...args) => testHarness.rasterize(...args);
      export const presentOffscreenPageBackBuffer = (...args) => testHarness.present(...args);
      export const presentOffscreenPageBackBufferWithCaret = (...args) => testHarness.presentCaret(...args);
      export const releaseOffscreenPageCanvas = (...args) => testHarness.release(...args);
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
    'performance',
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
  const transfers = new Map<number, Transferable[]>();
  const surfaces = new Map<string, Surface>();
  const scope = {
    onmessage: (_event: { data: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule }) => {},
    onmessageerror: null as (() => void) | null,
    postMessage(reply: ResidentEngineWorkerResponse, transfer: Transferable[] = []) {
      transfers.set(reply.id, transfer);
      answered.push(reply.id);
      replies.get(reply.id)?.(reply);
      replies.delete(reply.id);
    },
  };
  const harness = {
    now: () => performance.now(),
    initializations: 0,
    glyphCacheCreations: 0,
    clearFontCalls: 0,
    fontIds: [] as number[],
    loadedStates: [] as Uint8Array[],
    loadedMediaSources: [] as string[],
    partialDocuments: [] as boolean[],
    sessionsCreated: 0,
    wasmReady: false,
    failWarm: null as Error | null,
    preloadBlock: null as Promise<void> | null,
    preloadInputs: [] as (WebAssembly.Module | undefined)[],
    async preloadFrom(source: Promise<WebAssembly.Module | null>): Promise<void> {
      await harness.preload((await source) ?? undefined);
    },
    async preload(input?: WebAssembly.Module): Promise<void> {
      harness.preloadInputs.push(input);
      await harness.preloadBlock;
      if (harness.failWarm) {
        const error = harness.failWarm;
        harness.failWarm = null;
        throw error;
      }
      if (!harness.wasmReady) {
        harness.initializations += 1;
        harness.wasmReady = true;
      }
    },
    delta: null as DecodedFrameDelta | null,
    caret: null as YrsResidentCaretRect | null,
    displayWindows: [] as [number, number][],
    retainBuiltPages: [] as boolean[],
    windowedIncrementalBuilds: [] as boolean[],
    directBatches: [] as boolean[],
    rasterized: [] as number[],
    buffers: [] as Surface[],
    releaseCalls: [] as Array<{ pages: number[]; expectedFrameEpoch: number }>,
    releaseSuperseded: false,
    presented: [] as number[],
    failRaster: null as number | null,
    failPresent: null as number | null,
    memories: [{ label: 'docx-edit', bufferBytes: 65536, liveBytes: 100, peakBytes: 100, failedAllocationBytes: 0 }],
    session: {
      proposalEngine: { version: () => 'v' },
      markProjectionStories(_stories: readonly string[]) {},
      loadState(state: Uint8Array) {
        harness.loadedStates.push(state);
      },
      loadMediaSources(json: string) {
        harness.loadedMediaSources.push(json);
      },
      loadNoteSeparators(_state: Uint8Array) {},
      setPartialDocument(partial: boolean) {
        harness.partialDocuments.push(partial);
      },
      setDisplayWindow(start: number, end: number) {
        harness.displayWindows.push([start, end]);
      },
      setDisplayRetainBuiltPages(retain: boolean) {
        harness.retainBuiltPages.push(retain);
      },
      setWindowedIncrementalBuilds(enabled: boolean) {
        harness.windowedIncrementalBuilds.push(enabled);
      },
      setDirectBatches(enabled: boolean) {
        harness.directBatches.push(enabled);
      },
      directBatchesApplied() {
        return 0;
      },
      clearFonts() {
        harness.clearFontCalls += 1;
        harness.fontIds = [];
      },
      registerFont(_bytes: Uint8Array) {
        const id = harness.fontIds.length;
        harness.fontIds.push(id);
        return id;
      },
      registerSubstituteFont(_base: number, _family: string) {
        const id = harness.fontIds.length;
        harness.fontIds.push(id);
        return id;
      },
      layoutDocumentJson() {},
      layoutDocumentWithRegionsRetained() {},
      retainedHeadersFootersJson(): string | undefined {
        return undefined;
      },
      onUpdate() {
        return () => {};
      },
      buildDisplayListFrame() {
        return new Uint8Array([frameEpoch]);
      },
      buildDisplayPagesFrame(pages: number[], _expectedFrameEpoch: number) {
        delta(pages.map((index) => index + 1));
        return new Uint8Array([frameEpoch]);
      },
      releaseDisplayPagesFrame(pages: number[], expectedFrameEpoch: number): Uint8Array | null {
        harness.releaseCalls.push({ pages, expectedFrameEpoch });
        if (harness.releaseSuperseded || expectedFrameEpoch !== frameEpoch) return null;
        delta(pages.map((index) => index + 1), false, 100, 3, pages);
        return new Uint8Array([frameEpoch]);
      },
      applyUpdate(_update: Uint8Array) {
        harness.releaseSuperseded = true;
      },
      setSelection() {},
      applyInput() {
        delta([1]);
        return new Uint8Array([frameEpoch]);
      },
      applyDelete() {
        delta([1]);
        return new Uint8Array([frameEpoch]);
      },
      residentDeletedUnits() {
        return 1;
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
      revisionCount(_excluding?: ReadonlySet<string>) {
        return 0;
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
      harness.buffers.push(buffer);
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
    release(canvas: Surface) {
      canvas.width = 1;
      canvas.height = 1;
      canvas.pixels = null;
    },
    presentCaret(canvas: Surface, buffer: Surface, _stage: Surface, caret: { color: string }) {
      if (!buffer.pixels) throw new Error('caret used a detached buffer');
      harness.presented.push(Number(buffer.pixels.split(':')[0]));
      canvas.pixels = `${buffer.pixels}|caret:${caret.color}`;
      canvas.width = buffer.width;
      canvas.height = buffer.height;
    },
  };
  startWorker(scope, Surface, harness, { now: () => harness.now() });
  function send(request: ResidentEngineWorkerRequestWithoutId) {
    const id = ++nextId;
    return new Promise<ResidentEngineWorkerResponse>((resolve) => {
      replies.set(id, resolve);
      scope.onmessage({
        data: { ...request, id } as ResidentEngineWorkerRequest,
      });
    });
  }
  function delta(
    upserts: number[],
    full = false,
    width = 100,
    pageCount = 3,
    unbuilt: number[] = []
  ) {
    const baseFrameEpoch = frameEpoch++;
    const operations: FramePageOperation[] = upserts.map((id) => ({
      kind: 'upsert',
      pageId: BigInt(id),
      pageIndex: id - 1,
      fingerprint: BigInt(frameEpoch),
      primitiveIds: new BigUint64Array(),
      page: {
        pageIndex: id - 1,
        width,
        height: 100,
        primitives: [],
        ...(unbuilt.includes(id - 1)
          ? { unbuilt: true, positionSpan: [id, id + 1] as [number, number] }
          : {}),
      },
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
    scope,
    harness,
    surfaces,
    answered,
    transfers,
    send,
    delta,
    resetCalls() {
      harness.rasterized = [];
      harness.presented = [];
    },
    async bootstrap(
      pageCount = 3,
      { keepSurfaces = false, heapLimitBytes }: { keepSurfaces?: boolean; heapLimitBytes?: number } = {}
    ) {
      delta(Array.from({ length: pageCount }, (_, index) => index + 1), true, 100, pageCount);
      return send({
        type: 'bootstrap',
        expectedFrameEpoch: 0,
        extras: '',
        ...(keepSurfaces ? { keepSurfaces } : {}),
        ...(heapLimitBytes !== undefined ? { heapLimitBytes } : {}),
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
    build(
      upserts: number[],
      width = 100,
      caret: YrsResidentCaretRect | null = null,
      displayWindow?: [number, number],
      retainBuiltPages?: boolean
    ) {
      delta(upserts, false, width);
      harness.caret = caret;
      return send({
        type: 'buildFrame',
        extras: '',
        expectedFrameEpoch: frameEpoch - 1,
        paintCaret: !!caret,
        displayWindow,
        ...(retainBuiltPages ? { retainBuiltPages } : {}),
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

test.each([false, true])(
  'a background page batch yields and returns ordered frames once with editModule=%s',
  async (hostModule) => {
    const w = worker();
    await w.bootstrap(9);
    w.harness.now = () => 0;
    const calls: number[][] = [];
    Object.assign(w.harness.session, {
      buildDisplayPagesFrame(pages: number[]) {
        calls.push(pages);
        const bytes = w.harness.session.applyInput();
        w.harness.delta = { ...w.harness.delta!, pageCount: 9 };
        if (hostModule && calls.length === 1) {
          w.scope.onmessage({ data: { type: 'editModule', module: null } });
        }
        return bytes;
      },
    });
    const response = await w.send({
      type: 'buildPages', pages: Array.from({ length: 9 }, (_, index) => index),
      expectedFrameEpoch: 1, paintCaret: false, background: true,
    });
    expect(response.ok).toBe(true);
    expect(calls).toEqual([[0, 1, 2, 3], [4, 5, 6, 7, 8]]);
    if (!response.ok) throw new Error(response.error);
    expect(response.pageFrames?.map((frame) => new Uint8Array(frame)[0])).toEqual([2, 3]);
    expect(response.caret?.frameEpoch).toBe(3);
    expect(w.answered).toEqual([1, 2]);
  }
);

test('a document read between background page slices leaves the complete batch intact', async () => {
  const w = worker();
  await w.bootstrap(9);
  w.harness.now = () => 0;
  const calls: number[][] = [];
  const order: string[] = [];
  let read!: Promise<ResidentEngineWorkerResponse>;
  Object.assign(w.harness.session, {
    proposalEngine: { version: () => 'current' },
    paragraphIdentities: () => ({ paragraphs: [] }),
    buildDisplayPagesFrame(pages: number[]) {
      calls.push(pages);
      const bytes = w.harness.session.applyInput();
      w.harness.delta = { ...w.harness.delta!, pageCount: 9 };
      if (calls.length === 1) {
        read = w.send({ type: 'documentRead', read: { kind: 'paragraphIdentities' } })
          .then((reply) => { order.push('read'); return reply; });
      }
      return bytes;
    },
  });
  const response = await w.send({
    type: 'buildPages', pages: Array.from({ length: 9 }, (_, index) => index),
    expectedFrameEpoch: 1, paintCaret: false, background: true,
  }).then((reply) => { order.push('build'); return reply; });
  expect(await read).toMatchObject({
    ok: true, read: { version: 'current', value: { paragraphs: [] } },
  });
  expect(response.ok).toBe(true);
  if (!response.ok) throw new Error(response.error);
  expect(response.pageBuildSuperseded).toBeUndefined();
  expect(calls).toEqual([[0, 1, 2, 3], [4, 5, 6, 7, 8]]);
  expect(response.pageFrames?.map((frame) => new Uint8Array(frame)[0])).toEqual([2, 3]);
  expect(response.caret?.frameEpoch).toBe(3);
  expect(order).toEqual(['read', 'build']);
  expect(w.answered).toEqual([1, 3, 2]);
});

test.each(['cancelled as stale', 'replaced while queued'])(
  'a background page build %s after a trap is answered with the trap',
  async (variant) => {
    const w = worker();
    await w.bootstrap(9);
    w.harness.now = () => 0;
    Object.assign(w.harness.session, {
      proposalEngine: { version: () => 'current' },
      paragraphIdentities: () => {
        throw new WebAssembly.RuntimeError('unreachable');
      },
    });
    const build = () => w.send({
      type: 'buildPages', pages: Array.from({ length: 9 }, (_, index) => index),
      expectedFrameEpoch: 1, paintCaret: false, background: true,
    });
    const read = () => w.send({ type: 'documentRead', read: { kind: 'paragraphIdentities' } });
    const replies = variant === 'cancelled as stale'
      ? [build(), read(), w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null })]
      : (() => {
          const trapped = read();
          const queued = build();
          return [trapped, queued, trapped.then(build)];
        })();
    for (const reply of await Promise.all(replies)) {
      expect(!reply.ok && reply.terminal).toBe(true);
    }
    expect(new Set(w.answered).size).toBe(w.answered.length);
  }
);

test('document reads skip stale versions and read the current version', async () => {
  const w = worker();
  await w.bootstrap();
  let reads = 0;
  Object.assign(w.harness.session, {
    proposalEngine: { version: () => 'current' },
    paragraphIdentities: () => {
      reads += 1;
      return { paragraphs: [] };
    },
  });
  const stale = await w.send({
    type: 'documentRead', expectVersion: 'stale', read: { kind: 'paragraphIdentities' },
  });
  expect(stale).toMatchObject({ ok: true, superseded: true });
  expect(stale).not.toHaveProperty('read');
  expect(reads).toBe(0);
  const current = await w.send({
    type: 'documentRead', expectVersion: 'current', read: { kind: 'paragraphIdentities' },
  });
  expect(current).toMatchObject({
    ok: true, read: { version: 'current', value: { paragraphs: [] } },
  });
  expect(current).not.toHaveProperty('superseded');
  expect(reads).toBe(1);
});

test.each([false, true])(
  'input overtakes only version-checked reads behind a running request with expectVersion=%s',
  async (checked) => {
    const w = worker();
    await w.bootstrap();
    let version = 'before';
    let reads = 0;
    const applyInput = w.harness.session.applyInput;
    Object.assign(w.harness.session, {
      proposalEngine: { version: () => version },
      paragraphIdentities: () => {
        reads += 1;
        return { paragraphs: [] };
      },
      applyInput: () => {
        version = 'after';
        return applyInput();
      },
    });
    const entered = deferred();
    const held = deferred();
    const rasterize = w.harness.rasterize;
    w.harness.rasterize = async (...args) => {
      entered.resolve();
      await held.promise;
      return rasterize(...args);
    };
    const attached = w.attach([1]);
    await entered.promise;
    const read = w.send({
      type: 'documentRead', read: { kind: 'paragraphIdentities' },
      ...(checked ? { expectVersion: 'before' } : {}),
    });
    w.harness.caret = caret(1);
    const loc = { story: 'body', paraId: 'p1', offset: 0 };
    const input = w.send({
      type: 'applyInput', text: 'x', selection: { anchor: loc, head: loc },
      expectedFrameEpoch: 1, profile: false, paintCaret: false,
    });
    expect(w.answered).toEqual([1]);
    expect(reads).toBe(0);
    held.resolve();
    const [surface, answer, edited] = await Promise.all([attached, read, input]);
    expect(surface.ok).toBe(true);
    expect(edited).toMatchObject({ ok: true, caret: { frameEpoch: 2 } });
    expect(edited).toHaveProperty('frame');
    if (checked) {
      expect(answer).toMatchObject({ ok: true, superseded: true });
      expect(answer).not.toHaveProperty('read');
      expect(reads).toBe(0);
      expect(w.answered).toEqual([1, 2, 4, 3]);
    } else {
      expect(answer).toMatchObject({
        ok: true, read: { version: 'before', value: { paragraphs: [] } },
      });
      expect(reads).toBe(1);
      expect(w.answered).toEqual([1, 2, 3, 4]);
    }
  }
);

test('fast background page slices grow beyond four pages and keep every frame', async () => {
  const w = worker();
  await w.bootstrap(50);
  w.harness.now = () => 0;
  const calls: number[][] = [];
  Object.assign(w.harness.session, {
    buildDisplayPagesFrame(pages: number[]) {
      calls.push(pages);
      const bytes = w.harness.session.applyInput();
      w.harness.delta = { ...w.harness.delta!, pageCount: 50 };
      return bytes;
    },
  });
  const pages = Array.from({ length: 50 }, (_, index) => index);
  const response = await w.send({
    type: 'buildPages', pages, expectedFrameEpoch: 1, paintCaret: false, background: true,
  });
  expect(response.ok).toBe(true);
  if (!response.ok) throw new Error(response.error);
  expect(calls.map((slice) => slice.length)).toEqual([4, 32, 14]);
  expect(calls.flat()).toEqual(pages);
  expect(response.pageFrames?.map((frame) => new Uint8Array(frame)[0])).toEqual([2, 3, 4]);
  expect(response.caret?.frameEpoch).toBe(4);
  expect(w.answered).toEqual([1, 2]);
});

test('slow background page slices keep at least four pages', async () => {
  const w = worker();
  await w.bootstrap(12);
  let clock = 0;
  w.harness.now = () => clock;
  const calls: number[][] = [];
  Object.assign(w.harness.session, {
    buildDisplayPagesFrame(pages: number[]) {
      calls.push(pages);
      clock += 100;
      const bytes = w.harness.session.applyInput();
      w.harness.delta = { ...w.harness.delta!, pageCount: 12 };
      return bytes;
    },
  });
  const pages = Array.from({ length: 12 }, (_, index) => index);
  const response = await w.send({
    type: 'buildPages', pages, expectedFrameEpoch: 1, paintCaret: false, background: true,
  });
  expect(response.ok).toBe(true);
  expect(calls.map((slice) => slice.length)).toEqual([4, 4, 4]);
  expect(calls.flat()).toEqual(pages);
});

test('a visible page request supersedes the remaining background slices', async () => {
  const w = worker();
  await w.bootstrap(9);
  const calls: number[][] = [];
  let visible!: Promise<ResidentEngineWorkerResponse>;
  Object.assign(w.harness.session, {
    buildDisplayPagesFrame(pages: number[]) {
      calls.push(pages);
      const bytes = w.harness.session.applyInput();
      w.harness.delta = { ...w.harness.delta!, pageCount: 9 };
      if (calls.length === 1) {
        visible = w.send({
          type: 'buildPages', pages: [8], expectedFrameEpoch: 2, paintCaret: false,
        });
      }
      return bytes;
    },
  });
  const background = await w.send({
    type: 'buildPages', pages: Array.from({ length: 9 }, (_, index) => index),
    expectedFrameEpoch: 1, paintCaret: false, background: true,
  });
  expect(background).toMatchObject({ ok: true, pageBuildSuperseded: true });
  expect((await visible).ok).toBe(true);
  expect(calls).toEqual([[0, 1, 2, 3], [8]]);
  expect(w.answered).toEqual([1, 2, 3]);
});

test('font suffixes preserve ids and glyph caches; mismatches require a full snapshot', async () => {
  const w = worker();
  expect(await w.bootstrap()).toMatchObject({ ok: true });
  expect(await w.attach([1])).toMatchObject({ ok: true });
  const font = new Uint8Array([1]);
  const sync = (
    fontsRevision: number,
    fonts: YrsResidentWorkerSnapshot['fonts'],
    fontsBaseRevision?: number
  ) => {
    return w.send({
      type: 'sync',
      extras: '',
      expectedFrameEpoch: w.harness.delta!.baseFrameEpoch,
      paintCaret: false,
      snapshot: {
        clientId: 1,
        state: new Uint8Array(),
        selection: null,
        fonts,
        fontsRevision,
        ...(fontsBaseRevision === undefined ? {} : { fontsBaseRevision }),
        renderInputs: [],
        measureInputs: [],
        layoutInput: '',
        layoutWithRegions: false,
        layoutRevision: 1,
      },
    });
  };
  w.delta([]);
  expect(await sync(1, [font], 0)).toMatchObject({ ok: true });
  w.delta([]);
  expect(await sync(2, [{ substituteOf: 0, family: 'Calibri' }], 1)).toMatchObject({ ok: true });
  w.delta([]);
  expect(await sync(2, [], 2)).toMatchObject({ ok: true });
  expect(w.harness.fontIds).toEqual([0, 1]);
  expect(w.harness.clearFontCalls).toBe(1);
  expect(w.harness.glyphCacheCreations).toBe(1);
  const beforeMismatch = {
    states: w.harness.loadedStates.length,
    media: w.harness.loadedMediaSources.length,
    partial: w.harness.partialDocuments.length,
    windows: w.harness.displayWindows.length,
  };
  expect(await sync(3, [font], 0)).toMatchObject({
    ok: false, error: 'Resident engine worker font base revision mismatch',
  });
  expect(await sync(2, [], 1)).toMatchObject({
    ok: false, error: 'Resident engine worker font base revision mismatch',
  });
  expect(w.harness.fontIds).toEqual([0, 1]);
  expect(w.harness.clearFontCalls).toBe(1);
  expect(w.harness.glyphCacheCreations).toBe(1);
  expect(w.harness.loadedStates).toHaveLength(beforeMismatch.states);
  expect(w.harness.loadedMediaSources).toHaveLength(beforeMismatch.media);
  expect(w.harness.partialDocuments).toHaveLength(beforeMismatch.partial);
  expect(w.harness.displayWindows).toHaveLength(beforeMismatch.windows);
  w.delta([]);
  expect(await sync(3, [font])).toMatchObject({ ok: true });
  expect(w.harness.fontIds).toEqual([0]);
  expect(w.harness.clearFontCalls).toBe(2);
  expect(w.harness.glyphCacheCreations).toBe(2);
});

test('full font sync repaints retained pages while suffix appends preserve their pixels', async () => {
  for (const paintCaret of [false, true]) {
    const w = worker();
    expect(await w.bootstrap()).toMatchObject({ ok: true });
    w.harness.caret = paintCaret ? caret(1) : null;
    const sync = (
      fontsRevision: number,
      fonts: YrsResidentWorkerSnapshot['fonts'],
      fontsBaseRevision?: number
    ) => {
      w.delta([]);
      return w.send({
        type: 'sync',
        extras: '',
        expectedFrameEpoch: w.harness.delta!.baseFrameEpoch,
        paintCaret,
        snapshot: {
          clientId: 1,
          state: new Uint8Array(),
          selection: null,
          fonts,
          fontsRevision,
          ...(fontsBaseRevision === undefined ? {} : { fontsBaseRevision }),
          renderInputs: [],
          measureInputs: [],
          layoutInput: '',
          layoutWithRegions: false,
          layoutRevision: 1,
        },
      });
    };
    expect(await sync(1, [new Uint8Array([1])])).toMatchObject({ ok: true });
    expect(await w.attach([1, 2])).toMatchObject({ ok: true });
    expect(w.harness.rasterized).toEqual([1, 2]);
    w.resetCalls();
    expect(await sync(2, [new Uint8Array([2])])).toMatchObject({
      ok: true,
      replayedPages: 2,
      caretPainted: paintCaret,
    });
    expect(w.harness.delta!.operations).toEqual([]);
    expect(w.harness.rasterized).toEqual([1, 2]);
    expect(w.harness.presented).toEqual([1, 2]);
    expect(w.surfaces.get('1')!.pixels).toBe(paintCaret ? '1:100|caret:#000' : '1:100');
    expect(w.surfaces.get('2')!.pixels).toBe('2:100');
    w.resetCalls();
    expect(await w.build([], 100, w.harness.caret)).toMatchObject({
      ok: true,
      replayedPages: 0,
      caretPainted: paintCaret,
    });
    expect(await sync(3, [new Uint8Array([3])], 2)).toMatchObject({
      ok: true,
      replayedPages: 0,
      caretPainted: paintCaret,
    });
    expect(w.harness.rasterized).toEqual([]);
    expect(w.harness.presented).toEqual([]);
  }
});

describe('resident display page release', () => {
  test('answers build, release, then input in FIFO order', async () => {
    const w = worker();
    await w.bootstrap();
    w.harness.caret = caret(1);
    const order: string[] = [];
    const build = w
      .send({ type: 'buildPages', pages: [2], expectedFrameEpoch: 1, paintCaret: false })
      .then((reply) => {
        order.push('build');
        return reply;
      });
    const release = w
      .send({
        type: 'releasePages',
        pages: [{ index: 2, pageId: '3' }],
        expectedFrameEpoch: 2,
        paintCaret: false,
      })
      .then((reply) => {
        order.push('release');
        return reply;
      });
    const loc = { story: 'body', paraId: 'p1', offset: 0 };
    const input = w
      .send({
        type: 'applyInput',
        text: 'x',
        selection: { anchor: loc, head: loc },
        expectedFrameEpoch: 3,
        profile: false,
        paintCaret: false,
      })
      .then((reply) => {
        order.push('input');
        return reply;
      });
    const replies = await Promise.all([build, release, input]);
    expect(order).toEqual(['build', 'release', 'input']);
    expect(w.answered).toEqual([1, 2, 3, 4]);
    for (const [index, reply] of replies.entries()) {
      expect(reply).toMatchObject({ ok: true, caret: { frameEpoch: index + 2 }, selection: null });
      expect(reply).toHaveProperty('frame');
    }
    expect(w.harness.releaseCalls).toEqual([{ pages: [2], expectedFrameEpoch: 2 }]);
  });

  test('stale epochs and mismatched page identities supersede without calling the engine', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([3]);
    w.resetCalls();
    const before = w.harness.delta;
    for (const request of [
      { pages: [{ index: 2, pageId: '3' }], expectedFrameEpoch: 0 },
      { pages: [{ index: 2, pageId: '2' }], expectedFrameEpoch: 1 },
      { pages: [{ index: 8, pageId: '3' }], expectedFrameEpoch: 1 },
    ]) {
      const reply = await w.send({ type: 'releasePages', ...request, paintCaret: false });
      expect(reply).toMatchObject({ ok: true, superseded: true });
      expect(reply).not.toHaveProperty('frame');
    }
    expect(w.harness.releaseCalls).toEqual([]);
    expect(w.harness.delta).toBe(before);
    expect(w.harness.rasterized).toEqual([]);
    expect(w.surfaces.get('3')?.pixels).toBe('3:100');
    expect(
      await w.send({
        type: 'releasePages',
        pages: [{ index: 2, pageId: '3' }],
        expectedFrameEpoch: 1,
        paintCaret: false,
      })
    ).toHaveProperty('frame');
  });

  test('an engine superseded release after an update emits no frame and preserves surfaces', async () => {
    const w = worker();
    await w.bootstrap();
    await w.attach([3]);
    w.resetCalls();
    const before = w.harness.delta;
    void w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null });
    const reply = await w.send({
      type: 'releasePages',
      pages: [{ index: 2, pageId: '3' }],
      expectedFrameEpoch: 1,
      paintCaret: false,
    });
    expect(reply).toMatchObject({ ok: true, superseded: true });
    expect(reply).not.toHaveProperty('frame');
    expect(w.harness.delta).toBe(before);
    expect(w.harness.rasterized).toEqual([]);
    expect(w.surfaces.get('3')?.pixels).toBe('3:100');
  });

  test('a released page keeps its transferred canvas without pixels and paints into it once rebuilt', async () => {
    const w = worker();
    await w.bootstrap();
    await w.build([], 100, caret(3));
    await w.attach([3]);
    const canvas = w.surfaces.get('3')!;
    expect(canvas.pixels).toStartWith('3:100');
    const oldBuffer = w.harness.buffers.at(-1)!;
    w.resetCalls();
    const reply = await w.send({
      type: 'releasePages',
      pages: [{ index: 2, pageId: '3' }],
      expectedFrameEpoch: 2,
      paintCaret: false,
    });
    expect(reply).toMatchObject({ ok: true, replayedPages: 0, caretPainted: false });
    expect(w.harness.rasterized).toEqual([]);
    expect([canvas.pixels, canvas.width, canvas.height]).toEqual([null, 1, 1]);
    await w.attach([3]);
    expect(w.harness.rasterized).toEqual([]);
    expect(canvas.pixels).toBeNull();
    await w.send({ type: 'buildPages', pages: [2], expectedFrameEpoch: 3, paintCaret: false });
    expect(w.harness.rasterized).toEqual([3]);
    expect(w.surfaces.get('3')).toBe(canvas);
    expect(canvas.pixels).toBe('3:100');
    expect(w.harness.buffers.at(-1)).not.toBe(oldBuffer);
  });

  test('a canvas attached while its page is unbuilt is painted once the page builds', async () => {
    const w = worker();
    await w.bootstrap();
    await w.send({
      type: 'releasePages',
      pages: [{ index: 2, pageId: '3' }],
      expectedFrameEpoch: 1,
      paintCaret: false,
    });
    w.resetCalls();
    await w.attach([3]);
    const canvas = w.surfaces.get('3')!;
    expect(w.harness.rasterized).toEqual([]);
    expect(canvas.pixels).toBeNull();
    await w.send({ type: 'buildPages', pages: [2], expectedFrameEpoch: 2, paintCaret: false });
    expect(w.harness.rasterized).toEqual([3]);
    expect(canvas.pixels).toBe('3:100');
  });
});

describe('resident worker warmup', () => {
  const timers = new Map<number, { callback: () => void; ms: number }>();
  const realSetTimeout = globalThis.setTimeout;
  const realClearTimeout = globalThis.clearTimeout;
  let nextTimer = 1;

  beforeEach(() => {
    timers.clear();
    globalThis.setTimeout = ((callback: () => void, ms: number) => {
      const id = nextTimer++;
      timers.set(id, { callback, ms });
      return id;
    }) as unknown as typeof setTimeout;
    globalThis.clearTimeout = ((id: number) => {
      timers.delete(id);
    }) as unknown as typeof clearTimeout;
  });

  afterEach(() => {
    globalThis.setTimeout = realSetTimeout;
    globalThis.clearTimeout = realClearTimeout;
  });

  function advanceTimers(ms: number): void {
    for (const [id, timer] of [...timers]) {
      timer.ms -= ms;
      if (timer.ms > 0) continue;
      timers.delete(id);
      timer.callback();
    }
  }

  function armedBudgets(): number[] {
    return [...timers.values()].map((timer) => timer.ms);
  }

  function hostWarm(w: ReturnType<typeof worker>) {
    const started = deferred();
    const preloadFrom = w.harness.preloadFrom;
    w.harness.preloadFrom = (source) => {
      started.resolve();
      return preloadFrom(source);
    };
    return { reply: w.send({ type: 'warm', hostModule: true }), started: started.promise };
  }

  test('waits for the host module outside the request queue and does not answer it', async () => {
    const w = worker();
    const module = new WebAssembly.Module(
      new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
    );
    const { reply: warm, started } = hostWarm(w);
    const bootstrap = w.bootstrap();
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    advanceTimers(RESIDENT_HOST_MODULE_WAIT_MS - 1);
    await Promise.resolve();
    expect(w.harness.preloadInputs).toEqual([]);
    expect(w.harness.initializations).toBe(0);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([]);
    w.scope.onmessage({ data: { type: 'editModule', module } });
    expect((await warm).ok).toBe(true);
    expect((await bootstrap).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([module, undefined]);
    expect(w.harness.initializations).toBe(1);
    expect(w.answered).toEqual([1, 2]);
    expect(armedBudgets()).toEqual([]);
    advanceTimers(1);
    w.scope.onmessage({ data: { type: 'editModule', module: null } });
    expect((await w.send({ type: 'warm', hostModule: true })).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([module, undefined, module]);
    expect(w.answered).toEqual([1, 2, 3]);
  });

  test('waits for a fresh host module after a failed warm', async () => {
    const w = worker();
    const bytes = new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00]);
    const moduleA = new WebAssembly.Module(bytes);
    const moduleB = new WebAssembly.Module(bytes);
    w.harness.failWarm = new Error('init failed');
    const { reply: warm, started } = hostWarm(w);
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    const firstTimers = [...timers.keys()];
    w.scope.onmessage({ data: { type: 'editModule', module: moduleA } });
    const failed = await warm;
    expect(failed).toMatchObject({ id: 1, ok: false, error: 'init failed' });
    expect(failed).not.toHaveProperty('terminal');
    expect(w.harness.preloadInputs).toEqual([moduleA]);
    expect(w.harness.initializations).toBe(0);
    expect(w.answered).toEqual([1]);
    expect(armedBudgets()).toEqual([]);

    const { reply: retry, started: retryStarted } = hostWarm(w);
    await retryStarted;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    expect([...timers.keys()]).not.toEqual(firstTimers);
    advanceTimers(RESIDENT_HOST_MODULE_WAIT_MS - 1);
    await Promise.resolve();
    expect(w.harness.preloadInputs).toEqual([moduleA]);
    expect(w.harness.initializations).toBe(0);
    expect(w.answered).toEqual([1]);

    w.scope.onmessage({ data: { type: 'editModule', module: moduleB } });
    expect(await retry).toMatchObject({ id: 2, ok: true });
    expect(w.harness.preloadInputs).toEqual([moduleA, moduleB]);
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([1, 2]);
    expect(armedBudgets()).toEqual([]);
  });

  test('falls back to the asset preload when the host sends null', async () => {
    const w = worker();
    const { reply: warm, started } = hostWarm(w);
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    w.scope.onmessage({ data: { type: 'editModule', module: null } });
    expect((await warm).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([undefined]);
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([1]);
    expect(armedBudgets()).toEqual([]);
  });

  test('falls back at once when the host sends a value that is not a module', async () => {
    const w = worker();
    const { reply: warm, started } = hostWarm(w);
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    expect(w.harness.preloadInputs).toEqual([]);
    expect(w.answered).toEqual([]);
    w.scope.onmessage({ data: { type: 'editModule', module: {} as WebAssembly.Module } });
    expect((await warm).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([undefined]);
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([1]);
    expect(armedBudgets()).toEqual([]);
  });

  test('falls back at once when the host module message cannot be received', async () => {
    const w = worker();
    const { reply: warm, started } = hostWarm(w);
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    w.scope.onmessageerror?.();
    expect((await warm).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([undefined]);
    expect(w.answered).toEqual([1]);
    expect(armedBudgets()).toEqual([]);
  });

  test('falls back after the host module wait expires and ignores late modules', async () => {
    const w = worker();
    const module = new WebAssembly.Module(
      new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
    );
    const { reply: warm, started } = hostWarm(w);
    await started;
    expect(armedBudgets()).toEqual([RESIDENT_HOST_MODULE_WAIT_MS]);
    advanceTimers(RESIDENT_HOST_MODULE_WAIT_MS - 1);
    await Promise.resolve();
    expect(w.harness.preloadInputs).toEqual([]);
    expect(w.harness.initializations).toBe(0);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([]);
    advanceTimers(1);
    w.scope.onmessage({ data: { type: 'editModule', module } });
    expect((await warm).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([undefined]);
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([1]);
    expect(armedBudgets()).toEqual([]);
    w.scope.onmessage({ data: { type: 'editModule', module } });
    expect((await w.send({ type: 'warm', hostModule: true })).ok).toBe(true);
    expect(w.harness.preloadInputs).toEqual([undefined, undefined]);
    expect(w.answered).toEqual([1, 2]);
  });

  test('initializes wasm without creating a session, then bootstraps a frame', async () => {
    const w = worker();
    expect(await w.send({ type: 'warm' })).toMatchObject({ ok: true });
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.harness.preloadInputs).toEqual([undefined]);
    expect(await w.build([])).toMatchObject({
      ok: false,
      error: 'Resident engine worker is not initialized',
    });
    expect(await w.bootstrap()).toMatchObject({ ok: true, layoutRevision: 1 });
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(1);
    expect((await w.attach([1])).ok).toBe(true);
    expect(w.harness.rasterized).toEqual([1]);
  });

  test('retries initialization after a failed warm, including a wasm runtime error', async () => {
    const w = worker();
    w.harness.failWarm = new WebAssembly.RuntimeError('init failed');
    const failed = await w.send({ type: 'warm' });
    expect(failed).toMatchObject({ ok: false, error: 'init failed' });
    expect(failed).not.toHaveProperty('terminal');
    expect(w.harness.sessionsCreated).toBe(0);
    expect((await w.bootstrap()).ok).toBe(true);
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(1);
  });

  test('queues bootstrap behind an in-flight warm', async () => {
    const w = worker();
    const loading = deferred();
    w.harness.preloadBlock = loading.promise;
    const started = deferred();
    const preload = w.harness.preload;
    w.harness.preload = async (...args) => {
      started.resolve();
      return preload(...args);
    };
    const warm = w.send({ type: 'warm' });
    const bootstrap = w.bootstrap();
    await started.promise;
    expect(w.harness.sessionsCreated).toBe(0);
    expect(w.answered).toEqual([]);
    loading.resolve();
    expect((await warm).ok).toBe(true);
    expect((await bootstrap).ok).toBe(true);
    expect(w.answered).toEqual([1, 2]);
    expect(w.harness.initializations).toBe(1);
  });
});

describe('resident worker page damage', () => {
  test('frame requests opt into windowed incremental builds only with a display window', async () => {
    const w = worker();
    expect((await w.bootstrap()).ok).toBe(true);
    expect(w.harness.windowedIncrementalBuilds).toEqual([false]);
    expect((await w.build([], 100, null, [8, 11])).ok).toBe(true);
    expect(w.harness.displayWindows).toEqual([[8, 11]]);
    expect(w.harness.retainBuiltPages).toEqual([false]);
    expect(w.harness.windowedIncrementalBuilds).toEqual([false, true]);
    expect((await w.build([])).ok).toBe(true);
    expect(w.harness.displayWindows).toEqual([[8, 11]]);
    expect(w.harness.retainBuiltPages).toEqual([false]);
    expect(w.harness.windowedIncrementalBuilds).toEqual([false, true, false]);
    expect((await w.build([], 100, null, [8, 11], true)).ok).toBe(true);
    expect(w.harness.retainBuiltPages).toEqual([false, true]);
    expect((await w.build([], 100, null, [8, 11])).ok).toBe(true);
    expect(w.harness.retainBuiltPages).toEqual([false, true, false]);
  });

  test('input and delete requests without a display window disable a previous opt-in', async () => {
    const w = worker();
    await w.bootstrap();
    const loc = { story: 'body', paraId: 'p1', offset: 1 };
    const options = {
      selection: { anchor: loc, head: loc },
      profile: false,
      paintCaret: false,
    };
    expect((await w.build([], 100, caret(1), [8, 11])).ok).toBe(true);
    expect(
      (
        await w.send({
          type: 'applyInput',
          text: 'x',
          expectedFrameEpoch: 2,
          ...options,
        })
      ).ok
    ).toBe(true);
    expect((await w.build([], 100, caret(1), [8, 11])).ok).toBe(true);
    expect(
      (
        await w.send({
          type: 'applyDelete',
          direction: 'backward',
          count: 1,
          expectedFrameEpoch: 4,
          ...options,
        })
      ).ok
    ).toBe(true);
    expect(w.harness.windowedIncrementalBuilds).toEqual([false, true, false, true, false]);
  });

  test('a bootstrap that keeps surfaces paints the next document into the attached canvases', async () => {
    const w = worker();
    await w.bootstrap(3);
    await w.attach([1, 2]);
    w.resetCalls();
    expect((await w.bootstrap(3, { keepSurfaces: true })).ok).toBe(true);
    expect(w.harness.presented).toEqual([1, 2]);
    w.resetCalls();
    await w.bootstrap(3);
    expect(w.harness.presented).toEqual([]);
  });

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
    expect(w.surfaces.get('1')!.pixels).toBeNull();
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
    expect((await w.send({ type: 'warm' })).ok).toBe(true);
    expect(w.harness.sessionsCreated).toBe(0);
    const extras: string[] = [];
    const layoutJson = JSON.stringify({
      layout: { pages: [] },
      headersFooters: { parts: [] },
      notesConverged: true,
    });
    let epoch = 0;
    const layouts: string[] = [];
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsRetainedJson: () => {
        layouts.push('reply');
        return layoutJson;
      },
      layoutDocumentWithRegionsRetained: () => layouts.push('retained'),
      retainedHeadersFootersJson: () => JSON.stringify({ parts: [] }),
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
    expect(w.harness.initializations).toBe(1);
    expect(w.harness.sessionsCreated).toBe(1);
    expect(extras).toEqual([
      '{"headersFooters":{"parts":[]},"fontChains":{"a|0|0":[1]},"resolvedCommentIds":[4]}',
    ]);
    expect(w.harness.directBatches).toEqual([]);

    const plain = await w.send({
      type: 'sync',
      expectedFrameEpoch: 0,
      extras: 'given',
      paintCaret: false,
      snapshot,
    });
    expect(plain.ok && plain.layoutJson).toBeUndefined();
    expect(extras.at(-1)).toBe('given');
    expect(layouts).toEqual(['reply', 'retained']);
  });

  test('marks a replica as a preview or not before it lays it out', async () => {
    const w = worker();
    const calls: string[] = [];
    Object.assign(w.harness.session, {
      loadState: () => calls.push('load'),
      setPartialDocument: (partial: boolean) => calls.push(`partial:${partial}`),
      layoutDocumentWithRegionsRetained: () => calls.push('layout'),
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

  test('builds only the pages a cut preview lays out like the whole document', async () => {
    const w = worker();
    let epoch = 0;
    const built: number[][] = [];
    const frame = (full: boolean) => {
      epoch += 1;
      w.harness.delta = {
        protocolVersion: 1,
        full,
        frameEpoch: epoch,
        baseFrameEpoch: full ? 0 : epoch - 1,
        docEpoch: epoch,
        layoutEpoch: epoch,
        pageCount: 0,
        operations: [],
        bytes: new Uint8Array(),
      };
      return new Uint8Array([0]);
    };
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsPrefixRetainedJson: () =>
        '{"layout":{"pages":[1,2,3]},"notesConverged":true}',
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: () => frame(true),
      buildDisplayPagesFrame: (pages: number[]) => {
        built.push(pages);
        return frame(false);
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
      partialDocument: true,
    };
    // The seed ran out on its third page, so only its first is final.
    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '',
      snapshot,
      layoutExtras: '{}',
      provisionalPages: 3,
      displayWindow: [0, 2],
    });
    expect(w.harness.displayWindows.at(-1)).toEqual([0, 1]);
    await w.send({
      type: 'buildFrame',
      extras: '',
      expectedFrameEpoch: epoch,
      paintCaret: false,
      displayWindow: [0, 3],
    });
    expect(w.harness.displayWindows.at(-1)).toEqual([0, 1]);
    await w.send({ type: 'buildPages', pages: [0, 1, 2], expectedFrameEpoch: epoch, paintCaret: false });
    expect(built).toEqual([[0]]);
    await w.send({
      type: 'sync',
      expectedFrameEpoch: epoch,
      extras: '',
      paintCaret: false,
      snapshot: { ...snapshot, partialDocument: false },
      layoutExtras: '{}',
      provisionalPages: 3,
      displayWindow: [0, 3],
    });
    expect(w.harness.displayWindows.at(-1)).toEqual([0, 3]);
    await w.send({ type: 'buildPages', pages: [0, 1, 2], expectedFrameEpoch: epoch, paintCaret: false });
    expect(built).toEqual([[0], [0, 1, 2]]);
  });

  test('lays out the media sources a snapshot carries, and clears them when it carries none', async () => {
    const w = worker();
    const loaded: string[] = [];
    Object.assign(w.harness.session, {
      loadMediaSources: (json: string) => loaded.push(json),
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout: { pages: [] }, notesConverged: true }),
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
    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '{}',
      snapshot: { ...snapshot, mediaSources: '{"sources":1}' },
    });
    await w.send({ type: 'sync', expectedFrameEpoch: 0, extras: '{}', paintCaret: false, snapshot });
    expect(loaded).toEqual(['{"sources":1}', '']);
  });

  test('lays out the note separators a snapshot carries, clears absent ones, and preserves worker-owned ones', async () => {
    const w = worker();
    const loaded: Uint8Array[] = [];
    Object.assign(w.harness.session, {
      loadNoteSeparators: (state: Uint8Array) => loaded.push(state),
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout: { pages: [] }, notesConverged: true }),
    });
    const snapshot: YrsResidentWorkerSnapshot = {
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
    const noteSeparators = new Uint8Array([1, 2, 3]);
    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '{}',
      snapshot: { ...snapshot, noteSeparators },
    });
    expect(loaded[0]).toBe(noteSeparators);
    await w.send({ type: 'sync', expectedFrameEpoch: 0, extras: '{}', paintCaret: false, snapshot });
    expect(loaded).toEqual([noteSeparators, new Uint8Array(0)]);
    for (const workerSnapshot of [{ ...snapshot, noteSeparators }, snapshot]) {
      await w.send({
        type: 'sync',
        expectedFrameEpoch: 0,
        extras: '{}',
        paintCaret: false,
        snapshot: { ...workerSnapshot, workerAuthoritative: true },
      });
    }
    expect(loaded).toHaveLength(2);
  });

  test('finishes a provisional layout on request and before other work', async () => {
    const w = worker();
    const extras: string[] = [];
    const calls: string[] = [];
    let epoch = 0;
    const provisional = '{"layout":{"pages":[1]},"notesConverged":true,"provisional":true}';
    const full = '{"layout":{"pages":[1,2]},"notesConverged":true}';
    let headersFooters = '{"parts":["full"]}';
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsPrefixRetainedJson: (_input: string, pages: number) => {
        calls.push(`prefix:${pages}`);
        return provisional;
      },
      layoutDocumentWithRegionsRetainedJson: () => {
        calls.push('full');
        return full;
      },
      retainedHeadersFootersJson: () => headersFooters,
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
    // A later edit changes what the session retains; the completed layout's reply keeps its own.
    headersFooters = '{"parts":["edited"]}';
    const late = await w.send({ type: 'completeLayout', expectedFrameEpoch: 4, paintCaret: false });
    expect(late.ok && late.layoutJson).toBe(full);
    expect(extras.at(-1)).toBe('{"headersFooters":{"parts":["full"]}}');
    expect(calls).toHaveLength(4);

    await w.send({
      type: 'bootstrap',
      expectedFrameEpoch: 0,
      extras: '',
      snapshot,
      layoutExtras: '{}',
      provisionalPages: 3,
    });
    Object.assign(w.harness.session, {
      setSelection: () => {},
      applyInput: (_text: string, expected: number) => {
        calls.push('input');
        const session = w.harness.session as unknown as {
          buildDisplayListFrame: (input: string, expected: number) => Uint8Array;
        };
        return session.buildDisplayListFrame('input', expected);
      },
    });
    // A header caret needs no body caret geometry from the stub.
    const loc = { story: 'header1', paraId: '1', offset: 0 };
    const selection = { anchor: loc, head: loc };
    await w.send({
      type: 'applyInput',
      text: 'x',
      selection,
      expectedFrameEpoch: epoch,
      profile: false,
      paintCaret: false,
    });
    expect(calls.slice(4)).toEqual(['prefix:3', 'full', 'input']);
    // The edit re-paginated: a completion built now would pair its cached headers with the new pages.
    const framesBefore = extras.length;
    const afterEdit = await w.send({ type: 'completeLayout', expectedFrameEpoch: epoch, paintCaret: false });
    expect(afterEdit.ok && afterEdit.frame).toBeUndefined();
    expect(afterEdit.ok && afterEdit.layoutJson).toBeUndefined();
    expect(extras).toHaveLength(framesBefore);
  });
});

describe('resident worker whole-document provisional pages', () => {
  const snapshot: YrsResidentWorkerSnapshot = {
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
  const provisional = JSON.stringify({ layout: { pages: [0, 1, 2, 3, 4, 5] }, provisional: true });
  const full = JSON.stringify({ layout: { pages: [0, 1, 2, 3, 4, 5, 6] } });

  function provisionalWorker(prefixComplete = false) {
    const w = worker();
    let epoch = 0;
    const built: number[][] = [];
    const calls: string[] = [];
    const frameWindows: Array<[number, number] | undefined> = [];
    const frame = (isFull: boolean) => {
      epoch += 1;
      w.harness.delta = {
        protocolVersion: 1,
        full: isFull,
        frameEpoch: epoch,
        baseFrameEpoch: isFull ? 0 : epoch - 1,
        docEpoch: epoch,
        layoutEpoch: epoch,
        pageCount: 0,
        operations: [],
        bytes: new Uint8Array(),
      };
      return new Uint8Array([0]);
    };
    Object.assign(w.harness.session, {
      layoutDocumentWithRegionsPrefixRetainedJson: (_input: string, pages: number) => {
        calls.push(`prefix:${pages}`);
        return prefixComplete ? full : provisional;
      },
      layoutDocumentWithRegionsRetainedJson: () => {
        calls.push('full');
        return full;
      },
      beginRegionLayout: () => ({ measuredBlocks: 0, bodyBlocks: 10 }),
      resumeRegionLayout: () => ({ measuredBlocks: 10, bodyBlocks: 10, layoutJson: full }),
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: () => {
        frameWindows.push(w.harness.displayWindows.at(-1));
        return frame(true);
      },
      buildDisplayPagesFrame: (pages: number[]) => {
        built.push(pages);
        return frame(false);
      },
      applyInput: () => {
        calls.push('input');
        frameWindows.push(w.harness.displayWindows.at(-1));
        return frame(true);
      },
    });
    return { w, built, calls, frameWindows, epoch: () => epoch };
  }

  test('bootstrap builds only exact prefix pages and restores the window on completion', async () => {
    const { w, built, calls, frameWindows, epoch } = provisionalWorker();
    const bootstrap = await w.send({
      type: 'bootstrap', snapshot, provisionalPages: 3, displayWindow: [0, 6],
      retainBuiltPages: true, extras: '', layoutExtras: '{}', expectedFrameEpoch: 0,
    });
    expect(bootstrap.ok && bootstrap.layoutProvisional).toBe(true);
    expect(frameWindows).toEqual([[0, 3]]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(false);
    for (const background of [false, true]) {
      const reply = await w.send({
        type: 'buildPages', pages: [2, 3, 4, 5], background,
        expectedFrameEpoch: epoch(), paintCaret: false,
      });
      expect(reply.ok).toBe(true);
    }
    expect(built).toEqual([[2], [2]]);
    const completed = await w.send({
      type: 'completeLayout', expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(frameWindows).toEqual([[0, 3], [0, 6]]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(true);
    await w.send({ type: 'buildPages', pages: [4], expectedFrameEpoch: epoch(), paintCaret: false });
    expect(built.at(-1)).toEqual([4]);
    expect(calls).toEqual(['prefix:3', 'full']);
  });

  test('worker-authoritative sync limits the prefix until sliced completion', async () => {
    const { w, built, calls, frameWindows, epoch } = provisionalWorker();
    await w.bootstrap();
    const synced = await w.send({
      type: 'sync', snapshot: { ...snapshot, workerAuthoritative: true },
      provisionalPages: 3, displayWindow: [0, 6], retainBuiltPages: true,
      extras: '', layoutExtras: '{}', expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(synced.ok && synced.layoutProvisional).toBe(true);
    expect(frameWindows.at(-1)).toEqual([0, 3]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(false);
    await w.send({
      type: 'buildPages', pages: [2, 3, 4, 5], expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(built).toEqual([[2]]);
    const completed = await w.send({
      type: 'completeLayout', expectedFrameEpoch: epoch(), paintCaret: false, sliceBlocks: 8,
    });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(frameWindows.at(-1)).toEqual([0, 6]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(true);
    await w.send({ type: 'buildPages', pages: [4], expectedFrameEpoch: epoch(), paintCaret: false });
    expect(built.at(-1)).toEqual([4]);
    expect(calls).toEqual(['prefix:3']);
  });

  test('a complete whole-document prefix is not clamped', async () => {
    const { w, built, frameWindows, epoch } = provisionalWorker(true);
    const bootstrap = await w.send({
      type: 'bootstrap', snapshot, provisionalPages: 3, displayWindow: [0, 6],
      retainBuiltPages: true, extras: '', layoutExtras: '{}', expectedFrameEpoch: 0,
    });
    expect(bootstrap.ok && bootstrap.layoutProvisional).toBeUndefined();
    expect(frameWindows).toEqual([[0, 6]]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(true);
    await w.send({
      type: 'buildPages', pages: [2, 3, 4, 5], expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(built).toEqual([[2, 3, 4, 5]]);
  });

  test('background slices build only the requested pages below the provisional limit', async () => {
    const { w, built, epoch } = provisionalWorker();
    await w.send({
      type: 'bootstrap', snapshot, provisionalPages: 6, displayWindow: [0, 10],
      extras: '', expectedFrameEpoch: 0,
    });
    const reply = await w.send({
      type: 'buildPages', pages: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9], background: true,
      expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(reply.ok).toBe(true);
    expect(built).toEqual([[0, 1, 2, 3], [4, 5]]);
  });

  test('worker-authoritative buildFrame keeps the limit and preserves an omitted window', async () => {
    const { w, built, calls, frameWindows, epoch } = provisionalWorker();
    await w.send({
      type: 'bootstrap', snapshot: { ...snapshot, workerAuthoritative: true },
      provisionalPages: 3, displayWindow: [0, 6], extras: '', expectedFrameEpoch: 0,
    });
    await w.send({
      type: 'buildFrame', displayWindow: [1, 6], retainBuiltPages: true,
      extras: '', expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(frameWindows.at(-1)).toEqual([1, 3]);
    expect(w.harness.retainBuiltPages.at(-1)).toBe(false);
    await w.send({ type: 'buildFrame', extras: '', expectedFrameEpoch: epoch(), paintCaret: false });
    expect(frameWindows.at(-1)).toEqual([1, 3]);
    expect(w.harness.windowedIncrementalBuilds.at(-1)).toBe(false);
    expect(calls).toEqual(['prefix:3']);
    await w.send({
      type: 'completeLayout', expectedFrameEpoch: epoch(), paintCaret: false,
    });
    expect(frameWindows.at(-1)).toEqual([1, 6]);
    await w.send({ type: 'buildPages', pages: [4], expectedFrameEpoch: epoch(), paintCaret: false });
    expect(built.at(-1)).toEqual([4]);
  });

  test('sync and bootstrap replace the provisional limit with a complete layout', async () => {
    const { w, built, frameWindows, epoch } = provisionalWorker();
    for (const type of ['sync', 'bootstrap'] as const) {
      await w.send({
        type: 'bootstrap', snapshot, provisionalPages: 3, displayWindow: [0, 6],
        extras: '', expectedFrameEpoch: 0,
      });
      const reply = await w.send({
        type, snapshot, displayWindow: [0, 6], extras: '',
        expectedFrameEpoch: epoch(), paintCaret: false,
      });
      expect(reply.ok).toBe(true);
      expect(frameWindows.at(-1)).toEqual([0, 6]);
      await w.send({ type: 'buildPages', pages: [4], expectedFrameEpoch: epoch(), paintCaret: false });
      expect(built.at(-1)).toEqual([4]);
    }
  });

  test('input completes the layout and lifts the limit before building its frame', async () => {
    const { w, built, calls, frameWindows, epoch } = provisionalWorker();
    await w.send({
      type: 'bootstrap', snapshot: { ...snapshot, workerAuthoritative: true },
      provisionalPages: 3, displayWindow: [0, 6], extras: '', expectedFrameEpoch: 0,
    });
    const loc = { story: 'header1', paraId: '1', offset: 0 };
    const reply = await w.send({
      type: 'applyInput', text: 'x', selection: { anchor: loc, head: loc },
      expectedFrameEpoch: epoch(), profile: false, paintCaret: false,
    });
    expect(reply.ok).toBe(true);
    expect(calls).toEqual(['prefix:3', 'full', 'input']);
    expect(frameWindows.at(-1)).toEqual([0, 6]);
    await w.send({ type: 'buildPages', pages: [4], expectedFrameEpoch: epoch(), paintCaret: false });
    expect(built.at(-1)).toEqual([4]);
  });
});

describe('resident worker revision counts', () => {
  test('revisionCount excludes revisions created by worker proposals', async () => {
    const w = worker();
    let version = 'proposal-1';
    const revisions = ['document-1', 'document-2'];
    const excluded: ReadonlySet<string>[] = [];
    const engine = {
      version: () => version,
      resolveParagraphAnchor: (anchor: DocxParagraphAnchor) => ({ status: 'found', anchor }),
      applyEdits: (request: DocxEditRequest) => {
        version = 'proposal-2';
        return {
          ok: true,
          version,
          source: 'host',
          changedStories: ['body'],
          receipts: request.steps.map((_step, stepIndex) => {
            const revisionId = `worker-${stepIndex + 1}`;
            revisions.push(revisionId);
            return {
              stepIndex, changed: true, revisionIds: [revisionId],
              newParagraphs: [], removedParagraphs: [],
            };
          }),
        };
      },
    };
    Object.assign(w.harness.session, {
      proposalEngine: engine,
      geometryReader: {
        version: engine.version,
        listRevisions: () => [],
        resolveParagraphAnchor: () => ({ status: 'missing' }),
        hasStory: () => false,
      },
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
    });
    w.harness.session.revisionCount = (excluding) => {
      expect(excluding).toBeInstanceOf(Set);
      excluded.push(excluding!);
      return revisions.filter((id) => !excluding!.has(id)).length;
    };
    expect((await w.bootstrap()).ok).toBe(true);
    const proposed = await w.send({
      type: 'proposal',
      operation: {
        kind: 'propose',
        request: {
          expectVersion: 'proposal-1',
          proposals: [1, 2].map<DocxProposalInput>((index) => ({
            id: `host-${index}`,
            paragraph: {
              kind: 'session', sessionId: 'session', story: 'body', paraId: `p${index}`,
            },
            suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
            op: 'insertText',
            at: 'end',
            text: '!',
          })),
        },
      },
    });
    expect(proposed.ok && proposed.proposal?.result?.ok).toBe(true);
    expect(proposed.ok && proposed.proposal?.mirror.proposals.entries.map(
      ({ record }) => ({ id: record.id, revisionIds: record.revisionIds })
    )).toEqual([
      { id: 'host-1', revisionIds: ['worker-1'] },
      { id: 'host-2', revisionIds: ['worker-2'] },
    ]);
    expect(revisions).toEqual(['document-1', 'document-2', 'worker-1', 'worker-2']);

    const counted = await w.send({ type: 'revisionCount' });
    expect(counted.ok && counted.revisionCount).toBe(2);
    expect(excluded).toEqual([new Set(['worker-1', 'worker-2'])]);
  });

  test('revisionCount passes an empty exclusion set without a proposal registry', async () => {
    const w = worker();
    const revisions = ['document-1', 'document-2'];
    const excluded: ReadonlySet<string>[] = [];
    w.harness.session.revisionCount = (excluding) => {
      expect(excluding).toBeInstanceOf(Set);
      excluded.push(excluding!);
      return revisions.filter((id) => !excluding!.has(id)).length;
    };
    expect((await w.bootstrap()).ok).toBe(true);
    const counted = await w.send({ type: 'revisionCount' });
    expect(counted.ok && counted.revisionCount).toBe(2);
    expect(excluded).toEqual([new Set<string>()]);
  });
});

describe('resident worker proposal failures', () => {
  const operations: ResidentProposalOperation[] = [
    { kind: 'propose', request: { expectVersion: 'proposal', proposals: [] } },
    {
      kind: 'setStates',
      request: { expectVersion: 'proposal', expectPreviewVersion: 0, changes: [] },
    },
    { kind: 'withdraw', request: { expectVersion: 'proposal', ids: [] } },
  ];

  function proposalWorker() {
    const w = worker();
    const engine = { version: () => 'proposal' };
    Object.assign(w.harness.session, {
      proposalEngine: engine,
      geometryReader: engine,
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
    });
    return { w, engine };
  }

  for (const operation of operations) {
    test(
      `a reply failure after ${operation.kind} returns is terminal without document updates`,
      async () => {
        const { w } = proposalWorker();
        expect((await w.bootstrap()).ok).toBe(true);
        w.harness.session.encodeStateVector = () => {
          throw new Error('state vector failed');
        };
        const failed = await w.send({ type: 'proposal', operation });
        expect(failed).toMatchObject({
          ok: false,
          terminal: true,
          error: 'state vector failed',
        });
        expect(w.answered.filter((id) => id === failed.id)).toHaveLength(1);
      }
    );
  }

  test('a reply failure after a proposal snapshot is not terminal', async () => {
    const { w } = proposalWorker();
    expect((await w.bootstrap()).ok).toBe(true);
    const encodeStateVector = w.harness.session.encodeStateVector;
    w.harness.session.encodeStateVector = () => {
      throw new Error('state vector failed');
    };
    const failed = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
    expect(failed).toMatchObject({ ok: false, error: 'state vector failed' });
    expect(!failed.ok && failed.terminal).toBeUndefined();
    w.harness.session.encodeStateVector = encodeStateVector;
    const snapshot = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
    expect(snapshot.ok && snapshot.proposal?.mirror.proposals).toEqual({
      previewVersion: 0,
      entries: [],
    });
  });

  test('a proposal failure before the registry operation commits is not terminal', async () => {
    const { w } = proposalWorker();
    expect((await w.bootstrap()).ok).toBe(true);
    const failed = await w.send({
      type: 'proposal',
      operation: {
        kind: 'propose',
        request: { expectVersion: 'proposal', proposals: null },
      } as never,
    });
    expect(failed).toMatchObject({
      ok: false,
      error: 'a proposal request needs a proposals array',
    });
    expect(!failed.ok && failed.terminal).toBeUndefined();
    const snapshot = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
    expect(snapshot.ok && snapshot.proposal?.mirror.proposals.entries).toEqual([]);
  });

  test('a proposal failure with updates is terminal before the registry returns', async () => {
    const { w, engine } = proposalWorker();
    let onUpdate: ((update: Uint8Array) => void) | undefined;
    Object.assign(w.harness.session, {
      onUpdate: (listener: (update: Uint8Array) => void) => {
        onUpdate = listener;
        return () => {
          onUpdate = undefined;
        };
      },
    });
    expect((await w.bootstrap()).ok).toBe(true);
    engine.version = () => {
      onUpdate!(new Uint8Array([1]));
      throw new Error('registry failed after update');
    };
    const failed = await w.send({ type: 'proposal', operation: operations[0]! });
    expect(failed).toMatchObject({
      ok: false,
      terminal: true,
      error: 'registry failed after update',
    });
    expect(w.answered.filter((id) => id === failed.id)).toHaveLength(1);
  });

  test('a proposal trap after the registry returns refuses queued requests', async () => {
    const { w } = proposalWorker();
    expect((await w.bootstrap()).ok).toBe(true);
    w.harness.session.encodeStateVector = () => {
      throw new WebAssembly.RuntimeError('unreachable');
    };
    const failed = w.send({ type: 'proposal', operation: operations[0]! });
    const queued = w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
    for (const reply of await Promise.all([failed, queued])) {
      expect(reply).toMatchObject({
        ok: false,
        terminal: true,
        error: 'Resident engine worker trapped: unreachable',
      });
      expect(w.answered.filter((id) => id === reply.id)).toHaveLength(1);
    }
  });
});

describe('resident worker memory', () => {
  test('a bootstrap starts the session under its heap limit', async () => {
    const w = worker();
    await w.bootstrap(3, { heapLimitBytes: 1024 });
    await w.bootstrap();
    expect((w.harness as { heapLimits?: unknown[] }).heapLimits).toEqual([1024, undefined]);
  });

  test('every reply carries the worker memory', async () => {
    const w = worker();
    const reply = await w.bootstrap();
    expect(reply.memory).toEqual(w.harness.memories);
  });

  test('a trap after a failed allocation replies out of memory', async () => {
    const w = worker();
    await w.bootstrap();
    w.harness.memories = [
      { label: 'docx-edit', bufferBytes: 4294901760, liveBytes: 4172000000, peakBytes: 4172000000, failedAllocationBytes: 65536 },
    ];
    w.harness.session.buildDisplayListFrame = () => {
      throw new WebAssembly.RuntimeError('unreachable');
    };
    const trapped = await w.build([1]);
    expect(trapped.ok).toBe(false);
    expect(!trapped.ok && trapped.terminal).toBe(true);
    expect(!trapped.ok && trapped.outOfMemory).toBe(true);
    expect(!trapped.ok && trapped.error).toBe(
      'Resident engine worker ran out of memory allocating 65536 bytes: unreachable'
    );
    expect(trapped.memory).toEqual(w.harness.memories);
  });

  test('a trap that the raster paints past still answers the request as out of memory', async () => {
    const w = worker();
    await w.bootstrap();
    w.harness.memories = [
      { label: 'docx-edit', bufferBytes: 65536, liveBytes: 65000, peakBytes: 65000, failedAllocationBytes: 112 },
    ];
    Object.assign(w.harness.session, {
      outlineGlyphJson: () => {
        throw new WebAssembly.RuntimeError('unreachable');
      },
    });
    const rasterize = w.harness.rasterize;
    w.harness.rasterize = async (...args: Parameters<typeof rasterize>) => {
      try {
        (w.harness as { glyphs?: (fontId: number, glyphId: number) => string }).glyphs?.(1, 1);
      } catch {
        // painted with browser text instead
      }
      return rasterize(...args);
    };
    const attached = await w.attach([1]);
    expect(!attached.ok && attached.terminal && attached.outOfMemory).toBe(true);
    const next = await w.build([1]);
    expect(!next.ok && next.terminal).toBe(true);
  });

  test('a trap the raster paints past without a failed allocation keeps the worker', async () => {
    const w = worker();
    await w.bootstrap();
    Object.assign(w.harness.session, {
      outlineGlyphJson: () => {
        throw new WebAssembly.RuntimeError('unreachable');
      },
    });
    const rasterize = w.harness.rasterize;
    w.harness.rasterize = async (...args: Parameters<typeof rasterize>) => {
      try {
        (w.harness as { glyphs?: (fontId: number, glyphId: number) => string }).glyphs?.(1, 1);
      } catch {
        // painted with browser text instead
      }
      return rasterize(...args);
    };
    expect((await w.attach([1])).ok).toBe(true);
    expect((await w.build([1])).ok).toBe(true);
  });

  test('a trap without a failed allocation is not reported as out of memory', async () => {
    const w = worker();
    await w.bootstrap();
    w.harness.session.buildDisplayListFrame = () => {
      throw new WebAssembly.RuntimeError('unreachable');
    };
    const trapped = await w.build([1]);
    expect(!trapped.ok && trapped.terminal).toBe(true);
    expect(!trapped.ok && trapped.outOfMemory).toBeUndefined();
    expect(!trapped.ok && trapped.error).toBe('Resident engine worker trapped: unreachable');
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

  test('a collaboration update restarts the completion without holding it for idle input', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker();
    await bootstrap();
    let inputAt = 0;
    onResume.push(() => {
      inputAt = performance.now();
      void w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null });
    });
    const completed = await w.send({
      type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false, sliceBlocks: 4,
    });
    expect(performance.now() - inputAt).toBeLessThan(290);
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(completed).toHaveProperty('frame');
    expect(calls.slice(0, 3)).toEqual(['begin', 'resume:4', 'update']);
    expect(calls.filter((call) => call === 'begin')).toHaveLength(2);
    expect(calls).not.toContain('whole');
    expect(new Set(w.answered).size).toBe(w.answered.length);
  });

  test('a glyph trap while a frame request finishes the pass answers both requests once', async () => {
    const { w, onResume, bootstrap } = steppedWorker(100);
    await bootstrap();
    await w.attach([1]);
    // Presenting fails too once the trap is recorded; the trap still answers.
    w.harness.failPresent = 1;
    w.harness.memories = [
      { label: 'docx-edit', bufferBytes: 65536, liveBytes: 65000, peakBytes: 65000, failedAllocationBytes: 112 },
    ];
    Object.assign(w.harness.session, {
      outlineGlyphJson: () => {
        throw new WebAssembly.RuntimeError('unreachable');
      },
    });
    const rasterize = w.harness.rasterize;
    w.harness.rasterize = async (...args: Parameters<typeof rasterize>) => {
      try {
        (w.harness as { glyphs?: (fontId: number, glyphId: number) => string }).glyphs?.(1, 1);
      } catch {
        // painted with browser text instead
      }
      return rasterize(...args);
    };
    let frame: Promise<ResidentEngineWorkerResponse> | undefined;
    onResume.push(() => {
      frame = w.send({ type: 'buildFrame', extras: 'given', expectedFrameEpoch: 2, paintCaret: false });
    });
    const completion = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 2,
    });
    const framed = await frame!;
    expect(!completion.ok && completion.terminal && completion.outOfMemory).toBe(true);
    expect(!framed.ok && framed.terminal && framed.outOfMemory).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(new Set(w.answered).size).toBe(w.answered.length);
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

  test('a measurement failure answers the completion once without a synchronous retry', async () => {
    const { w, calls, onResume, bootstrap } = steppedWorker();
    await bootstrap();
    onResume.push(() => {
      Object.assign(w.harness.session, {
        resumeRegionLayout: () => { throw new Error('measurement failed'); },
      });
    });
    const failed = await w.send({
      type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false, sliceBlocks: 1,
    });
    expect(failed).toMatchObject({ ok: false, error: 'measurement failed' });
    expect((await w.send({ type: 'revisionCount' })).ok).toBe(true);
    expect(new Set(w.answered).size).toBe(w.answered.length);
    expect(calls).not.toContain('whole');
  });

  test('without a slice size the rest is laid out in one step', async () => {
    const { w, calls, bootstrap } = steppedWorker();
    await bootstrap();
    const completed = await w.send({ type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(calls).toEqual(['whole']);
  });
});

describe('worker proposals during sliced completion', () => {
  const layoutInput = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
    renderEnv: {},
  });

  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
  )))));

  test('content-control reads equal a main session opened from the same package', async () => {
    const bytes = new Uint8Array(readFileSync(resolve(
      import.meta.dir, '__fixtures__/content-controls/template.docx'
    )));
    const engine = await createResidentEngineSession(undefined, 97200);
    const main = await createYrsSession({ clientId: 97200 });
    const w = worker();
    Object.assign(w.harness.session, engine);
    const normalize = (result: DocxContentControlsResult) => ({ ...result, version: '<version>' });
    try {
      expect((await w.send({ type: 'open', bytes: bytes.buffer })).ok).toBe(true);
      main.openDocx(bytes, true);
      const listed = main.listContentControls();
      if (!listed.ok) throw new Error(listed.failure.message);
      expect(new Set(listed.content.controls.map(({ placement }) => placement))).toEqual(
        new Set(['inline', 'block'])
      );
      const query = { kind: 'ooxmlId', ooxmlId: listed.content.controls[0]!.ooxmlId! } as const;
      for (const options of [{}, { maxControls: 1 }, { stories: ['body'] as const }]) {
        for (const read of [
          { kind: 'listContentControls', options },
          { kind: 'findContentControls', query, options },
        ] satisfies ResidentDocumentRead[]) {
          const reply = await w.send({ type: 'documentRead', read });
          if (!reply.ok || !reply.read) throw new Error('expected a content-control read');
          expect(normalize(reply.read.value as DocxContentControlsResult)).toEqual(normalize(
            read.kind === 'listContentControls'
              ? main.listContentControls(options)
              : main.findContentControls(query, options)
          ));
        }
      }
    } finally {
      main.destroy();
      engine.destroy();
    }
  });

  async function proposalWorker(extraBody = '', comments?: string, bytes?: Uint8Array, options: { clientId?: number; headersFooters?: boolean } = {}) {
    const parts: PartsMap = new Map();
    parts.set('[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ));
    parts.set('_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ));
    const body = Array.from({ length: 40 }, (_, index) =>
      `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Paragraph ${index + 1}</w:t></w:r></w:p>${index === 0 ? extraBody : ''}`
    ).join('');
    parts.set('word/document.xml', toBytes(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`
    ));
    if (comments !== undefined) {
      const types = new TextDecoder().decode(parts.get('[Content_Types].xml')!);
      parts.set('[Content_Types].xml', toBytes(types.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')));
      parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
      parts.set('word/comments.xml', toBytes(`<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`));
    }
    if (options.headersFooters) {
      const document = new TextDecoder().decode(parts.get('word/document.xml')!);
      parts.set('word/document.xml', toBytes(document.replace('<w:sectPr/>',
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>'
      ).replace('<w:document ', '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')));
      const types = new TextDecoder().decode(parts.get('[Content_Types].xml')!);
      parts.set('[Content_Types].xml', toBytes(types.replace('</Types>',
        '<Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
      )));
      parts.set('word/_rels/document.xml.rels', toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer.xml"/></Relationships>'
      ));
      parts.set('word/header.xml', toBytes('<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>'));
      parts.set('word/footer.xml', toBytes('<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>'));
    }
    const engine = await createResidentEngineSession(undefined, options.clientId);
    engine.openDocx(bytes ?? new Uint8Array(rezipPartsToArrayBuffer(parts)));
    const w = worker();
    const calls: string[] = [];
    const onResume: Array<() => void> = [];
    Object.assign(w.harness.session, engine, {
      layoutDocumentWithRegionsPrefixRetainedJson: (input: string, pages: number) => {
        calls.push(`prefix:${pages}`);
        return engine.layoutDocumentWithRegionsPrefixRetainedJson(input, pages);
      },
      layoutDocumentWithRegionsRetainedJson: (input: string) => {
        calls.push('whole');
        return engine.layoutDocumentWithRegionsRetainedJson(input);
      },
      beginRegionLayout: (input: string) => {
        calls.push('begin');
        return engine.beginRegionLayout(input);
      },
      resumeRegionLayout: (blocks: number) => {
        calls.push('resume');
        const progress = engine.resumeRegionLayout(Math.min(blocks, 1));
        onResume.shift()?.();
        return progress;
      },
      buildDisplayListFrame: (extras: string, epoch: number) => {
        const frame = engine.buildDisplayListFrame(extras, epoch);
        w.harness.delta = decodeFrameDelta(frame);
        return frame;
      },
      buildDisplayPagesFrame: (pages: number[], epoch: number) => {
        const frame = engine.buildDisplayPagesFrame(pages, epoch);
        w.harness.delta = decodeFrameDelta(frame);
        return frame;
      },
    });
    const snapshot: YrsResidentWorkerSnapshot = {
      workerAuthoritative: true,
      clientId: 1,
      state: new Uint8Array(),
      fontsRevision: 0,
      fonts: [new Uint8Array(readFileSync(resolve(
        import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
      )))],
      renderInputs: [],
      measureInputs: [],
      layoutInput,
      layoutWithRegions: true,
      layoutRevision: 1,
      selection: null,
    };
    const booted = await w.send({
      type: 'bootstrap', snapshot, extras: '', layoutExtras: '{}',
      provisionalPages: 1, displayWindow: [0, 1], expectedFrameEpoch: 0,
    });
    expect(booted.ok).toBe(true);
    if (!bytes) expect(booted.ok && booted.layoutProvisional).toBe(true);
    const proposal = (index = 1): Extract<DocxProposalInput, { op: 'replaceText' }> => ({
      id: `p${index}`,
      paragraph: {
        kind: 'persisted',
        story: { kind: 'body', partUri: '/word/document.xml' },
        paraId: index.toString(16).padStart(8, '0'),
      },
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'replaceText', search: `Paragraph ${index}`, replaceWith: `Changed ${index}`,
    });
    const complete = () => w.send({
      type: 'completeLayout', expectedFrameEpoch: 1, paintCaret: false, sliceBlocks: 1,
    });
    const expectFullLayout = async (reply: ResidentEngineWorkerResponse, input = layoutInput) => {
      const fresh = await createResidentEngineSession();
      try {
        fresh.loadState(engine.encodeState());
        expect(reply.ok && reply.layoutJson).toBe(
          fresh.layoutDocumentWithRegionsRetainedJson(input)
        );
      } finally {
        fresh.destroy();
      }
      expect(calls).not.toContain('whole');
    };
    return { w, engine, calls, onResume, snapshot, booted, proposal, complete, expectFullLayout };
  }

  test('sidebar and headings reads match a main session and reject stale versions', async () => {
    const bytes = sidebarDocx();
    const { w, engine } = await proposalWorker('', undefined, bytes);
    const main = await createYrsSession();
    try {
      main.openDocx(bytes, true);
      const version = engine.geometryReader.version();
      const expected = {
        sidebar: readSidebar(main, ['7', 'missing'], main.version()),
        headings: readOutlineHeadings(main, main.version()),
      };
      for (const expectVersion of [version, 'stale']) {
        const reads: ResidentDocumentRead[] = [
          { kind: 'sidebar', commentIds: ['7', 'missing'], expectVersion },
          { kind: 'headings', expectVersion },
        ];
        for (const read of reads) {
          const reply = await w.send({ type: 'documentRead', read });
          expect(reply).toMatchObject({ ok: true, read: {
            version,
            value: expectVersion === version ? expected[read.kind as keyof typeof expected] : null,
          } });
        }
      }
    } finally {
      main.destroy();
      engine.destroy();
    }
  });

  test('paged export replies with the resident session export of its retained layout', async () => {
    const { w, engine, complete } = await proposalWorker();
    try {
      const completed = await complete();
      expect(completed.ok && completed.layoutProvisional).not.toBe(true);
      const options = { revisionView: 'markup' } as const;
      for (const currentRequest of [
        layoutInput,
        JSON.stringify({ ...JSON.parse(layoutInput), renderEnv: { revisionPreview: { '1': 'accepted' } } }),
      ]) {
        const reply = await w.send({
          type: 'documentRead',
          read: { kind: 'exportStructuredWithPages', options, currentRequest },
        });
        if (!reply.ok || !reply.read) throw new Error('expected a paged export read');
        expect(reply.read.value).toBe(engine.exportStructuredWithPagesJson(options, currentRequest));
        expect(JSON.parse(reply.read.value as string).version).toBe(reply.read.version);
      }
      const preview = JSON.parse(engine.exportStructuredWithPagesJson(options, JSON.stringify({
        ...JSON.parse(layoutInput), renderEnv: { revisionPreview: { '1': 'accepted' } },
      })));
      expect(preview).toMatchObject({ ok: false, failure: { code: 'unsupported-revision-layout' } });
    } finally {
      engine.destroy();
    }
  });

  test('proposal mirrors retain the same navigation target as repeated worker reads', async () => {
    const { w, engine, proposal } = await proposalWorker();
    try {
      let projectionReads = 0;
      const storyIds = engine.geometryReader.storyIds;
      engine.geometryReader.storyIds = () => { projectionReads += 1; return storyIds(); };
      const snapshot = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
      expect(snapshot.ok).toBe(true);
      expect(projectionReads).toBe(0);
      const applied = await w.send({
        type: 'proposal', operation: {
          kind: 'propose', request: { expectVersion: engine.proposalEngine.version(), proposals: [proposal()] },
        },
      });
      expect(applied.ok).toBe(true);
      if (!applied.ok || !applied.proposal?.result?.ok) throw new Error('expected a proposal');
      expect(applied.proposal.changedStories).toEqual(['body']);
      expect(applied.proposal.geometry.navigationTargets).toBeUndefined();
      expect(projectionReads).toBe(0);
      const paragraph = applied.proposal.result.snapshot.proposals[0]!.paragraph;
      const read = { kind: 'navigationTarget', story: paragraph.story, paraId: paragraph.paraId } as const;
      const toggled = await w.send({
        type: 'proposal', operation: {
          kind: 'setStates', request: {
            expectVersion: applied.proposal.mirror.version,
            expectPreviewVersion: applied.proposal.mirror.proposals.previewVersion,
            changes: [{ id: 'p1', state: 'rejected' }],
          },
        },
      });
      expect(toggled.ok).toBe(true);
      if (!toggled.ok || !toggled.proposal?.result?.ok) throw new Error('expected a toggle');
      expect(toggled.proposal.mirror.version).toBe(applied.proposal.mirror.version);
      expect(toggled.proposal.geometry.previewVersion).toBe(applied.proposal.geometry.previewVersion + 1);
      expect(toggled.proposal.changedStories).toEqual([]);
      expect(projectionReads).toBe(1);
      const first = await w.send({ type: 'documentRead', read });
      const second = await w.send({ type: 'documentRead', read });
      expect(first.ok).toBe(true);
      expect(second.ok).toBe(true);
      if (!first.ok || !second.ok) throw new Error('expected navigation reads');
      expect(first.read).toEqual(second.read);
      expect(first.read).toEqual({
        version: applied.proposal.mirror.version,
        value: toggled.proposal.geometry.navigationTargets?.p1,
      });
      expect(projectionReads).toBe(1);
    } finally {
      engine.destroy();
    }
  });

  const requirementsInput = (preview?: ReturnType<typeof proposalRevisionPreview>) => JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    renderEnv: { revisionPreview: preview },
  });

  async function decided(
    w: Awaited<ReturnType<typeof proposalWorker>>['w'],
    engine: Awaited<ReturnType<typeof proposalWorker>>['engine'],
    proposal: Awaited<ReturnType<typeof proposalWorker>>['proposal']
  ) {
    const applied = await w.send({
      type: 'proposal', operation: {
        kind: 'propose', request: { expectVersion: engine.proposalEngine.version(), proposals: [proposal()] },
      },
    });
    if (!applied.ok || !applied.proposal?.result?.ok) throw new Error('expected a proposal');
    let previewVersion = applied.proposal.mirror.proposals.previewVersion;
    return async (state: 'proposed' | 'accepted' | 'rejected') => {
      const reply = await w.send({
        type: 'proposal', operation: {
          kind: 'setStates', request: {
            expectVersion: applied.proposal!.mirror.version,
            expectPreviewVersion: previewVersion,
            changes: [{ id: 'p1', state }],
          },
        },
      });
      const result = reply.ok ? reply.proposal?.result : undefined;
      if (!reply.ok || !reply.proposal || !result?.ok) throw new Error('expected a decision');
      previewVersion = reply.proposal.mirror.proposals.previewVersion;
      return { fontRequirements: reply.proposal.fontRequirements, preview: proposalRevisionPreview(result.snapshot) };
    };
  }

  test('meta and JSON replies preserve decisions, frame bytes, sizes and header epochs', async () => {
    const run = async (mode: 'json' | 'meta') => {
      const { w, engine, proposal, snapshot, calls } = await proposalWorker('', undefined, undefined, {
        clientId: 9501, headersFooters: true,
      });
      const frames: Uint8Array[] = [];
      const layouts: string[] = [];
      const metas: LayoutMetaV1[] = [];
      let revision = snapshot.layoutRevision;
      let epoch = w.harness.delta!.frameEpoch;
      let input = layoutInput;
      const sync = async () => {
        const response = await w.send({
          type: 'sync', expectedFrameEpoch: epoch, paintCaret: false, extras: '', layoutExtras: '{}',
          ...(mode === 'meta' ? { layoutReply: 'meta' as const } : {}),
          snapshot: { ...snapshot, fonts: [], layoutRevision: ++revision, layoutInput: input },
        });
        if (!response.ok || !response.frame) throw new Error('expected layout frame');
        frames.push(new Uint8Array(response.frame));
        epoch = response.caret!.frameEpoch;
        const full = await w.send({ type: 'layoutJson', layoutRevision: revision });
        if (!full.ok || full.layoutJsonStatus !== 'ok' || full.layoutJson === undefined) {
          throw new Error('expected retained JSON');
        }
        layouts.push(full.layoutJson);
        if (mode === 'json') {
          expect(response.layoutJson).toBe(full.layoutJson);
          expect(response.layoutMeta).toBeUndefined();
        } else {
          expect(response.layoutJson).toBeUndefined();
          if (!isLayoutMetaV1(response.layoutMeta)) throw new Error('expected v1 meta');
          const meta = response.layoutMeta;
          metas.push(meta);
          expect(w.transfers.get(response.id)).toContain(meta.pageSizes.buffer);
          const output = JSON.parse(full.layoutJson) as { layout: Layout; notesConverged: boolean; provisional?: boolean };
          expect(meta.pageCount).toBe(output.layout.pages.length);
          expect(meta.pageCount).toBeGreaterThan(1);
          expect(meta.partial).toBe(output.layout.partial === true);
          expect(meta.provisional).toBe(output.provisional === true);
          expect(meta.notesConverged).toBe(output.notesConverged);
          const sizes = new Float64Array(output.layout.pages.flatMap((page) => [page.size.w, page.size.h]));
          expect(new Uint8Array(meta.pageSizes.buffer)).toEqual(new Uint8Array(sizes.buffer));
          expect(JSON.parse(meta.layoutShell)).toEqual({
            ...output.layout,
            pages: output.layout.pages.map((page) => ({ ...page, fragments: [] })),
          });
        }
        expect(await w.send({ type: 'layoutJson', layoutRevision: revision - 1 })).toMatchObject({
          ok: true, layoutJsonStatus: 'stale',
        });
      };
      try {
        const decide = await decided(w, engine, proposal);
        for (const state of ['accepted', 'rejected', 'proposed', 'accepted'] as const) {
          const decision = await decide(state);
          input = JSON.stringify({ ...JSON.parse(layoutInput),
            regions: { sections: [{ sectionId: 'main', headerFooterRefs: {
              headerDefault: 'rIdHeader', footerDefault: 'rIdFooter',
            } }] },
            renderEnv: { revisionPreview: decision.preview },
          });
          await sync();
        }
        engine.applyRawOps('hf:rIdHeader', [{ op: 'insert', index: 0, text: 'Changed ' }]);
        await sync();
        await sync();
        input = JSON.stringify({ ...JSON.parse(input), regions: { sections: [{ sectionId: 'main' }] } });
        await sync();
        await sync();
        if (mode === 'meta') {
          expect(calls).not.toContain('whole');
          expect(metas[0]!.headersFooters).toContain('Header');
          expect(metas[0]!.headersFooters).toContain('Footer');
          for (const meta of metas.slice(1, 4)) {
            expect(meta.headersFootersEpoch).toBe(metas[0]!.headersFootersEpoch);
            expect(meta.headersFooters).toBeUndefined();
          }
          expect(metas[4]!.headersFootersEpoch).toBe(metas[0]!.headersFootersEpoch + 1);
          expect(metas[4]!.headersFooters).toContain('Changed ');
          expect(metas[5]!.headersFootersEpoch).toBe(metas[4]!.headersFootersEpoch);
          expect(metas[5]!.headersFooters).toBeUndefined();
          expect(metas[6]!.headersFootersEpoch).toBe(metas[4]!.headersFootersEpoch + 1);
          expect(metas[6]!.headersFooters).toBe('null');
          expect(metas[7]!.headersFootersEpoch).toBe(metas[6]!.headersFootersEpoch);
          expect(metas[7]!.headersFooters).toBeUndefined();
        }
        return { frames, layouts };
      } finally {
        void w.send({ type: 'destroy' });
      }
    };
    const full = await run('json');
    const meta = await run('meta');
    expect(meta.layouts).toEqual(full.layouts);
    expect(meta.frames).toEqual(full.frames);
  });

  test('a decision answers with the font requirements of the layout input the host builds next, and an undo reads cached ones', async () => {
    const { w, engine, proposal } = await proposalWorker();
    try {
      const decide = await decided(w, engine, proposal);
      const asked = await w.send({ type: 'fontRequirements', layoutInput: requirementsInput() });
      if (!asked.ok) throw new Error('expected font requirements');
      let reads = 0;
      Object.assign(w.harness.session, {
        layoutFontRequirementsJson: (input: string) => {
          reads += 1;
          return engine.layoutFontRequirementsJson(input);
        },
      });

      const accepted = await decide('accepted');
      expect(accepted.preview).toBeDefined();
      expect(accepted.fontRequirements).toEqual({
        layoutInput: requirementsInput(accepted.preview),
        requirementsJson: engine.layoutFontRequirementsJson(requirementsInput(accepted.preview)),
      });
      await w.send({ type: 'fontRequirements', layoutInput: requirementsInput(accepted.preview) });

      const undone = await decide('proposed');
      expect(undone.fontRequirements).toEqual({
        layoutInput: requirementsInput(),
        requirementsJson: asked.requirementsJson!,
      });
      const base = await w.send({ type: 'fontRequirements', layoutInput: requirementsInput() });
      expect(base.ok && base.requirementsJson).toBe(asked.requirementsJson);
      const rejected = await decide('rejected');
      expect(rejected.fontRequirements?.layoutInput).toBe(requirementsInput(rejected.preview));
      const again = await w.send({ type: 'fontRequirements', layoutInput: requirementsInput(rejected.preview) });
      expect(again.ok && again.requirementsJson).toBe(rejected.fontRequirements!.requirementsJson);
      expect(reads).toBe(2);
    } finally {
      void w.send({ type: 'destroy' });
    }
  });

  test('a decision whose font requirements cannot be read leaves them to the host', async () => {
    const { w, engine, proposal } = await proposalWorker();
    try {
      const decide = await decided(w, engine, proposal);
      await w.send({ type: 'fontRequirements', layoutInput: requirementsInput() });
      Object.assign(w.harness.session, {
        layoutFontRequirementsJson: () => {
          throw new Error('unreadable');
        },
      });
      const accepted = await decide('accepted');
      expect(accepted.fontRequirements).toBeUndefined();
      const asked = await w.send({
        type: 'fontRequirements',
        layoutInput: requirementsInput(accepted.preview),
      });
      expect(asked.ok).toBe(false);
    } finally {
      void w.send({ type: 'destroy' });
    }
  });

  test('cached navigation targets populate the first mirror and changed versions defer rebuilding', async () => {
    const { w, engine, proposal } = await proposalWorker();
    try {
      let projectionReads = 0;
      const storyIds = engine.geometryReader.storyIds;
      engine.geometryReader.storyIds = () => { projectionReads += 1; return storyIds(); };
      const read = { kind: 'navigationTarget', story: 'body', paraId: '00000001' } as const;
      const initial = await w.send({ type: 'documentRead', read });
      expect(initial.ok).toBe(true);
      expect(projectionReads).toBe(1);
      const unchanged = await w.send({
        type: 'proposal', operation: {
          kind: 'propose', request: {
            expectVersion: engine.proposalEngine.version(),
            proposals: [{ ...proposal(), replaceWith: 'Paragraph 1' }],
          },
        },
      });
      expect(unchanged.ok).toBe(true);
      if (!initial.ok || !unchanged.ok || !unchanged.proposal?.result?.ok) {
        throw new Error('expected unchanged proposal and navigation');
      }
      expect(unchanged.proposal.changedStories).toEqual([]);
      expect(initial.read).toEqual({
        version: unchanged.proposal.mirror.version,
        value: unchanged.proposal.geometry.navigationTargets?.p1,
      });
      expect(unchanged.proposal.geometry.navigationTargets?.p1).toMatchObject({
        loc: { story: read.story, paraId: read.paraId, offset: 0 }, position: 1,
      });
      expect(projectionReads).toBe(1);

      const changed = await w.send({
        type: 'proposal', operation: {
          kind: 'propose', request: { expectVersion: engine.proposalEngine.version(), proposals: [proposal(2)] },
        },
      });
      expect(changed.ok).toBe(true);
      if (!changed.ok || !changed.proposal?.result?.ok) throw new Error('expected changed proposal');
      expect(changed.proposal.changedStories).toEqual(['body']);
      expect(changed.proposal.mirror.version).not.toBe(unchanged.proposal.mirror.version);
      expect(changed.proposal.geometry.navigationTargets).toBeUndefined();
      expect(projectionReads).toBe(1);
      const snapshot = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
      expect(snapshot.ok).toBe(true);
      if (!snapshot.ok || !snapshot.proposal) throw new Error('expected snapshot');
      expect(projectionReads).toBe(2);
      const navigation = await w.send({ type: 'documentRead', read });
      expect(navigation.ok).toBe(true);
      expect(navigation.ok && navigation.read).toEqual({
        version: snapshot.proposal.mirror.version,
        value: snapshot.proposal.geometry.navigationTargets?.p1,
      });
      expect(projectionReads).toBe(2);
      const withdrawn = await w.send({
        type: 'proposal', operation: {
          kind: 'withdraw', request: { expectVersion: snapshot.proposal.mirror.version, ids: ['p2'] },
        },
      });
      expect(withdrawn.ok).toBe(true);
      if (!withdrawn.ok || !withdrawn.proposal?.result?.ok) throw new Error('expected withdrawal');
      expect(withdrawn.proposal.changedStories).toEqual(['body']);
      expect(withdrawn.proposal.geometry.navigationTargets).toBeUndefined();
      expect(projectionReads).toBe(2);
      const first = await w.send({ type: 'documentRead', read });
      expect(first.ok).toBe(true);
      expect(projectionReads).toBe(3);
      const warmed = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
      expect(warmed.ok).toBe(true);
      if (!warmed.ok || !warmed.proposal) throw new Error('expected warmed snapshot');
      expect(first.ok && first.read).toEqual({
        version: warmed.proposal.mirror.version,
        value: warmed.proposal.geometry.navigationTargets?.p1,
      });
      expect(projectionReads).toBe(3);
    } finally {
      engine.destroy();
    }
  });

  test('snapshot and document reads answer before the background layout without restarting it', async () => {
    const { w, engine, calls, onResume, proposal, complete, expectFullLayout } = await proposalWorker();
    try {
      const order: string[] = [];
      let snapshot: Promise<ResidentEngineWorkerResponse> | undefined;
      const reads: Array<Promise<ResidentEngineWorkerResponse>> = [];
      const tail = engine.paragraphIdentities().paragraphs.at(-1)!.session!;
      const requests: ResidentDocumentRead[] = [
        { kind: 'paragraphIdentities' },
        { kind: 'resolveParagraphAnchors', anchors: [proposal().paragraph] },
        { kind: 'readParagraphs', request: { view: 'accepted' } },
        { kind: 'navigationTarget', story: tail.story, paraId: tail.paraId },
      ];
      onResume.push(() => {
        snapshot = w.send({ type: 'proposal', operation: { kind: 'snapshot' } }).then((reply) => {
          order.push('snapshot');
          return reply;
        });
        for (const read of requests) {
          reads.push(w.send({ type: 'documentRead', read }).then((reply) => {
            order.push(read.kind);
            return reply;
          }));
        }
      });
      const completed = await complete().then((reply) => {
        order.push('complete');
        return reply;
      });
      expect((await snapshot!).ok).toBe(true);
      const replies = await Promise.all(reads);
      for (const reply of replies) expect(reply.ok).toBe(true);
      expect(replies.at(-1)).toMatchObject({
        ok: true, read: { value: { loc: { story: tail.story, paraId: tail.paraId, offset: 0 } } },
      });
      expect(order).toEqual(['snapshot', ...requests.map((read) => read.kind), 'complete']);
      expect(calls.filter((call) => call === 'begin')).toHaveLength(1);
      await expectFullLayout(completed);
    } finally {
      engine.destroy();
    }
  });

  test('find matches read equals the replica and rejects a different version', async () => {
    const { w, engine } = await proposalWorker();
    const replica = await createYrsSession();
    try {
      replica.loadState(engine.encodeState());
      const expectVersion = engine.proposalEngine.version();
      const options = { matchCase: false, matchWholeWord: true };
      const matches = findBodyMatches(replica, (loc) => yrsLocToProjectedDisplayPosition(
        replica,
        (root) => createYrsPositionProjection(replica, root),
        loc,
        'body',
        (story) => createYrsInputPositionMap(story, replica.paragraphSpans(story))
      ), 'paragraph', options);
      expect(matches).toHaveLength(40);
      const read = { kind: 'findMatches', searchText: 'paragraph', options, expectVersion } as const;
      const reply = await w.send({ type: 'documentRead', read });
      expect(reply.ok && reply.read).toEqual({ version: expectVersion, value: matches });
      const stale = await w.send({
        type: 'documentRead', read: { ...read, expectVersion: `${expectVersion}-stale` },
      });
      expect(stale.ok && stale.read).toEqual({ version: expectVersion, value: null });
    } finally {
      replica.destroy();
      engine.destroy();
    }
  });

  test('comment deletion removes only its worker anchor and missing ids are harmless', async () => {
    const extraBody = [1, 2].map((id) =>
      `<w:p w14:paraId="0000010${id}"><w:commentRangeStart w:id="${id}"/><w:r><w:t>Marked ${id}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`
    ).join('');
    const comments = [1, 2].map((id) =>
      `<w:comment w:id="${id}" w:author="Reviewer"><w:p><w:r><w:t>Comment ${id}</w:t></w:r></w:p></w:comment>`
    ).join('');
    const { w, engine } = await proposalWorker(extraBody, comments);
    try {
      expect(engine.resolveComment('1')).not.toHaveLength(0);
      const untouched = engine.resolveComment('2');
      const reply = await w.send({ type: 'proposal', operation: { kind: 'removeComment', id: '1' } });
      if (!reply.ok || !reply.proposal) throw new Error('expected comment deletion reply');
      expect(reply.proposal.result).toBeUndefined();
      expect(reply.proposal.changedStories.length).toBeGreaterThan(0);
      expect(reply.proposal.updates.length).toBeGreaterThan(0);
      let anchors: ReturnType<typeof engine.resolveComment> = [];
      try { anchors = engine.resolveComment('1'); } catch {}
      expect(anchors).toEqual([]);
      expect(engine.resolveComment('2')).toEqual(untouched);
      expect(readSidebar(engine.geometryReader, ['1'], engine.proposalEngine.version())!.comments)
        .toEqual([{ id: '1', anchors: [] }]);

      const version = engine.proposalEngine.version();
      const missing = await w.send({ type: 'proposal', operation: { kind: 'removeComment', id: 'missing' } });
      expect(missing).toMatchObject({ ok: true, proposal: { changedStories: [], updates: [] } });
      expect(engine.proposalEngine.version()).toBe(version);
      expect(engine.resolveComment('2')).toEqual(untouched);
    } finally {
      engine.destroy();
    }
  });

  test('viewer document reads reply with the current version and value', async () => {
    const extraBody = '<w:p w14:paraId="00000100"><w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      '<w:commentRangeStart w:id="1"/><w:r><w:t>the phrase</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
      '<w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>' +
      '<w:ins w:id="9" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>' +
      '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>';
    const comments = '<w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>Check phrase</w:t></w:r></w:p></w:comment>';
    const { w, engine } = await proposalWorker(extraBody, comments);
    const main = await createYrsSession();
    try {
      main.loadState(engine.encodeState());
      const version = engine.proposalEngine.version();
      const projection = createYrsPositionProjection(main, 'body');
      const position = (offset: number) => yrsLocToProjectedDisplayPosition(main, () => projection, {
        story: 'body', paraId: '00000100', offset,
      })!;
      const revision = main.listRevisions().find((candidate) => candidate.kind === 'insertion')!;
      const requests: Array<{ read: ResidentDocumentRead; value?: unknown; text?: string }> = [
        {
          read: { kind: 'findText', request: { text: 'phrase', within: { kind: 'story', story: 'body' }, view: 'accepted', limit: 1 } },
          value: { ...main.findText({ text: 'phrase', within: { kind: 'story', story: 'body' }, view: 'accepted', limit: 1 }), version },
        },
        {
          read: { kind: 'findParagraphs', query: 'phrase', caseSensitive: true, limit: 1 },
          value: [{ paraId: '00000100', match: 'phrase', before: 'Before the ', after: ' new after' }],
        },
        {
          read: { kind: 'selectionInfo', story: 'body', anchor: position(17), head: position(7), expectVersion: version },
          value: main.selectionText({
            story: 'body', start: { paraId: '00000100', offset: 7 }, end: { paraId: '00000100', offset: 17 },
          }),
        },
        {
          read: { kind: 'commentTarget', story: 'body', commentId: '1', expectVersion: version },
          text: 'the phrase',
        },
        {
          read: { kind: 'revisionTarget', story: 'body', revisionId: revision.revisionId, expectVersion: version },
          text: 'new',
        },
      ];
      for (const { read, value, text } of requests) {
        const reply = await w.send({ type: 'documentRead', read });
        expect(reply.ok).toBe(true);
        if (text === undefined) {
          expect(reply.ok && reply.read).toEqual({ version, value });
        } else {
          const range = reply.ok ? (reply.read!.value as { anchor: number; head: number } | null) : null;
          expect(reply.ok && reply.read!.version).toBe(version);
          expect(range).not.toBeNull();
          const covered = await w.send({
            type: 'documentRead',
            read: { kind: 'selectionText', story: 'body', anchor: range!.anchor, head: range!.head, expectVersion: version },
          });
          expect(covered.ok && (covered.read!.value as { text: string } | null)?.text).toBe(text);
        }
        if ('expectVersion' in read) {
          const stale = await w.send({ type: 'documentRead', read: { ...read, expectVersion: `${version}-stale` } });
          expect(stale.ok).toBe(true);
          expect(stale.ok && stale.read).toEqual({ version, value: null });
        }
      }
    } finally {
      main.destroy();
      engine.destroy();
    }
  });

  test('search reads match main display ranges and carry anchors without restarting background layout', async () => {
    const extraBody = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr/><w:p w14:paraId="00000100"><w:r><w:t>Paragraph cell</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr/><w:p w14:paraId="00000101"><w:r><w:t>PARAGRAPH cell</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl><w:sdt><w:sdtPr><w:id w:val="100"/></w:sdtPr><w:sdtContent>' +
      '<w:p w14:paraId="00000102"><w:r><w:t>Paragraph boxed</w:t></w:r></w:p>' +
      '</w:sdtContent></w:sdt>';
    const { w, engine, calls, onResume, complete, expectFullLayout } = await proposalWorker(extraBody);
    const main = await createYrsSession();
    try {
      main.loadState(engine.encodeState());
      const projection = createYrsPositionProjection(main, 'body');
      const expected = main.searchText('paragraph').filter((hit) =>
        hit.story === 'body' || hit.story.startsWith('body:')
      ).map((hit) => ({
        story: hit.story, paraId: hit.paraId, start: hit.start,
        displayFrom: yrsLocToProjectedDisplayPosition(main, () => projection, {
          story: hit.story, paraId: hit.paraId, offset: hit.start,
        })!,
        displayTo: yrsLocToProjectedDisplayPosition(main, () => projection, {
          story: hit.story, paraId: hit.paraId, offset: hit.end,
        })!,
      })).sort((a, b) => a.displayFrom! - b.displayFrom!);
      const read = { kind: 'searchText', query: 'paragraph', caseSensitive: false } as const;
      let pending!: Promise<ResidentEngineWorkerResponse>;
      const order: string[] = [];
      onResume.push(() => {
        pending = w.send({ type: 'documentRead', read }).then((reply) => {
          order.push('search');
          return reply;
        });
      });
      const completed = await complete().then((reply) => { order.push('complete'); return reply; });
      const reply = await pending;
      expect(reply.ok).toBe(true);
      if (!reply.ok || !reply.read) throw new Error('expected search read');
      const value = reply.read.value as ReturnType<typeof readResidentSearch>;
      expect(value).toEqual(readResidentSearch({ ...engine.geometryReader, ...engine }, 'paragraph', false));
      expect(value.matches).toEqual(expected);
      expect(value.matches).toHaveLength(43);
      expect(value.matches.slice(0, 5).map(({ paraId }) => paraId)).toEqual([
        '00000001', '00000100', '00000101', '00000102', '00000002',
      ]);
      expect(value.matches.every((match) => !('anchor' in match))).toBe(true);
      expect(value.carried).toBe(0);
      expect(reply.read.version).toBe(engine.proposalEngine.version());
      expect(order).toEqual(['search', 'complete']);
      expect(calls.filter((call) => call === 'begin')).toHaveLength(1);
      await expectFullLayout(completed);
      const anchors = async (indices: number[]) => {
        const locs = indices.map((index) => {
          const match = value.matches[index];
          return { story: match.story, paraId: match.paraId, offset: match.start };
        });
        const answer = await w.send({ type: 'documentRead', read: {
          kind: 'stickyAnchors', locs, version: reply.read!.version,
        } });
        expect(answer.ok).toBe(true);
        if (!answer.ok || !answer.read) throw new Error('expected sticky anchor read');
        const sticky = answer.read.value as Array<ReturnType<typeof engine.encodeStickyPosition> | null>;
        expect(answer.read.version).toBe(engine.proposalEngine.version());
        expect(sticky).toHaveLength(locs.length);
        sticky.forEach((anchor, index) => {
          expect(anchor).not.toBeNull();
          expect(anchor).toEqual(engine.encodeStickyPosition(locs[index]));
          expect(engine.resolveStickyPosition(anchor!)).toEqual(locs[index]);
          expect(main.resolveStickyPosition(anchor!)).toEqual(locs[index]);
        });
        return sticky;
      };
      const [firstAnchor, cellAnchor, lastAnchor] = await anchors([0, 2, value.matches.length - 1]);
      const validLoc = { story: 'body', paraId: '00000001', offset: 0 };
      const invalid = await w.send({ type: 'documentRead', read: {
        kind: 'stickyAnchors', version: reply.read.version, locs: [
          { story: 'missing', paraId: 'missing', offset: 0 }, validLoc,
        ],
      } });
      expect(invalid).toMatchObject({
        ok: true, read: { value: [null, engine.encodeStickyPosition(validLoc)] },
      });
      const search = async (caseSensitive: boolean, carry = cellAnchor) => {
        const answer = await w.send({ type: 'documentRead', read: {
          ...read, query: 'Paragraph', caseSensitive, carry,
        } });
        expect(answer.ok).toBe(true);
        if (!answer.ok || !answer.read) throw new Error('expected search read');
        return answer.read.value as ReturnType<typeof readResidentSearch>;
      };
      expect((await search(false)).carried).toBe(2);
      const sensitive = await search(true);
      expect(sensitive.matches).toHaveLength(42);
      expect(sensitive.matches.some(({ paraId }) => paraId === '00000101')).toBe(false);
      expect(sensitive.carried).toBe(2);
      expect(engine.proposalEngine.applyEdits({
        expectVersion: engine.proposalEngine.version(),
        steps: [{
          op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId: '00000001' },
          text: 'Changed 1',
        }],
      }).ok).toBe(true);
      main.loadState(engine.encodeState());
      expect(engine.proposalEngine.version()).not.toBe(reply.read.version);
      await anchors([0, 2]);
      const after = await search(false, firstAnchor);
      expect(after.matches).toHaveLength(42);
      expect(after.carried).toBe(0);
      expect(after.matches[after.carried].paraId).toBe('00000100');
      expect((await search(false, lastAnchor)).carried).toBe(41);
      expect(engine.proposalEngine.applyEdits({
        expectVersion: engine.proposalEngine.version(),
        steps: [{
          op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId: '00000028' },
          text: 'Changed tail',
        }],
      }).ok).toBe(true);
      const last = await search(false, lastAnchor);
      expect(last.matches).toHaveLength(41);
      expect(last.carried).toBe(40);
      const empty = await w.send({ type: 'documentRead', read: { ...read, query: '' } });
      expect(empty).toMatchObject({ ok: true, read: { value: { matches: [], carried: -1 } } });
    } finally {
      main.destroy();
      engine.destroy();
    }
  });

  test('repeated proposals restart the background layout without switching to synchronous completion', async () => {
    const { w, engine, calls, onResume, proposal, complete, expectFullLayout } = await proposalWorker();
    try {
      const order: string[] = [];
      const proposed: Array<Promise<ResidentEngineWorkerResponse>> = [];
      for (let index = 1; index <= 6; index += 1) {
        onResume.push(() => {
          proposed.push(w.send({
            type: 'proposal',
            operation: {
              kind: 'propose',
              request: { expectVersion: engine.proposalEngine.version(), proposals: [proposal(index)] },
            },
          }).then((reply) => {
            order.push(`proposal:${index}`);
            return reply;
          }));
        });
      }
      const completed = await complete().then((reply) => {
        order.push('complete');
        return reply;
      });
      expect(order).toEqual([
        'proposal:1', 'proposal:2', 'proposal:3', 'proposal:4', 'proposal:5', 'proposal:6', 'complete',
      ]);
      for (const reply of await Promise.all(proposed)) {
        expect(reply.ok && reply.proposal?.result?.ok).toBe(true);
        expect(reply.ok && reply.proposal?.changedStories).toEqual(['body']);
      }
      expect(calls.filter((call) => call === 'begin')).toHaveLength(7);
      await expectFullLayout(completed);
    } finally {
      engine.destroy();
    }
  });

  test.each(['propose', 'setStates', 'withdraw'] as const)(
    '%s replies during completion and its relayout paints the visible prefix before completing',
    async (kind) => {
      const { w, engine, calls, onResume, snapshot, booted, proposal, complete, expectFullLayout } =
        await proposalWorker();
      try {
        if (kind !== 'propose') {
          const seeded = await w.send({
            type: 'proposal',
            operation: {
              kind: 'propose',
              request: { expectVersion: engine.proposalEngine.version(), proposals: [proposal()] },
            },
          });
          expect(seeded.ok && seeded.proposal?.result?.ok).toBe(true);
        }
        const order: string[] = [];
        let operation: Promise<ResidentEngineWorkerResponse> | undefined;
        let relayout: Promise<ResidentEngineWorkerResponse> | undefined;
        let input = layoutInput;
        onResume.push(() => {
          const expectVersion = engine.proposalEngine.version();
          operation = w.send({
            type: 'proposal',
            operation: kind === 'propose'
              ? { kind, request: { expectVersion, proposals: [proposal()] } }
              : kind === 'setStates'
                ? { kind, request: { expectVersion, expectPreviewVersion: 0, changes: [{ id: 'p1', state: 'rejected' }] } }
                : { kind, request: { expectVersion, ids: ['p1'] } },
          }).then((reply) => {
            order.push(kind);
            expect(reply.ok && reply.proposal?.result?.ok).toBe(true);
            if (!reply.ok || !reply.proposal?.result?.ok) throw new Error('proposal failed');
            if (kind === 'withdraw') expect(reply.proposal.geometry.targets.p1).toBeUndefined();
            else expect(reply.proposal.geometry.targets.p1?.ok).toBe(true);
            input = JSON.stringify({
              ...JSON.parse(layoutInput),
              renderEnv: { revisionPreview: proposalRevisionPreview(reply.proposal.result.snapshot) },
            });
            relayout = w.send({
              type: 'sync',
              snapshot: { ...snapshot, fonts: [], layoutInput: input, layoutRevision: 2 },
              extras: '', layoutExtras: '{}', provisionalPages: 1, displayWindow: [0, 1],
              expectedFrameEpoch: 1, paintCaret: false,
            }).then((frame) => {
              order.push('prefix');
              return frame;
            });
            return reply;
          });
        });
        const superseded = await complete();
        expect(superseded.ok && superseded.frame).toBeUndefined();
        await operation!;
        const prefix = await relayout!;
        expect(prefix.ok && prefix.layoutProvisional).toBe(true);
        expect(order).toEqual([kind, 'prefix']);
        expect(calls.filter((call) => call === 'prefix:1')).toHaveLength(2);
        if (!booted.ok || !booted.frame || !prefix.ok || !prefix.frame) throw new Error('frame missing');
        const visible = applyFrameDeltaOwned(
          applyFrameDeltaOwned(null, decodeFrameDelta(new Uint8Array(booted.frame))),
          decodeFrameDelta(new Uint8Array(prefix.frame))
        );
        const pageText = visible.displayList.pages.flatMap((page) => page.primitives.map((primitive) =>
          primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : ''
        )).join('');
        expect(pageText).toContain(kind === 'propose' ? 'Changed 1' : 'Paragraph 1');
        let framed: Promise<ResidentEngineWorkerResponse> | undefined;
        onResume.push(() => {
          framed = w.send({ type: 'buildFrame', extras: '{}', expectedFrameEpoch: 2, paintCaret: false })
            .then((frame) => {
              order.push('frame');
              return frame;
            });
        });
        const completed = await complete().then((reply) => {
          order.push('complete');
          return reply;
        });
        expect((await framed!).ok).toBe(true);
        expect(order).toEqual([kind, 'prefix', 'frame', 'complete']);
        await expectFullLayout(completed, input);
      } finally {
        engine.destroy();
      }
    }
  );
});

describe('resident worker opening', () => {
  test.each([undefined, 256])('forwards the optional preview paragraph budget %s', async (paragraphBudget) => {
    await preloadEditWasm(new Uint8Array(readFileSync(resolve(
      import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
    ))));
    const bytes = syntheticDocx('plain', 44, 17, {
      tableDense: true,
      blocks: 40,
      trailingShortParagraphs: 170,
    });
    const engine = await createResidentEngineSession();
    const w = worker();
    const forwarded: Array<number | undefined> = [];
    Object.assign(w.harness.session, engine, {
      openDocxPreview: (source: Uint8Array, blocks: number, budget?: number) => {
        forwarded.push(budget);
        return engine.openDocxPreview(source, blocks, budget);
      },
    });
    try {
      const reply = await w.send({
        type: 'open',
        bytes: bytes.buffer as ArrayBuffer,
        previewBlocks: 200,
        ...(paragraphBudget === undefined ? {} : { previewParagraphBudget: paragraphBudget }),
      });
      expect(reply.ok).toBe(true);
      expect(reply.ok && reply.hostJson).toBeDefined();
      expect(forwarded).toEqual([paragraphBudget]);
      const expected = await createResidentEngineSession();
      try {
        const hostJson = expected.openDocxPreview(bytes, 200, paragraphBudget);
        expect(reply.ok && reply.hostJson).toBe(hostJson ?? undefined);
        expect(engine.paragraphIdentities().paragraphs.length).toBe(
          expected.paragraphIdentities().paragraphs.length
        );
        const blockCount = await createResidentEngineSession();
        try {
          blockCount.openDocxPreview(bytes, 200);
          const count = blockCount.paragraphIdentities().paragraphs.length;
          if (paragraphBudget === undefined) {
            expect(engine.paragraphIdentities().paragraphs.length).toBe(count);
          } else {
            expect(engine.paragraphIdentities().paragraphs.length).toBeLessThan(count);
          }
        } finally {
          blockCount.destroy();
        }
      } finally {
        expected.destroy();
      }
    } finally {
      engine.destroy();
    }
  });

  const provisional = '{"layout":{"pages":[1]},"notesConverged":true,"provisional":true}';
  const full = '{"layout":{"pages":[1,2]},"notesConverged":true}';
  const snapshot = {
    clientId: 1,
    state: new Uint8Array([5]),
    fontsRevision: 1,
    fonts: [new Uint8Array([9])],
    renderInputs: [],
    measureInputs: [],
    layoutInput: '{"request":1}',
    layoutWithRegions: true,
    layoutRevision: 1,
    selection: null,
  };

  function openingWorker() {
    const w = worker();
    const calls: string[] = [];
    let epoch = 0;
    Object.assign(w.harness.session, {
      openDocx: (bytes: Uint8Array, digest?: string, generation?: string) => {
        calls.push(`open:${bytes.join(',')}:${digest}:${generation}`);
        return '{"host":1}';
      },
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
      layoutFontRequirementsJson: (input: string) => {
        calls.push(`requirements:${input}`);
        return '[{"key":"a"}]';
      },
      loadState: () => calls.push('loadState'),
      registerFont: () => {
        calls.push('font');
        return 1;
      },
      layoutDocumentWithRegionsPrefixRetainedJson: (input: string, pages: number) => {
        calls.push(`prefix:${input}:${pages}`);
        return provisional;
      },
      layoutDocumentWithRegionsRetainedJson: (input: string) => {
        calls.push(`layout:${input}`);
        return full;
      },
      beginRegionLayout: (input: string) => {
        calls.push(`begin:${input}`);
        return { measuredBlocks: 0, bodyBlocks: 2 };
      },
      resumeRegionLayout: (blocks: number) => {
        calls.push(`resume:${blocks}`);
        return { measuredBlocks: 2, bodyBlocks: 2, layoutJson: full };
      },
      residentCaretSnapshot: () => ({ frameEpoch: epoch, caretRect: null }),
      buildDisplayListFrame: (_input: string, expectedFrameEpoch: number) => {
        calls.push(`frame:${expectedFrameEpoch}`);
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
      encodeState: () => {
        calls.push('state');
        return new Uint8Array([7, 8]);
      },
      proposalEngine: { version: () => 'opened' },
    });
    return { w, calls };
  }

  async function peerReply(w: ReturnType<typeof worker>, request: ResidentEngineWorkerRequestWithoutId) {
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      return await Promise.race([
        w.send(request),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => reject(new Error(`No ${request.type} reply`)), 1000);
        }),
      ]);
    } finally {
      clearTimeout(timer);
    }
  }

  test('encodeState captures peer metadata, state, version and proposals in one synchronous handler', async () => {
    const { w } = openingWorker();
    await peerReply(w, { type: 'open', bytes: new Uint8Array([1]).buffer });
    let revision = 7;
    const captures: string[] = [];
    Object.assign(w.harness.session, {
      encodeState: () => {
        captures.push(`state:${revision}`);
        queueMicrotask(() => { revision += 1; });
        return new Uint8Array([0, revision, 0]).subarray(1, 2);
      },
      encodePeerMetadata: () => {
        captures.push(`metadata:${revision}`);
        return new Uint8Array([0, revision, 0]).subarray(1, 2);
      },
      proposalEngine: { version: () => {
        captures.push(`version:${revision}`);
        return `v${revision}`;
      } },
    });
    const reply = await peerReply(w, { type: 'encodeState', peerMetadata: true });
    if (!reply.ok || !reply.state || !reply.peerMetadata) throw new Error('Expected peer snapshot');
    expect(captures).toEqual(['state:7', 'metadata:7', 'version:7']);
    expect(new Uint8Array(reply.state)).toEqual(new Uint8Array([7]));
    expect(new Uint8Array(reply.peerMetadata)).toEqual(new Uint8Array([7]));
    expect(reply.version).toBe('v7');
    expect(reply.proposals).toEqual({ previewVersion: 0, entries: [] });
    expect(w.transfers.get(reply.id)).toEqual([reply.state, reply.peerMetadata]);
  });

  test('state-only encodeState keeps its reply and transfer unchanged', async () => {
    const { w } = openingWorker();
    await peerReply(w, { type: 'open', bytes: new Uint8Array([1]).buffer });
    const metadata = mock(() => new Uint8Array([9]));
    Object.assign(w.harness.session, { encodePeerMetadata: metadata });
    const reply = await peerReply(w, { type: 'encodeState' });
    if (!reply.ok || !reply.state) throw new Error('Expected state');
    expect(reply).toEqual({
      id: reply.id, ok: true, state: new Uint8Array([7, 8]).buffer,
      version: 'opened', proposals: { previewVersion: 0, entries: [] },
      memory: w.harness.memories,
    });
    expect(metadata).not.toHaveBeenCalled();
    expect(w.transfers.get(reply.id)).toEqual([reply.state]);
  });

  test.each(['missing', 'rejected'] as const)('encodeState preserves state when metadata is %s', async (kind) => {
    const { w } = openingWorker();
    await peerReply(w, { type: 'open', bytes: new Uint8Array([1]).buffer });
    if (kind === 'rejected') Object.assign(w.harness.session, {
      encodePeerMetadata: () => { throw new PeerMetadataError('unopened', 'No peer source'); },
    });
    const reply = await peerReply(w, { type: 'encodeState', peerMetadata: true });
    if (!reply.ok || !reply.state) throw new Error('Expected valid state despite metadata absence');
    expect(new Uint8Array(reply.state)).toEqual(new Uint8Array([7, 8]));
    expect(reply.peerMetadata).toBeUndefined();
    expect(reply.peerMetadataReason).toContain(kind === 'missing' ? 'missing-capability' : 'unopened');
    expect(reply.version).toBe('opened');
    expect(w.transfers.get(reply.id)).toEqual([reply.state]);
  });

  test('retains the opened source, keeps save history across an opened bootstrap and resets it on destroy/open', async () => {
    const { w } = openingWorker();
    const host = { package: { document: { content: [] } } };
    let source = new Uint8Array([1, 2, 3]).buffer;
    const records: ResidentSaveRecord[] = [];
    const save: ResidentEngineSession['save'] = async (bytes, hostJson, metadata, comments, record) => {
      expect(bytes.buffer).toBe(source);
      expect(hostJson).toBe('{"host":1}');
      expect(metadata).toBe(host);
      expect(comments).toEqual([]);
      if (records.at(-1) !== record) {
        expect(record).toEqual({ full: false });
        records.push(record);
      } else {
        expect(record.full).toBe(true);
        expect(record.saved).toBeDefined();
        expect(record.base).toBe(host);
      }
      const saved = new Uint8Array([4, 5, 6]).buffer;
      record.full = true;
      record.saved = saved;
      record.base = host as unknown as NonNullable<ResidentSaveRecord['base']>;
      return saved;
    };
    Object.assign(w.harness.session, { save });
    expect((await w.send({ type: 'open', bytes: source })).ok).toBe(true);
    const first = await w.send({ type: 'save', host, comments: [] });
    expect(first.ok && first.saved).not.toBe(records[0]!.saved);
    expect(first.ok && [...new Uint8Array(first.saved!)]).toEqual([4, 5, 6]);
    expect((await w.send({ type: 'save', host, comments: [] })).ok).toBe(true);
    expect((await w.send({
      type: 'bootstrap', opened: true, snapshot, extras: '', layoutExtras: '{}', expectedFrameEpoch: 0,
    })).ok).toBe(true);
    expect((await w.send({ type: 'save', host, comments: [] })).ok).toBe(true);
    expect(records).toHaveLength(1);
    w.scope.onmessage({ data: { id: 999, type: 'destroy' } });
    source = new Uint8Array([7, 8]).buffer;
    expect((await w.send({ type: 'open', bytes: source })).ok).toBe(true);
    expect((await w.send({ type: 'save', host, comments: [] })).ok).toBe(true);
    expect(records).toHaveLength(2);
    expect((await w.bootstrap()).ok).toBe(true);
    expect(await w.send({ type: 'save', host, comments: [] })).toMatchObject({
      ok: false, code: 'save-unavailable',
    });
  });

  test('save returns exactly the peer diff only when given a state vector', async () => {
    const { w } = openingWorker();
    const diff = new Uint8Array([7, 8, 9]);
    const encodeStateAsUpdate = mock((_stateVector: Uint8Array) => diff);
    Object.assign(w.harness.session, {
      save: async () => new ArrayBuffer(4),
      encodeStateAsUpdate,
    });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1) })).ok).toBe(true);
    const stateVector = new Uint8Array([1, 2, 3]);
    const saved = await w.send({ type: 'save', comments: [], stateVector });
    expect(saved.ok && saved.updates).toEqual([diff.buffer]);
    expect(encodeStateAsUpdate).toHaveBeenCalledTimes(1);
    expect(encodeStateAsUpdate).toHaveBeenCalledWith(stateVector);
    encodeStateAsUpdate.mockClear();
    expect(await w.send({ type: 'save', comments: [] })).toMatchObject({ ok: true, updates: [] });
    expect(encodeStateAsUpdate).not.toHaveBeenCalled();
  });

  test('syncUpdate applies in mutation order and acknowledges its version, vector and repairs', async () => {
    const { w } = openingWorker();
    const calls: string[] = [];
    let version = 'opened';
    const vector = new Uint8Array([1, 2]);
    const repair = new Uint8Array([3, 4]);
    Object.assign(w.harness.session, {
      applyUpdate: (update: Uint8Array) => { version = String(update[0]); calls.push(`apply:${version}`); },
      encodeStateAsUpdate: (captured: Uint8Array) => {
        expect(captured).toEqual(vector);
        calls.push(`repair:${version}`);
        return repair;
      },
      encodeStateVector: () => vector,
      proposalEngine: { version: () => version },
      save: async () => { calls.push(`save:${version}`); return new ArrayBuffer(0); },
    });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1) })).ok).toBe(true);
    void w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null });
    const acknowledged = w.send({ type: 'syncUpdate', update: new Uint8Array([2]), stateVector: vector });
    void w.send({ type: 'applyUpdate', update: new Uint8Array([3]), selection: null });
    const saved = w.send({ type: 'save', comments: [] });
    expect(await acknowledged).toMatchObject({ ok: true, version: '2', stateVector: vector.buffer, repair: repair.buffer });
    expect((await saved).ok).toBe(true);
    expect(calls).toEqual(['apply:1', 'apply:2', 'repair:2', 'apply:3', 'save:3']);
  });

  test('syncUpdate acknowledges an empty diff and a deletion with an unchanged state vector', async () => {
    const { w } = openingWorker();
    const vector = new Uint8Array([1, 2]);
    let version = 'opened';
    const applyUpdate = mock((update: Uint8Array) => { if (update.length > 0) version = 'deleted'; });
    const encodeStateAsUpdate = mock(() => new Uint8Array([0, 0]));
    Object.assign(w.harness.session, {
      applyUpdate, encodeStateAsUpdate,
      encodeStateVector: () => vector,
      proposalEngine: { version: () => version },
    });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1) })).ok).toBe(true);
    expect(await w.send({ type: 'syncUpdate', update: new Uint8Array(), stateVector: vector })).toMatchObject({
      ok: true, version: 'opened', stateVector: vector.buffer,
    });
    expect(await w.send({ type: 'syncUpdate', update: new Uint8Array([9]), stateVector: vector })).toMatchObject({
      ok: true, version: 'deleted', stateVector: vector.buffer,
    });
    expect(applyUpdate).toHaveBeenCalledTimes(1);
    expect(encodeStateAsUpdate).toHaveBeenCalledTimes(2);
  });

  test('syncUpdate reports a worker trap as a terminal failure', async () => {
    const { w } = openingWorker();
    Object.assign(w.harness.session, { applyUpdate: () => { throw new WebAssembly.RuntimeError('unreachable'); } });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1) })).ok).toBe(true);
    expect(await w.send({ type: 'syncUpdate', update: new Uint8Array([1]), stateVector: new Uint8Array([0]) })).toMatchObject({
      ok: false, terminal: true,
    });
  });

  test('save forwards the peer story set, including an empty set, without reading revisions', async () => {
    const { w } = openingWorker();
    const save = mock(async (..._args: Parameters<ResidentEngineSession['save']>) => new ArrayBuffer(4));
    const storiesChangedSince = mock(() => { throw new Error('save must not read revisions'); });
    Object.assign(w.harness.session, { save, storiesChangedSince });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1) })).ok).toBe(true);
    for (const stories of [['hf:rIdH1'], []]) {
      expect((await w.send({ type: 'save', comments: [], stories })).ok).toBe(true);
      expect(save.mock.calls.at(-1)![5]).toEqual(stories);
    }
    expect((await w.send({ type: 'save', comments: [] })).ok).toBe(true);
    expect(save.mock.calls.at(-1)![5]).toBeUndefined();
    expect(storiesChangedSince).not.toHaveBeenCalled();
  });

  test('a preview rejects save as still opening', async () => {
    const { w } = openingWorker();
    Object.assign(w.harness.session, { openDocxPreview: () => '{"host":"preview"}' });
    expect((await w.send({ type: 'open', bytes: new ArrayBuffer(1), previewBlocks: 1 })).ok).toBe(true);
    expect(await w.send({ type: 'save', comments: [] })).toMatchObject({
      ok: false, code: 'save-unavailable', error: 'Resident engine worker is still opening',
    });
  });

  test('opens a package, lays it out without a state to load, and hands its state over', async () => {
    const { w, calls } = openingWorker();
    const opened = await w.send({
      type: 'open',
      bytes: new Uint8Array([1, 2, 3]).buffer,
      digest: 'abc',
      generation: 'opening',
      heapLimitBytes: 1024,
    });
    expect(opened.ok && opened.hostJson).toBe('{"host":1}');
    expect(opened.ok && opened.stateVector).toBeDefined();
    expect(w.harness.directBatches).toEqual([true]);
    expect(opened.memory).toEqual(w.harness.memories);
    const requirements = await w.send({ type: 'fontRequirements', layoutInput: '{"request":1}' });
    expect(requirements.ok && requirements.requirementsJson).toBe('[{"key":"a"}]');
    expect(requirements.memory).toEqual(w.harness.memories);

    const framed = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
      provisionalPages: 3,
    });
    expect(framed.ok && framed.layoutJson).toBe(provisional);
    expect(framed.ok && framed.layoutProvisional).toBe(true);
    expect((w.harness as { heapLimits?: unknown[] }).heapLimits).toEqual([1024]);
    expect(calls).toEqual([
      'open:1,2,3:abc:opening',
      'requirements:{"request":1}',
      'font',
      'prefix:{"request":1}:3',
      'frame:0',
    ]);

    const state = await w.send({ type: 'encodeState' });
    expect(state.ok && [...new Uint8Array(state.state!)]).toEqual([7, 8]);
    expect(state.ok && state.version).toBe('opened');
    expect(state.ok && state.proposals).toEqual({ previewVersion: 0, entries: [] });
    expect(state.memory).toEqual(w.harness.memories);
    expect(calls[calls.length - 1]).toBe('state');

    const completed = await w.send({
      type: 'completeLayout',
      expectedFrameEpoch: 1,
      paintCaret: false,
      sliceBlocks: 2,
    });
    expect(completed.ok && completed.layoutJson).toBe(full);
    expect(completed.ok && completed.layoutProvisional).toBeUndefined();
    expect(calls.slice(-3)).toEqual(['begin:{"request":1}', 'resume:2', 'frame:1']);
  });

  test('answers a repeated font requirements request from the last answer until the document updates', async () => {
    const { w, calls } = openingWorker();
    const listeners = new Set<() => void>();
    Object.assign(w.harness.session, {
      onUpdate: (listener: () => void) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
    });
    expect((await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer })).ok).toBe(true);
    const ask = async (layoutInput: string) => {
      const answer = await w.send({ type: 'fontRequirements', layoutInput });
      expect(answer.ok && answer.requirementsJson).toBe('[{"key":"a"}]');
    };
    for (const input of ['{"request":1}', '{"request":1}', '{"request":2}', '{"request":2}']) {
      await ask(input);
    }
    for (const listener of [...listeners]) listener();
    expect(listeners.size).toBe(0);
    await ask('{"request":2}');
    await ask('{"request":2}');
    expect(calls.filter((call) => call.startsWith('requirements:'))).toEqual([
      'requirements:{"request":1}',
      'requirements:{"request":2}',
      'requirements:{"request":2}',
    ]);
  });

  test('opens a preview, lays it out, and replaces it with the whole package', async () => {
    const { w, calls } = openingWorker();
    Object.assign(w.harness.session, {
      openDocxPreview: (bytes: Uint8Array, blocks: number) => {
        calls.push(`preview:${bytes.join(',')}:${blocks}`);
        return '{"host":"preview"}';
      },
      destroy: () => calls.push('destroy'),
    });
    const bootstrap = {
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
      provisionalPages: 3,
    } as const;

    const preview = await w.send({ type: 'open', bytes: new Uint8Array([1, 2]).buffer, previewBlocks: 200 });
    expect(preview.ok && preview.hostJson).toBe('{"host":"preview"}');
    expect(preview.ok && preview.stateVector).toBeDefined();
    expect(w.harness.directBatches).toEqual([]);
    const framed = await w.send(bootstrap);
    expect(framed.ok && framed.layoutJson).toBe(provisional);
    expect(framed.ok && framed.documentPreview).toBe(true);
    expect(framed.ok && framed.documentAsOpened).toBeUndefined();
    // A second preview never replaces the first.
    const again = await w.send({ type: 'open', bytes: new Uint8Array([1, 2]).buffer, previewBlocks: 200 });
    expect(again.ok).toBe(false);

    const opened = await w.send({ type: 'open', bytes: new Uint8Array([3]).buffer, digest: 'abc' });
    expect(opened.ok && opened.hostJson).toBe('{"host":1}');
    expect(w.harness.directBatches).toEqual([true]);
    const full = await w.send({ ...bootstrap, expectedFrameEpoch: 1 });
    expect(full.ok && full.layoutJson).toBe(provisional);
    expect(full.ok && full.documentPreview).toBeUndefined();
    expect(full.ok && full.documentAsOpened).toBe(true);
    expect(calls).toEqual([
      'preview:1,2:200',
      'font',
      'prefix:{"request":1}:3',
      'frame:0',
      'destroy',
      'open:3:abc:undefined',
      'font',
      'prefix:{"request":1}:3',
      'frame:1',
    ]);
    expect(w.harness.sessionsCreated).toBe(2);

    // The whole package is never replaced.
    const replaced = await w.send({ type: 'open', bytes: new Uint8Array([4]).buffer });
    expect(replaced.ok).toBe(false);
    expect(calls).not.toContain('open:4:undefined:undefined');

    // A package that cannot open as a preview opens nothing, and the whole package opens after.
    const refusing = openingWorker();
    Object.assign(refusing.w.harness.session, {
      openDocxPreview: () => null,
      destroy: () => refusing.calls.push('destroy'),
    });
    const refused = await refusing.w.send({ type: 'open', bytes: new Uint8Array([5]).buffer, previewBlocks: 200 });
    expect(refused.ok && refused.previewRefused).toBe(true);
    expect(refused.ok && refused.hostJson).toBeUndefined();
    const fallback = await refusing.w.send({ type: 'open', bytes: new Uint8Array([6]).buffer });
    expect(fallback.ok && fallback.hostJson).toBe('{"host":1}');
    expect(refusing.calls).toEqual(['destroy', 'open:6:undefined:undefined']);
  });

  test('a frame after the document version changes is no longer as opened', async () => {
    const { w } = openingWorker();
    let version = 'opened';
    w.harness.session.proposalEngine.version = () => version;
    expect((await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer })).ok).toBe(true);
    const first = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
    });
    expect(first.ok && first.documentVersion).toBe('opened');
    expect(first.ok && first.documentAsOpened).toBe(true);

    version = 'changed';
    const changed = await w.send({
      type: 'buildFrame',
      extras: '',
      expectedFrameEpoch: 1,
      paintCaret: false,
    });
    expect(changed.ok && changed.documentVersion).toBe('changed');
    expect(changed.ok && changed.documentAsOpened).toBeUndefined();
  });

  test('an update before the first frame of the whole document leaves that frame not as opened', async () => {
    const { w } = openingWorker();
    let version = 'opened';
    Object.assign(w.harness.session, {
      applyUpdate: () => { version = 'changed'; },
    });
    w.harness.session.proposalEngine.version = () => version;
    expect((await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer })).ok).toBe(true);
    void w.send({ type: 'applyUpdate', update: new Uint8Array([1]), selection: null });
    const first = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
    });
    expect(first.ok && first.documentVersion).toBe('changed');
    expect(first.ok && first.documentAsOpened).toBeUndefined();
  });

  test('a replica state synced before the first frame of the whole document leaves that frame not as opened', async () => {
    const { w } = openingWorker();
    let version = 'opened';
    Object.assign(w.harness.session, {
      loadState: () => { version = 'changed'; },
    });
    w.harness.session.proposalEngine.version = () => version;
    expect((await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer })).ok).toBe(true);
    const synced = await w.send({
      type: 'sync',
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
      paintCaret: false,
    });
    expect(synced.ok && synced.documentVersion).toBe('changed');
    expect(synced.ok && synced.documentAsOpened).toBeUndefined();
  });

  test('proposal requests between open and bootstrap leave the worker registry empty', async () => {
    const { w } = openingWorker();
    expect((await w.send({ type: 'open', bytes: new Uint8Array([4]).buffer })).ok).toBe(true);
    const operations: ResidentProposalOperation[] = [
      {
        kind: 'propose',
        request: {
          expectVersion: 'opened',
          proposals: [
            {
              id: 'too-early',
              paragraph: { kind: 'session', sessionId: 'opened', story: 'body', paraId: 'p1' },
              suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
              op: 'insertText',
              at: 'end',
              text: '!',
            },
          ],
        },
      },
      { kind: 'snapshot' },
    ];
    for (const operation of operations) {
      const refused = await w.send({ type: 'proposal', operation });
      expect(refused).toMatchObject({
        ok: false,
        error: 'Resident engine worker has not laid out its document',
      });
      expect(!refused.ok && refused.terminal).toBeUndefined();
    }
    Object.assign(w.harness.session, {
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
      geometryReader: { version: () => 'opened' },
    });
    const bootstrapped = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      expectedFrameEpoch: 0,
    });
    expect(bootstrapped.ok).toBe(true);
    const empty = await w.send({ type: 'proposal', operation: { kind: 'snapshot' } });
    expect(empty.ok && empty.proposal?.mirror).toEqual({
      version: 'opened',
      proposals: { previewVersion: 0, entries: [] },
    });
  });

  test('a snapshot bootstrap lays out the same provisional prefix after loading state', async () => {
    const { w, calls } = openingWorker();
    const framed = await w.send({
      type: 'bootstrap',
      snapshot,
      extras: '',
      layoutExtras: '{}',
      expectedFrameEpoch: 0,
      provisionalPages: 3,
    });
    expect(framed.ok && framed.layoutJson).toBe(provisional);
    expect(framed.ok && framed.layoutProvisional).toBe(true);
    expect(framed.ok && framed.documentAsOpened).toBeUndefined();
    expect(calls).toEqual(['loadState', 'font', 'prefix:{"request":1}:3', 'frame:0']);
  });

  for (const count of [0, 2]) {
    test(`reports ${count} revisions from the opened session`, async () => {
      const { w } = openingWorker();
      w.harness.session.revisionCount = () => count;
      await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer });
      const response = await w.send({ type: 'revisionCount' });
      expect(response.ok && response.revisionCount).toBe(count);
      expect(response.memory).toEqual(w.harness.memories);
    });
  }

  test('revision count requires an initialized session', async () => {
    const { w } = openingWorker();
    const response = await w.send({ type: 'revisionCount' });
    expect(response).toMatchObject({
      ok: false,
      error: 'Resident engine worker is not initialized',
    });
  });

  for (const type of ['open', 'fontRequirements', 'encodeState', 'revisionCount'] as const) {
    test(`${type} attributes an OOM trap and refuses queued requests`, async () => {
      const { w, calls } = openingWorker();
      if (type !== 'open') {
        await w.send({ type: 'open', bytes: new Uint8Array([4]).buffer });
      }
      calls.length = 0;
      w.harness.memories[0].failedAllocationBytes = 64;
      const method = {
        open: 'openDocx',
        fontRequirements: 'layoutFontRequirementsJson',
        encodeState: 'encodeState',
        revisionCount: 'revisionCount',
      }[type];
      Object.assign(w.harness.session, {
        [method]: () => {
          calls.push('trap');
          throw new WebAssembly.RuntimeError('unreachable');
        },
      });
      const failed = w.send(
        type === 'open'
          ? { type, bytes: new Uint8Array([4]).buffer }
          : type === 'fontRequirements'
            ? { type, layoutInput: snapshot.layoutInput }
            : { type }
      );
      const queued = w.send({ type: 'encodeState' });
      for (const reply of await Promise.all([failed, queued])) {
        expect(!reply.ok && reply.terminal && reply.outOfMemory).toBe(true);
        expect(!reply.ok && reply.error).toBe(
          'Resident engine worker ran out of memory allocating 64 bytes: unreachable'
        );
        expect(reply.memory).toEqual(w.harness.memories);
        expect(w.answered.filter((id) => id === reply.id)).toHaveLength(1);
      }
      expect(calls).toEqual(['trap']);
    });
  }

  test('a second open fails and keeps the document the first one opened', async () => {
    const { w, calls } = openingWorker();
    expect((await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer })).ok).toBe(true);
    const second = await w.send({ type: 'open', bytes: new Uint8Array([2]).buffer });
    expect(!second.ok && second.error).toBe('Resident engine worker already holds a document');
    expect(!second.ok && second.terminal).toBeFalsy();
    const framed = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      expectedFrameEpoch: 0,
    });
    expect(framed.ok).toBe(true);
    expect(calls.filter((call) => call.startsWith('open:'))).toEqual(['open:1:undefined:undefined']);
    expect((w.harness as { heapLimits?: unknown[] }).heapLimits).toHaveLength(1);
  });

  test('an opened bootstrap refuses a heap limit other than the one it opened under', async () => {
    const { w } = openingWorker();
    await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer, heapLimitBytes: 1024 });
    const bootstrap = (heapLimitBytes?: number) =>
      w.send({
        type: 'bootstrap',
        opened: true,
        snapshot,
        extras: '',
        expectedFrameEpoch: 0,
        ...(heapLimitBytes !== undefined ? { heapLimitBytes } : {}),
      });
    const refused = await bootstrap(2048);
    expect(!refused.ok && refused.error).toBe(
      'Resident engine worker opened its document under another heap limit'
    );
    expect((await bootstrap(1024)).ok).toBe(true);
  });

  test('a package that fails to open frees its session and leaves the worker able to open', async () => {
    const { w, calls } = openingWorker();
    const { openDocx } = w.harness.session as { openDocx?: unknown };
    Object.assign(w.harness.session, {
      openDocx: () => {
        throw new Error('not a package');
      },
      destroy: () => calls.push('destroy'),
    });
    const failed = await w.send({ type: 'open', bytes: new Uint8Array([1]).buffer });
    expect(!failed.ok && failed.error).toBe('not a package');
    expect(!failed.ok && failed.terminal).toBeFalsy();
    expect(calls).toEqual(['destroy']);
    Object.assign(w.harness.session, { openDocx });
    expect((await w.send({ type: 'open', bytes: new Uint8Array([2]).buffer })).ok).toBe(true);
  });

  test('a bootstrap of an opened document fails, and keeps the worker, when nothing was opened', async () => {
    const { w, calls } = openingWorker();
    const refused = await w.send({
      type: 'bootstrap',
      opened: true,
      snapshot,
      extras: '',
      expectedFrameEpoch: 0,
    });
    expect(refused.ok).toBe(false);
    expect(!refused.ok && refused.terminal).toBeFalsy();
    expect(calls).toEqual([]);
    const opened = await w.send({ type: 'open', bytes: new Uint8Array([4]).buffer });
    expect(opened.ok).toBe(true);
  });
});
