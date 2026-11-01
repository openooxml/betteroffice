import { beforeAll, expect, mock, test } from 'bun:test';
import { resolve } from 'node:path';

type Renderer = typeof import('./raster');
type Fetch = typeof globalThis.fetch;
let createRenderer: (fetch: Fetch) => Renderer;
const png = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10);

beforeAll(async () => {
  const entry = resolve(import.meta.dir, 'raster.ts');
  const result = await Bun.build({
    entrypoints: [import.meta.path],
    target: 'bun',
    format: 'esm',
    plugins: [{
      name: 'isolated-xlsx-raster',
      setup(build) {
        build.onLoad({ filter: /raster\.test\.ts$/ }, () => ({
          contents: `import * as raster from ${JSON.stringify(entry)}; testHarness.raster = raster;`,
          loader: 'js',
        }));
        build.onLoad({ filter: /\/wasm\/asset\.ts$/ }, () => ({
          contents: "export const rasterWasmAssetUrl = () => new URL('https://example.test/xlsx_raster_bg.wasm');",
          loader: 'js',
        }));
        build.onResolve({ filter: /\/generated\/raster\/xlsx_raster\.js$/ }, () => ({
          path: 'raster', namespace: 'test-raster',
        }));
        build.onLoad({ filter: /.*/, namespace: 'test-raster' }, () => ({
          contents: `
            export default async function init({ module_or_path }) {
              await (module_or_path instanceof Response ? module_or_path : await fetch(module_or_path)).arrayBuffer();
            }
            export const render_png = () => testHarness.png;
          `,
          loader: 'js',
        }));
      },
    }],
  });
  if (!result.success) throw new Error(result.logs.join('\n'));
  const load = new Function('testHarness', 'fetch', await result.outputs[0].text());
  createRenderer = (fetch) => {
    const harness = { raster: undefined as unknown as Renderer, png };
    load(harness, fetch);
    return harness.raster;
  };
});

test('shares a failed raster initialization and retries on the next export', async () => {
  let rejectFetch!: (error: Error) => void;
  const pending = new Promise<Response>((_, reject) => { rejectFetch = reject; });
  let fetched!: () => void;
  const started = new Promise<void>((resolve) => { fetched = resolve; });
  const requests: string[] = [];
  const fetch = mock((async (input) => {
    requests.push(String(input));
    fetched();
    return requests.length === 1 ? pending : new Response();
  }) as Fetch);
  const renderer = createRenderer(fetch as unknown as Fetch);
  const failing = Promise.allSettled([renderer.renderPng('first'), renderer.renderPng('concurrent')]);
  await started;
  expect(fetch).toHaveBeenCalledTimes(1);
  const error = new Error('Raster fetch failed');
  rejectFetch(error);
  expect(await failing).toEqual([
    { status: 'rejected', reason: error },
    { status: 'rejected', reason: error },
  ]);
  await expect(renderer.renderPng('retry')).resolves.toEqual(png);
  expect(fetch).toHaveBeenCalledTimes(2);
  await expect(renderer.renderPng('cached')).resolves.toEqual(png);
  expect(requests).toEqual(Array(2).fill('https://example.test/xlsx_raster_bg.wasm'));
});

function streamed(chunks: number, everyMs: number, stopAfter = chunks) {
  let sent = 0;
  let timer: ReturnType<typeof setInterval> | undefined;
  return new Response(new ReadableStream<Uint8Array>({
    start(controller) {
      timer = setInterval(() => {
        if (sent >= stopAfter) { clearInterval(timer); if (stopAfter >= chunks) controller.close(); return; }
        sent += 1;
        controller.enqueue(new Uint8Array(64));
      }, everyMs);
    },
    cancel() { clearInterval(timer); },
  }), { headers: { 'content-type': 'application/wasm' } });
}

async function rejection(call: Promise<unknown>) {
  return call.then(() => { throw new Error('Expected a rejected export'); }, (error: unknown) => error as Error & { code?: string });
}

test('a slow download that keeps sending bytes outlives the stall timeout and reports progress', async () => {
  const fetch = mock((async () => streamed(12, 30)) as unknown as Fetch);
  const renderer = createRenderer(fetch as unknown as Fetch);
  let progress = 0;
  const started = performance.now();
  await expect(renderer.renderPng('slow', { stallMs: 250, totalMs: 5_000, progress: () => { progress += 1; } }))
    .resolves.toEqual(png);
  expect(performance.now() - started).toBeGreaterThan(300);
  expect(progress).toBe(12);
  expect(fetch).toHaveBeenCalledTimes(1);
});

test('a download that stops sending bytes rejects at the stall timeout, aborts and retries on the next export', async () => {
  const signals: AbortSignal[] = [];
  const fetch = mock((async (_input: unknown, init?: RequestInit) => {
    signals.push(init!.signal!);
    return signals.length === 1 ? streamed(10, 5, 2) : new Response();
  }) as unknown as Fetch);
  const renderer = createRenderer(fetch as unknown as Fetch);
  const started = performance.now();
  const error = await rejection(renderer.renderPng('stalled', { stallMs: 40 }));
  const elapsed = performance.now() - started;
  expect(error.code).toBe('png-raster-timeout');
  expect(error.name).toBe('WorkbookPngRasterLoadError');
  expect(error.message).toBe('PNG export failed: no rasterizer bytes arrived for 40 ms');
  expect(elapsed).toBeGreaterThanOrEqual(40);
  expect(elapsed).toBeLessThan(1_000);
  expect(signals[0].aborted).toBe(true);
  await expect(renderer.renderPng('retry', { stallMs: 40 })).resolves.toEqual(png);
  expect(fetch).toHaveBeenCalledTimes(2);
});

test('a download whose response never starts rejects at the stall timeout', async () => {
  const fetch = mock((async () => new Promise<Response>(() => {})) as unknown as Fetch);
  const renderer = createRenderer(fetch as unknown as Fetch);
  const error = await rejection(renderer.renderPng('stalled', { stallMs: 30 }));
  expect(error.message).toBe('PNG export failed: no rasterizer bytes arrived for 30 ms');
});

test('the overall backstop rejects a download that trickles forever', async () => {
  const fetch = mock((async () => streamed(Number.MAX_SAFE_INTEGER, 5, 100)) as unknown as Fetch);
  const renderer = createRenderer(fetch as unknown as Fetch);
  const error = await rejection(renderer.renderPng('trickle', { stallMs: 50, totalMs: 120 }));
  expect(error.code).toBe('png-raster-timeout');
  expect(error.message).toBe('PNG export failed: the rasterizer did not load within 120 ms');
});
