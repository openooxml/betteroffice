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
            export default async function init({ module_or_path }) { await fetch(module_or_path); }
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
