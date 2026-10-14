import { test, expect, type Page } from 'playwright/test';
import type { PixelComparison } from './pptx-worker-viewer-probe';

interface MainWasmProbe {
  instantiate: number;
  instantiateStreaming: number;
  workers: { url: string; type: WorkerOptions['type'] }[];
}

declare global {
  interface Window { __mainWasmProbe: MainWasmProbe }
}

const failures = new WeakMap<Page, { console: string[]; page: string[] }>();
test.beforeEach(async ({ page }) => {
  const errors = { console: [] as string[], page: [] as string[] };
  failures.set(page, errors);
  page.on('console', (message) => {
    if (message.type() === 'error') errors.console.push(message.text());
  });
  page.on('pageerror', (error) => errors.page.push(error.message));
});
test.afterEach(async ({ page }) => {
  const errors = failures.get(page)!;
  expect.soft(errors.console, 'Console errors').toEqual([]);
  expect.soft(errors.page, 'Page errors').toEqual([]);
});

function assertParity(result: PixelComparison, images = false) {
  const label = `${result.thumbnail ? 'thumbnail' : 'main'} slide=${result.slide} ` +
    `dpr=${result.dpr} zoom=${result.zoom}`;
  expect.soft(result.differingPixels, label).toBe(0);
  expect.soft(result.maxChannelDelta, label).toBe(0);
  if (images) {
    expect.soft(result.localImageDraws, `${label} local media draws`).toBeGreaterThan(0);
    expect.soft(result.workerImageDraws, `${label} worker media draws`).toBeGreaterThan(0);
    expect.soft(result.localImagePixels, `${label} local image region pixels`).toBeGreaterThan(0);
    expect.soft(result.workerImagePixels, `${label} worker image region pixels`).toBeGreaterThan(0);
  }
}

for (const dpr of [1, 2]) {
  test.describe(`worker viewer DPR ${dpr}`, () => {
    test.use({ deviceScaleFactor: dpr });
    for (const [fixture, count] of [['demo', 3], ['tiff', 1]] as const) {
      test(`${fixture}: all slides, zooms, thumbnails and navigation match`, async ({ page }) => {
        const initialSlide = fixture === 'demo' ? 2 : 1;
        await page.goto(`/pptx-worker-viewer.html?fixture=${fixture}&initialSlide=${initialSlide}`);
        await page.waitForFunction(() => '__pptxWorkerViewer' in window);
        expect(await page.evaluate(() => window.__pptxWorkerViewer.ready))
          .toEqual({ slideCount: count, initialSlide });
        await expect(page.locator('[data-arm="worker"] aside [aria-current="page"]'))
          .toHaveAttribute('data-slide-index', String(initialSlide - 1));
        expect(await page.locator('[data-arm="in-thread"] aside button').evaluateAll((rows) =>
          rows.findIndex((row) => row.getAttribute('aria-current') === 'page'))).toBe(initialSlide - 1);

        for (const zoom of [1, 1.5]) {
          for (let slide = 1; slide <= count; slide += 1) {
            const result = await page.evaluate(({ slide, zoom }) =>
              window.__pptxWorkerViewer.show(slide, zoom), { slide, zoom });
            expect(result.dpr).toBe(dpr);
            assertParity(result, slide === 1);
          }
          const thumbnails = await page.evaluate(() => window.__pptxWorkerViewer.thumbnails());
          expect(thumbnails).toHaveLength(Math.min(3, count));
          for (const result of thumbnails) assertParity(result, result.slide === 1);
        }

        if (count > 1) {
          await page.locator('[data-arm="worker"] aside button').nth(1).click();
          assertParity(await page.evaluate(() => window.__pptxWorkerViewer.compareCurrent(2)));
          await page.locator('[data-arm="worker"] [tabindex="0"]').focus();
          await page.keyboard.press('PageDown');
          assertParity(await page.evaluate(() => window.__pptxWorkerViewer.compareCurrent(3)));
        }
        expect(await page.evaluate(() => window.__pptxWorkerViewer.errors)).toEqual([]);
      });
    }
  });
}

for (const fixture of ['demo', 'tiff']) {
  test(`${fixture}: a fresh worker-only page keeps wasm off the main thread`, async ({ page }) => {
    const workers: string[] = [];
    page.on('worker', (worker) => workers.push(worker.url()));
    await page.addInitScript(() => {
      const probe: MainWasmProbe = { instantiate: 0, instantiateStreaming: 0, workers: [] };
      window.__mainWasmProbe = probe;
      const instantiate = WebAssembly.instantiate;
      WebAssembly.instantiate = function (...args: Parameters<typeof instantiate>) {
        probe.instantiate += 1;
        return Reflect.apply(instantiate, WebAssembly, args);
      } as typeof instantiate;
      const streaming = WebAssembly.instantiateStreaming;
      WebAssembly.instantiateStreaming = function (...args: Parameters<typeof streaming>) {
        probe.instantiateStreaming += 1;
        return Reflect.apply(streaming, WebAssembly, args);
      } as typeof streaming;
      const NativeWorker = window.Worker;
      window.Worker = class extends NativeWorker {
        constructor(url: string | URL, options?: WorkerOptions) {
          super(url, options);
          probe.workers.push({ url: String(url), type: options?.type });
        }
      };
    });
    await page.goto(`/pptx-worker-viewer.html?arm=worker&fixture=${fixture}`);
    await page.waitForFunction(() => '__pptxWorkerViewer' in window);
    await page.evaluate(() => window.__pptxWorkerViewer.ready);
    const result = await page.evaluate(async () => ({
      wasm: window.__mainWasmProbe,
      resources: performance.getEntriesByType('resource').map((entry) => entry.name),
      image: await window.__pptxWorkerViewer.workerImage(),
      errors: window.__pptxWorkerViewer.errors,
    }));
    await expect(page.locator('[data-arm]')).toHaveCount(1);
    expect(result.wasm.instantiate).toBe(0);
    expect(result.wasm.instantiateStreaming).toBe(0);
    expect(result.resources.filter((url) => url.includes('pptx_wasm_bg.wasm'))).toEqual([]);
    expect(workers).toHaveLength(1);
    expect(result.wasm.workers).toHaveLength(1);
    expect(result.wasm.workers[0].type).toBe('module');
    expect(result.wasm.workers[0].url).toContain('pptxSessionWorker');
    expect(result.image.draws).toBeGreaterThan(0);
    expect(result.image.pixels).toBeGreaterThan(0);
    expect(result.errors).toEqual([]);
  });
}
