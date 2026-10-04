import { test, expect, type Page } from 'playwright/test';
import type { ViewerPixelComparison } from './xlsx-worker-viewer-probe';

interface XlsxMainWasmProbe {
  instantiate: number;
  instantiateStreaming: number;
  compile: number;
  compileStreaming: number;
  modules: number;
  instances: number;
  fetches: string[];
  workers: { url: string; type: WorkerOptions['type'] }[];
}

declare global {
  interface Window { __xlsxMainWasmProbe: XlsxMainWasmProbe }
}

test.setTimeout(600_000);

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

async function open(page: Page, fixture: string, workerOnly = false) {
  await page.goto(`/xlsx-worker-viewer.html?fixture=${fixture}${workerOnly ? '&arm=worker' : ''}`);
  await page.waitForFunction(() => '__xlsxWorkerViewer' in window, undefined, { timeout: 120_000 });
  return page.evaluate(() => window.__xlsxWorkerViewer.ready);
}

function assertParity(result: ViewerPixelComparison, dpr: number, zoom: number, sheet: number) {
  const label = `sheet=${sheet} dpr=${dpr} zoom=${zoom} scroll=${result.scrollLeft},${result.scrollTop}`;
  expect(result.dpr).toBe(dpr);
  expect(result.zoom).toBe(zoom);
  expect(result.sheet).toBe(sheet);
  expect(result.width).toBeGreaterThan(0);
  expect(result.height).toBeGreaterThan(0);
  expect.soft(result.differingPixels, label).toBe(0);
  expect.soft(result.maxChannelDelta, label).toBe(0);
}

for (const dpr of [1, 2]) {
  test.describe(`xlsx worker viewer DPR ${dpr}`, () => {
    test.use({ deviceScaleFactor: dpr });
    for (const fixture of ['sample', 'charts']) {
      test(`${fixture}: editor canvas pixels match at each zoom, viewport and sheet`, async ({ page }) => {
        const { sheetCount } = await open(page, fixture);
        expect(sheetCount).toBe(fixture === 'sample' ? 3 : 1);
        await expect(page.locator('[data-arm]')).toHaveCount(2);
        for (const zoom of [1, 1.5]) {
          for (const position of ['origin', 'scrolled'] as const) {
            const result = await page.evaluate(({ zoom, position }) =>
              window.__xlsxWorkerViewer.show(zoom, position), { zoom, position });
            assertParity(result, dpr, zoom, 0);
            expect(result.scrollLeft).toBe(position === 'origin' ? 0 : 400 * zoom);
            expect(result.scrollTop).toBe(position === 'origin' ? 0 : 480 * zoom);
          }
        }

        if (fixture === 'sample') {
          for (const arm of ['in-thread', 'worker']) {
            const tabs = page.locator(`[data-arm="${arm}"] [data-testid="xlsx-sheet-tabs"]`);
            await tabs.getByRole('tab').nth(1).click();
            await expect(tabs.getByRole('tab', { selected: true })).toHaveText('Summary');
          }
          assertParity(await page.evaluate(() => window.__xlsxWorkerViewer.compareCurrent(1)), dpr, 1.5, 1);
          for (const zoom of [1, 1.5]) {
            for (const position of ['origin', 'scrolled'] as const) {
              const result = await page.evaluate(({ zoom, position }) =>
                window.__xlsxWorkerViewer.show(zoom, position), { zoom, position });
              assertParity(result, dpr, zoom, 1);
              expect(result.scrollLeft).toBe(position === 'origin' ? 0 : 400 * zoom);
              expect(result.scrollTop).toBe(position === 'origin' ? 0 : 480 * zoom);
            }
          }
        }
        expect(await page.evaluate(() => window.__xlsxWorkerViewer.errors)).toEqual([]);
      });
    }
  });
}

async function installMainWasmProbe(page: Page) {
  await page.addInitScript(() => {
    const probe: XlsxMainWasmProbe = {
      instantiate: 0, instantiateStreaming: 0, compile: 0, compileStreaming: 0,
      modules: 0, instances: 0, fetches: [], workers: [],
    };
    window.__xlsxMainWasmProbe = probe;
    const instantiate = WebAssembly.instantiate;
    WebAssembly.instantiate = function (...args: Parameters<typeof instantiate>) {
      probe.instantiate += 1;
      return Reflect.apply(instantiate, WebAssembly, args);
    } as typeof instantiate;
    const instantiateStreaming = WebAssembly.instantiateStreaming;
    WebAssembly.instantiateStreaming = function (...args: Parameters<typeof instantiateStreaming>) {
      probe.instantiateStreaming += 1;
      return Reflect.apply(instantiateStreaming, WebAssembly, args);
    } as typeof instantiateStreaming;
    const compile = WebAssembly.compile;
    WebAssembly.compile = function (...args: Parameters<typeof compile>) {
      probe.compile += 1;
      return Reflect.apply(compile, WebAssembly, args);
    } as typeof compile;
    const compileStreaming = WebAssembly.compileStreaming;
    WebAssembly.compileStreaming = function (...args: Parameters<typeof compileStreaming>) {
      probe.compileStreaming += 1;
      return Reflect.apply(compileStreaming, WebAssembly, args);
    };
    const NativeModule = WebAssembly.Module;
    WebAssembly.Module = new Proxy(NativeModule, {
      construct(target, args, newTarget) {
        probe.modules += 1;
        return Reflect.construct(target, args, newTarget);
      },
    });
    const NativeInstance = WebAssembly.Instance;
    WebAssembly.Instance = new Proxy(NativeInstance, {
      construct(target, args, newTarget) {
        probe.instances += 1;
        return Reflect.construct(target, args, newTarget);
      },
    });
    const nativeFetch = window.fetch;
    window.fetch = function (...args: Parameters<typeof nativeFetch>) {
      const input = args[0];
      const url = input instanceof Request ? input.url : String(input);
      if (url.includes('xlsx_wasm_bg.wasm')) probe.fetches.push(url);
      return Reflect.apply(nativeFetch, window, args);
    } as typeof nativeFetch;
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        probe.workers.push({ url: String(url), type: options?.type });
      }
    };
  });
}

async function assertWorkerOnly(page: Page, workers: string[]) {
  const result = await page.evaluate(() => ({
    wasm: window.__xlsxMainWasmProbe,
    resources: performance.getEntriesByType('resource').map((entry) => entry.name),
    errors: window.__xlsxWorkerViewer.errors,
  }));
  await expect(page.locator('[data-arm]')).toHaveCount(1);
  expect(result.wasm.instantiate).toBe(0);
  expect(result.wasm.instantiateStreaming).toBe(0);
  expect(result.wasm.compile).toBe(0);
  expect(result.wasm.compileStreaming).toBe(0);
  expect(result.wasm.modules).toBe(0);
  expect(result.wasm.instances).toBe(0);
  expect(result.wasm.fetches).toEqual([]);
  expect(result.resources.filter((url) => url.includes('xlsx_wasm_bg.wasm'))).toEqual([]);
  expect(workers).toHaveLength(1);
  expect(result.wasm.workers).toHaveLength(1);
  expect(result.wasm.workers[0].type).toBe('module');
  expect(result.wasm.workers[0].url).toContain('xlsxSessionWorker');
  expect(result.errors).toEqual([]);
}

for (const fixture of ['sample', 'charts']) {
  test(`xlsx ${fixture}: a fresh worker-only viewer keeps wasm off the main thread`, async ({ page }) => {
    const workers: string[] = [];
    page.on('worker', (worker) => workers.push(worker.url()));
    await installMainWasmProbe(page);
    await open(page, fixture, true);
    await assertWorkerOnly(page, workers);
  });
}

test('xlsx: worker onReady exposes async save and reveals selections before resolving', async ({ page }) => {
  const workers: string[] = [];
  page.on('worker', (worker) => workers.push(worker.url()));
  await installMainWasmProbe(page);
  await open(page, 'sample', true);
  const result = await page.evaluate(() => window.__xlsxWorkerViewer.contract());
  expect(result.handleIsNull).toBe(true);
  expect(result.syncSaveIsNull).toBe(true);
  expect(result.syncSelection).toBe(false);
  expect(result.asyncSelection).toBe(true);
  expect(result.scrollBefore).toEqual({ left: 0, top: 0 });
  expect(result.scrollAfter.left).toBeGreaterThan(0);
  expect(result.scrollAfter.top).toBeGreaterThan(0);
  expect(result.targetPaintedAtResolution).toBe(true);
  expect(result.targetVisibleAtResolution).toBe(true);
  await expect.poll(() => page.evaluate(() => {
    const scroll = document.querySelector<HTMLElement>('[data-arm="worker"] [data-testid="xlsx-scroll"]')!;
    const selection = scroll.querySelector('[data-testid="xlsx-selection"]')?.getBoundingClientRect();
    const viewport = scroll.getBoundingClientRect();
    return !!selection && selection.width > 0 && selection.height > 0 &&
      selection.left >= viewport.left && selection.top >= viewport.top &&
      selection.right <= viewport.left + scroll.clientWidth && selection.bottom <= viewport.top + scroll.clientHeight;
  }), { timeout: 120_000 }).toBe(true);
  expect(result.savedLength).toBeGreaterThan(4);
  expect(result.signature).toEqual([0x50, 0x4b, 0x03, 0x04]);
  await assertWorkerOnly(page, workers);
});
