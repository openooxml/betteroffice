import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

type Loader = typeof import('./loader');
type Fetch = typeof globalThis.fetch;

let createLoader: (fetch: Fetch) => Loader;
let createNativeLoader: (fetch: Fetch) => Loader;
const viewport = { x: 0, y: 0, width: 320, height: 240 };
const fixture = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')));
const unchanged = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../test-fixtures/no-edit-save.xlsx')));
const generated = pathToFileURL(resolve(import.meta.dir, 'generated/')).href + '/';

async function loaderFactory(native = false): Promise<(fetch: Fetch) => Loader> {
  const entry = resolve(import.meta.dir, 'loader.ts');
  const glue = resolve(import.meta.dir, 'generated/xlsx_wasm.js');
  const result = await Bun.build({
    entrypoints: [import.meta.path],
    target: 'bun',
    format: 'esm',
    plugins: [{
      name: 'isolated-xlsx-loader',
      setup(build) {
        build.onLoad({ filter: /loader\.png\.test\.ts$/ }, () => ({
          contents: `import * as loader from ${JSON.stringify(entry)}; testHarness.loader = loader;`,
          loader: 'js',
        }));
        build.onLoad({ filter: /\/wasm\/engine\.ts$/ }, () => ({
          contents: `
            export * from ${JSON.stringify(glue)};
            export const engineKind = ${JSON.stringify(native ? 'native' : 'wasm')};
            ${native
              ? 'export default async function init() { return { memory: undefined }; }'
              : `export { default } from ${JSON.stringify(glue)};`}
          `,
          loader: 'js',
        }));
      },
    }],
  });
  if (!result.success) throw new Error(result.logs.join('\n'));
  const code = (await result.outputs[0].text()).split('import.meta.url').join(JSON.stringify(pathToFileURL(entry).href));
  const load = new Function('testHarness', 'fetch', code);
  return (fetch) => {
    const harness = { loader: undefined as unknown as Loader };
    load(harness, fetch);
    return harness.loader;
  };
}

function assetFetch(requests: string[]): Fetch {
  return (async (input) => {
    const url = input instanceof Request ? input.url : String(input);
    requests.push(url);
    return new Response(readFileSync(new URL(url)), {
      headers: { 'Content-Type': 'application/wasm' },
    });
  }) as Fetch;
}

beforeAll(async () => {
  createLoader = await loaderFactory();
  createNativeLoader = await loaderFactory(true);
});

describe('wasm PNG initialization', () => {
  test.each([undefined, {}, { png: true }])('keeps the default full asset with %j', async (options) => {
    const requests: string[] = [];
    const loader = createLoader(assetFetch(requests));
    await loader.initWasm(undefined, options);
    expect(requests).toEqual([new URL('xlsx_wasm_bg.wasm', generated).href]);
    expect(loader.isPngExportAvailable()).toBe(true);
    const handle = loader.openWorkbook(fixture);
    try {
      expect(Array.from(handle.renderPng(viewport).subarray(0, 8)))
        .toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    } finally { handle.dispose(); }
  });

  test('loads the lean asset and preserves open, edit, display lists and saved bytes', async () => {
    const requests: string[] = [];
    const lean = createLoader(assetFetch(requests));
    const full = createLoader(assetFetch([]));
    await Promise.all([lean.initWasm(undefined, { png: false }), full.initWasm()]);
    expect(requests).toEqual([new URL('xlsx_wasm_lean_bg.wasm', generated).href]);
    expect(lean.isWasmAvailable()).toBe(true);
    expect(lean.isPngExportAvailable()).toBe(false);
    expect(full.isPngExportAvailable()).toBe(true);
    for (const loader of [lean, full]) {
      const handle = loader.openWorkbook(unchanged);
      try {
        handle.displayList(viewport);
        expect(handle.save()).toEqual(unchanged);
      } finally { handle.dispose(); }
    }
    const options = { calculation: { nowSerial: 45000, randSeed: 1 } };
    const leanHandle = lean.openWorkbook(fixture, options);
    const fullHandle = full.openWorkbook(fixture, options);
    try {
      expect(leanHandle.sheetInfo()).toEqual(fullHandle.sheetInfo());
      expect(leanHandle.displayList(viewport)).toEqual(fullHandle.displayList(viewport));
      expect(() => leanHandle.renderPng(viewport)).toThrow('png export not in this build');
      expect(() => leanHandle.renderRangePng({ range: 'A1:B2' })).toThrow('png export not in this build');
      for (const handle of [leanHandle, fullHandle]) {
        expect(handle.editCell(0, 2, 1, '1234').applied).toBe(true);
        expect(handle.cell(0, 2, 1).input).toBe('1234');
      }
      expect(leanHandle.displayList(viewport)).toEqual(fullHandle.displayList(viewport));
      expect(leanHandle.save()).toEqual(fullHandle.save());
      const reopened = lean.openWorkbook(leanHandle.save());
      try { expect(reopened.cell(0, 2, 1).input).toBe('1234'); }
      finally { reopened.dispose(); }
    } finally {
      leanHandle.dispose();
      fullHandle.dispose();
    }
  });

  test('honors explicit full bytes with png disabled and detects an explicit lean module', async () => {
    const requests: string[] = [];
    const full = createLoader(assetFetch(requests));
    const lean = createLoader(assetFetch(requests));
    const fullBytes = readFileSync(new URL('xlsx_wasm_bg.wasm', generated));
    const leanModule = await WebAssembly.compile(readFileSync(new URL('xlsx_wasm_lean_bg.wasm', generated)));
    await full.initWasm(fullBytes, { png: false });
    await lean.initWasm(Promise.resolve(leanModule));
    expect(requests).toEqual([]);
    expect(full.isPngExportAvailable()).toBe(true);
    expect(lean.isPngExportAvailable()).toBe(false);
    const handle = lean.openWorkbook(fixture);
    try { expect(() => handle.renderPng(viewport)).toThrow('png export not in this build'); }
    finally { handle.dispose(); }
  });

  test('shares the first initialization when concurrent callers request different assets', async () => {
    const requests: string[] = [];
    const loader = createLoader(assetFetch(requests));
    const first = loader.initWasm(undefined, { png: false });
    expect(loader.initWasm()).toBe(first);
    await first;
    await loader.initWasm(undefined, { png: true });
    expect(requests).toEqual([new URL('xlsx_wasm_lean_bg.wasm', generated).href]);
    expect(loader.isPngExportAvailable()).toBe(false);
  });

  test('ignores png opt-out on the native backend', async () => {
    const requests: string[] = [];
    const loader = createNativeLoader(assetFetch(requests));
    await loader.initWasm(undefined, { png: false });
    expect(requests).toEqual([]);
    expect(loader.isPngExportAvailable()).toBe(true);
  });
});
