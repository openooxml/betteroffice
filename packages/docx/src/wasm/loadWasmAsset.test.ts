/**
 * The default `file:` init path must not reach global fetch: Node's rejects
 * that scheme, and a DOM shim replaces Bun's with one that rejects it too.
 */

import { afterEach, describe, expect, it } from 'bun:test';

import {
  createWasmModuleState,
  readWasmSync,
  wasmModuleMemories,
  type WasmAsyncInput,
} from './loadWasmAsset';

const ASSET = new URL('./generated/opc/ooxml_opc_bg.wasm', import.meta.url);
const realFetch = globalThis.fetch;

/** Mirrors wasm-bindgen's `__wbg_init`: string/URL inputs go through global fetch. */
function recordingInit(seen: WasmAsyncInput[]) {
  return async ({
    module_or_path,
  }: {
    module_or_path: WasmAsyncInput | Promise<WasmAsyncInput>;
  }): Promise<void> => {
    const resolved = await module_or_path;
    seen.push(resolved);
    const bytes =
      typeof resolved === 'string' || resolved instanceof URL
        ? await (await fetch(resolved)).arrayBuffer()
        : (resolved as BufferSource);
    await WebAssembly.compile(bytes);
  };
}

function stateFor(assetUrl: () => URL, seen: WasmAsyncInput[]) {
  return createWasmModuleState({
    label: 'test',
    preloadName: 'preloadTestWasm',
    assetUrl,
    initAsync: recordingInit(seen),
    initSync: () => {},
  });
}

describe('preload', () => {
  afterEach(() => {
    globalThis.fetch = realFetch;
  });

  it('initializes a file: asset when global fetch rejects the file: scheme', async () => {
    globalThis.fetch = (() =>
      Promise.reject(
        new Error(`Failed to fetch from "${ASSET.href}": URL scheme "file" is not supported.`)
      )) as unknown as typeof fetch;
    const seen: WasmAsyncInput[] = [];

    await stateFor(() => ASSET, seen).preload();

    expect(seen).toHaveLength(1);
    expect(ArrayBuffer.isView(seen[0])).toBe(true);
  });

  it('still streams a non-file: asset through fetch', async () => {
    const wasm = readWasmSync(ASSET);
    expect(wasm).toBeDefined();
    const requested: unknown[] = [];
    globalThis.fetch = ((input: unknown) => {
      requested.push(input);
      return Promise.resolve(new Response(wasm as Uint8Array<ArrayBuffer>));
    }) as unknown as typeof fetch;
    const remote = new URL('https://cdn.example/ooxml_opc_bg.wasm');
    const seen: WasmAsyncInput[] = [];

    await stateFor(() => remote, seen).preload();

    expect(seen).toEqual([remote]);
    expect(requested).toEqual([remote]);
  });

  it('honours an explicit input over the packaged asset', async () => {
    const bytes = readWasmSync(ASSET) as Uint8Array;
    globalThis.fetch = (() => Promise.reject(new Error('unreachable'))) as unknown as typeof fetch;
    const seen: WasmAsyncInput[] = [];

    await stateFor(() => ASSET, seen).preload(bytes);

    expect(seen).toEqual([bytes]);
  });
});

describe('readWasmSync', () => {
  it('reads through a foreign URL implementation', () => {
    // fileURLToPath brand-checks objects; happy-dom's URL is native, but a
    // shim that supplies its own must not defeat the read.
    const foreign = { href: ASSET.href, protocol: 'file:' } as URL;
    expect(readWasmSync(foreign)?.byteLength).toBeGreaterThan(0);
  });
});

describe('wasmModuleMemories', () => {
  it('lists a module once it is initialized, with its allocator counters', async () => {
    const memory = new WebAssembly.Memory({ initial: 2 });
    const heap = { liveBytes: 10, peakBytes: 20, failedAllocationBytes: 0 };
    const state = createWasmModuleState({
      label: 'memory-test',
      preloadName: 'preloadMemoryTestWasm',
      assetUrl: () => ASSET,
      initAsync: async () => ({ memory }),
      initSync: () => ({ memory }),
      heap: () => heap,
    });
    expect(wasmModuleMemories().find((module) => module.label === 'memory-test')).toBeUndefined();

    await state.preload(new Uint8Array());

    expect(wasmModuleMemories().find((module) => module.label === 'memory-test')).toEqual({
      label: 'memory-test',
      bufferBytes: 2 * 65536,
      ...heap,
    });
    memory.grow(1);
    expect(wasmModuleMemories().find((module) => module.label === 'memory-test')?.bufferBytes).toBe(
      3 * 65536
    );
  });

  it('omits allocator counters for a module without them', () => {
    const memory = new WebAssembly.Memory({ initial: 1 });
    createWasmModuleState({
      label: 'memory-test-sync',
      preloadName: 'preloadMemoryTestSyncWasm',
      assetUrl: () => ASSET,
      initAsync: async () => ({ memory }),
      initSync: () => ({ memory }),
    }).ensure();

    expect(wasmModuleMemories().find((module) => module.label === 'memory-test-sync')).toEqual({
      label: 'memory-test-sync',
      bufferBytes: 65536,
    });
  });
});
