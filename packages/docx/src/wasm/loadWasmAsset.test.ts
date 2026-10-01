/**
 * The default `file:` init path must not reach global fetch: Node's rejects
 * that scheme, and a DOM shim replaces Bun's with one that rejects it too.
 */

import { afterEach, describe, expect, it, mock, spyOn } from 'bun:test';

import {
  compileWasm,
  createWasmModuleState,
  readWasmSync,
  wasmModuleMemories,
  type WasmAsyncInput,
} from './loadWasmAsset';

const ASSET = new URL('./generated/opc/ooxml_opc_bg.wasm', import.meta.url);
const realFetch = globalThis.fetch;
const WASM_BYTES = new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00]);

afterEach(() => {
  globalThis.fetch = realFetch;
});

/** Mirrors wasm-bindgen's `__wbg_init`: string/URL inputs go through global fetch. */
function recordingInit(seen: WasmAsyncInput[]) {
  return async ({
    module_or_path,
  }: {
    module_or_path: WasmAsyncInput | Promise<WasmAsyncInput>;
  }): Promise<void> => {
    const resolved = await module_or_path;
    seen.push(resolved);
    if (resolved instanceof WebAssembly.Module) return;
    const bytes =
      typeof resolved === 'string' || resolved instanceof URL
        ? await (await fetch(resolved)).arrayBuffer()
        : (resolved as BufferSource);
    await WebAssembly.compile(bytes);
  };
}

function stateFor(assetUrl: () => URL, seen: WasmAsyncInput[], shareModule = false) {
  return createWasmModuleState({
    label: 'test',
    preloadName: 'preloadTestWasm',
    assetUrl,
    initAsync: recordingInit(seen),
    initSync: () => {},
    shareModule,
  });
}

describe('preload', () => {
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

describe('compileWasm', () => {
  it('returns an existing module', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    expect(await compileWasm(module)).toBe(module);
  });

  it('compiles bytes', async () => {
    expect(await compileWasm(WASM_BYTES)).toBeInstanceOf(WebAssembly.Module);
  });

  it.each(['url', 'string', 'request'])('fetches a %s input', async (kind) => {
    const url = new URL('https://cdn.example/module.wasm');
    const input = kind === 'url' ? url : kind === 'string' ? url.href : new Request(url);
    const fetch = mock(() =>
      Promise.resolve(new Response(WASM_BYTES, { headers: { 'Content-Type': 'application/wasm' } }))
    );
    globalThis.fetch = fetch as unknown as typeof globalThis.fetch;
    expect(await compileWasm(input)).toBeInstanceOf(WebAssembly.Module);
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(fetch).toHaveBeenCalledWith(input);
  });

  it('compiles a wasm response', async () => {
    const response = new Response(WASM_BYTES, { headers: { 'Content-Type': 'application/wasm' } });
    expect(await compileWasm(response)).toBeInstanceOf(WebAssembly.Module);
  });

  it.each(['basic', 'cors', 'default'])('falls back to bytes for a wrong-MIME %s response', async (type) => {
    const response = new Response(WASM_BYTES, { headers: { 'Content-Type': 'text/plain' } });
    Object.defineProperty(response, 'type', { value: type });
    const streaming = spyOn(WebAssembly, 'compileStreaming').mockRejectedValue(new Error('MIME'));
    const compile = spyOn(WebAssembly, 'compile');
    try {
      expect(await compileWasm(response)).toBeInstanceOf(WebAssembly.Module);
      expect(streaming).toHaveBeenCalledWith(response);
      expect(compile).toHaveBeenCalledTimes(1);
    } finally {
      streaming.mockRestore();
      compile.mockRestore();
    }
  });

  it('rejects a non-ok response without falling back to bytes', async () => {
    const response = new Response(WASM_BYTES, { status: 404 });
    const error = new Error('HTTP');
    const streaming = spyOn(WebAssembly, 'compileStreaming').mockRejectedValue(error);
    const compile = spyOn(WebAssembly, 'compile');
    try {
      await expect(compileWasm(response)).rejects.toBe(error);
      expect(compile).not.toHaveBeenCalled();
    } finally {
      streaming.mockRestore();
      compile.mockRestore();
    }
  });

  it.each(['application/wasm', 'opaque'])('rethrows a streaming error for %s', async (kind) => {
    const response = new Response(WASM_BYTES, {
      headers: { 'Content-Type': kind === 'opaque' ? 'text/plain' : kind },
    });
    if (kind === 'opaque') Object.defineProperty(response, 'type', { value: 'opaque' });
    const error = new Error('streaming failed');
    const streaming = spyOn(WebAssembly, 'compileStreaming').mockRejectedValue(error);
    const compile = spyOn(WebAssembly, 'compile');
    try {
      await expect(compileWasm(response)).rejects.toBe(error);
      expect(compile).not.toHaveBeenCalled();
    } finally {
      streaming.mockRestore();
      compile.mockRestore();
    }
  });

  it('compiles response bytes when compileStreaming is unavailable', async () => {
    const streaming = WebAssembly.compileStreaming;
    (WebAssembly as { compileStreaming?: typeof streaming }).compileStreaming = undefined;
    try {
      expect(await compileWasm(new Response(WASM_BYTES))).toBeInstanceOf(WebAssembly.Module);
    } finally {
      WebAssembly.compileStreaming = streaming;
    }
  });
});

describe('shared module', () => {
  const remote = new URL('https://cdn.example/docx_edit_bg.wasm');

  function installFetch() {
    const fetch = mock(() =>
      Promise.resolve(new Response(WASM_BYTES, { headers: { 'Content-Type': 'application/wasm' } }))
    );
    globalThis.fetch = fetch as unknown as typeof globalThis.fetch;
    return fetch;
  }

  it('does not compile a shared module after synchronous disk initialization', async () => {
    const fetch = installFetch();
    const streaming = spyOn(WebAssembly, 'compileStreaming');
    const compile = spyOn(WebAssembly, 'compile');
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => ASSET, seen, true);
    try {
      state.ensure();
      expect(await state.sharedModule()).toBeNull();
      expect(seen).toEqual([]);
      expect(fetch).not.toHaveBeenCalled();
      expect(streaming).not.toHaveBeenCalled();
      expect(compile).not.toHaveBeenCalled();
    } finally {
      streaming.mockRestore();
      compile.mockRestore();
    }
  });

  it('does not compile the default asset during or after an explicit bytes preload', async () => {
    const fetch = installFetch();
    const module = new WebAssembly.Module(WASM_BYTES);
    let resolve!: (module: WebAssembly.Module) => void;
    const compile = spyOn(WebAssembly, 'compile').mockReturnValue(
      new Promise((settle) => {
        resolve = settle;
      })
    );
    const streaming = spyOn(WebAssembly, 'compileStreaming');
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    try {
      const loading = state.preload(WASM_BYTES);
      expect(await state.sharedModule()).toBeNull();
      expect(seen).toEqual([WASM_BYTES]);
      expect(fetch).not.toHaveBeenCalled();
      expect(streaming).not.toHaveBeenCalled();
      expect(compile).toHaveBeenCalledTimes(1);
      expect(compile).toHaveBeenCalledWith(WASM_BYTES);

      resolve(module);
      await loading;
      expect(await state.sharedModule()).toBeNull();
      expect(seen).toEqual([WASM_BYTES]);
      expect(fetch).not.toHaveBeenCalled();
      expect(streaming).not.toHaveBeenCalled();
      expect(compile).toHaveBeenCalledTimes(1);
    } finally {
      streaming.mockRestore();
      compile.mockRestore();
    }
  });

  it('starts a shared compile that a following preload reuses', async () => {
    const fetch = installFetch();
    const streaming = spyOn(WebAssembly, 'compileStreaming');
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    try {
      const compiling = state.sharedModule();
      expect(state.sharedModule()).toBe(compiling);
      expect(fetch).toHaveBeenCalledTimes(1);
      const loading = state.preload();
      expect(state.module()).toBe(compiling as Promise<WebAssembly.Module>);
      await loading;
      const module = await compiling;
      if (!module) throw new Error('Expected a shared module');
      expect(module).toBeInstanceOf(WebAssembly.Module);
      expect(seen).toEqual([module]);
      expect(state.sharedModule()).toBe(compiling);
      expect(fetch).toHaveBeenCalledTimes(1);
      expect(fetch).toHaveBeenCalledWith(remote);
      expect(streaming).toHaveBeenCalledTimes(1);
    } finally {
      streaming.mockRestore();
    }
  });

  it('shares one fetch and compile between preload and module', async () => {
    const fetch = installFetch();
    const streaming = spyOn(WebAssembly, 'compileStreaming');
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    try {
      const loading = state.preload();
      const compiling = state.module();
      expect(state.module()).toBe(compiling);
      expect(state.sharedModule()).toBe(compiling);
      expect(state.preload()).toBe(loading);
      await loading;
      const module = await compiling;
      expect(seen).toEqual([module]);
      expect(module).toBeInstanceOf(WebAssembly.Module);
      expect(state.module()).toBe(compiling);
      expect(state.sharedModule()).toBe(compiling);
      expect(fetch).toHaveBeenCalledTimes(1);
      expect(fetch).toHaveBeenCalledWith(remote);
      expect(streaming).toHaveBeenCalledTimes(1);
    } finally {
      streaming.mockRestore();
    }
  });

  it('retries a rejected compile and initialization', async () => {
    const fetch = installFetch();
    fetch.mockImplementationOnce(() => Promise.resolve(new Response(new Uint8Array([0]))));
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    const loading = state.preload();
    const failed = state.module();
    await expect(loading).rejects.toThrow();
    const retried = state.module();
    expect(retried).not.toBe(failed);
    await state.preload();
    expect(seen).toEqual([await retried]);
    expect(fetch).toHaveBeenCalledTimes(2);
  });

  it('recompiles after a shared module fails to initialize', async () => {
    const fetch = installFetch();
    const streaming = spyOn(WebAssembly, 'compileStreaming');
    const seen: WasmAsyncInput[] = [];
    const init = recordingInit(seen);
    const error = new Error('init failed');
    const state = createWasmModuleState({
      label: 'test',
      preloadName: 'preloadTestWasm',
      assetUrl: () => remote,
      initAsync: async (input) => {
        await init(input);
        if (seen.length === 1) throw error;
      },
      initSync: () => {},
      shareModule: true,
    });
    try {
      const loading = state.preload();
      const compiling = state.module();
      await expect(loading).rejects.toBe(error);
      const first = await compiling;
      expect(seen).toEqual([first]);
      expect(fetch).toHaveBeenCalledTimes(1);
      expect(streaming).toHaveBeenCalledTimes(1);

      const retry = state.preload();
      const recompiling = state.module();
      expect(recompiling).not.toBe(compiling);
      await retry;
      const second = await recompiling;
      expect(second).not.toBe(first);
      expect(seen).toEqual([first, second]);
      expect(fetch).toHaveBeenCalledTimes(2);
      expect(streaming).toHaveBeenCalledTimes(2);
    } finally {
      streaming.mockRestore();
    }
  });

  it('initializes and caches a module supplied by the host without fetching', async () => {
    const fetch = installFetch();
    const module = new WebAssembly.Module(WASM_BYTES);
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    const loading = state.preloadFrom(Promise.resolve(module));
    expect(state.preload()).toBe(loading);
    expect(state.preloadFrom(Promise.resolve(null))).toBe(loading);
    await loading;
    expect(seen).toEqual([module]);
    expect(await state.module()).toBe(module);
    await state.preloadFrom(Promise.resolve(null));
    expect(seen).toEqual([module]);
    expect(fetch).not.toHaveBeenCalled();
  });

  it.each(['null', 'rejection'])('falls back to module on a host %s', async (kind) => {
    const fetch = installFetch();
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    const source = kind === 'null' ? Promise.resolve(null) : Promise.reject(new Error('host failed'));
    await state.preloadFrom(source);
    expect(seen).toEqual([await state.module()]);
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it('honours and caches the first explicit module', async () => {
    const fetch = installFetch();
    const first = new WebAssembly.Module(WASM_BYTES);
    const second = new WebAssembly.Module(WASM_BYTES);
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    const loading = state.preload(first);
    expect(state.preload(second)).toBe(loading);
    await loading;
    expect(await state.module()).toBe(first);
    expect(seen).toEqual([first]);
    expect(fetch).not.toHaveBeenCalled();
  });

  it('preserves explicit bytes and retries a failed explicit initialization', async () => {
    const fetch = installFetch();
    const seen: WasmAsyncInput[] = [];
    const state = stateFor(() => remote, seen, true);
    await expect(state.preload(new Uint8Array([0]))).rejects.toThrow();
    await state.preload(WASM_BYTES);
    expect(seen).toEqual([new Uint8Array([0]), WASM_BYTES]);
    expect(fetch).not.toHaveBeenCalled();
  });

  it.each(['null', 'rejection'])(
    'preserves the default URL fallback without sharing on %s',
    async (kind) => {
      const fetch = installFetch();
      const seen: WasmAsyncInput[] = [];
      const state = stateFor(() => remote, seen);
      const source = kind === 'null' ? Promise.resolve(null) : Promise.reject(new Error('host failed'));
      await state.preloadFrom(source);
      expect(seen).toEqual([remote]);
      expect(fetch).toHaveBeenCalledTimes(1);
    }
  );
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
