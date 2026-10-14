/**
 * Shared init-state machinery for the four generated wasm modules (opc /
 * layout / edit / parse). The binaries are external assets (gitignored,
 * rebuilt by scripts/build-docx-wasm.ts, shipped in dist/generated/) —
 * never embedded base64.
 *
 * Two ways a module becomes ready:
 *  - `preload()` (async): browsers and workers — passes the asset URL to the
 *    wasm-bindgen glue, which fetches and instantiates it. A `file:` asset is
 *    read from disk instead, since `fetch` may reject that scheme.
 *  - `ensure()` (sync): Node/Bun/SSR — reads the asset from disk via
 *    `process.getBuiltinModule` (no static `node:fs` import, so browser
 *    bundlers never see a node builtin) and feeds `initSync`. In a browser
 *    without a prior `preload()` this throws with a call-to-action.
 *
 * IMPORTANT — URL geometry: every `new URL('./generated/…', import.meta.url)`
 * literal must live in a module physically inside `src/wasm/`, and each loader
 * is its own root-named tsup entry. That keeps the relative path valid in BOTH
 * layouts: `src/wasm/*` next to `src/wasm/generated/` in source mode, and
 * root-level chunks next to `dist/generated/` in package builds (the xlsx
 * package established the pattern).
 */

export type WasmSyncInput = BufferSource | WebAssembly.Module;
export type WasmAsyncInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

/** @internal */
export async function compileWasm(input: WasmAsyncInput): Promise<WebAssembly.Module> {
  if (input instanceof WebAssembly.Module) return input;
  if (
    typeof input === 'string' ||
    (typeof Request === 'function' && input instanceof Request) ||
    (typeof URL === 'function' && input instanceof URL)
  ) {
    input = await fetch(input);
  }
  if (typeof Response === 'function' && input instanceof Response) {
    if (typeof WebAssembly.compileStreaming === 'function') {
      try {
        return await WebAssembly.compileStreaming(input);
      } catch (error) {
        if (
          !input.ok ||
          !['basic', 'cors', 'default'].includes(input.type) ||
          input.headers.get('Content-Type') === 'application/wasm'
        ) {
          throw error;
        }
      }
    }
    return WebAssembly.compile(await input.arrayBuffer());
  }
  return WebAssembly.compile(input as BufferSource);
}

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}
interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

function builtinModule<T>(name: string): T | undefined {
  const proc = (
    globalThis as {
      process?: { getBuiltinModule?: (id: string) => unknown };
    }
  ).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

/** Read a `file:` wasm asset synchronously from disk; undefined off-Node. */
export function readWasmSync(url: URL): Uint8Array | undefined {
  if (url.protocol !== 'file:') return undefined;
  const fs = builtinModule<NodeFsLike>('node:fs');
  const nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) return undefined;
  try {
    // `.href` and not the URL object: fileURLToPath brand-checks its argument,
    // so a shim supplying a non-native URL would otherwise fail the read.
    return fs.readFileSync(nodeUrl.fileURLToPath(url.href));
  } catch {
    return undefined;
  }
}

/**
 * Default async input. wasm-bindgen fetches URL inputs, but `file:` is served
 * only by Bun's own `fetch` — Node's rejects it, and a DOM shim replaces Bun's
 * with one that rejects it too. So `file:` assets resolve to disk bytes, as
 * `ensure()` already does; http(s) stays a URL for the browser streaming path.
 */
function defaultAsyncInput(url: URL): WasmAsyncInput {
  if (url.protocol !== 'file:') return url;
  return readWasmSync(url) ?? url;
}

export interface WasmModuleState {
  /** Async init from the packaged asset URL (or an explicit override). */
  preload(input?: WasmAsyncInput): Promise<void>;
  /** @internal */
  module(): Promise<WebAssembly.Module>;
  /** @internal This thread's shared compile, or null when it loaded the engine from another input. */
  sharedModule(): Promise<WebAssembly.Module | null>;
  /** @internal */
  preloadFrom(source: Promise<WebAssembly.Module | null>): Promise<void>;
  /** Sync guard used by every call site; disk-inits on Node/Bun, throws in a browser before `preload()`. */
  ensure(): void;
}

/** A wasm module's Rust heap, from the allocator's counters. */
export interface WasmHeapStats {
  /** Bytes allocated now. */
  liveBytes: number;
  /** The most bytes allocated at once since the module started. */
  peakBytes: number;
  /** Size of the allocation the memory could not satisfy, or 0. */
  failedAllocationBytes: number;
}

/** One instantiated wasm module's memory in this thread; heap counters only for modules that keep them. */
export interface WasmModuleMemory extends Partial<WasmHeapStats> {
  label: string;
  /**
   * Size of its linear memory. It only grows, so it is also the module's
   * high-water mark.
   */
  bufferBytes: number;
}

/** The address space of a wasm32 memory: 4 GiB. */
export const WASM32_MEMORY_LIMIT_BYTES = 4 * 1024 * 1024 * 1024;

interface RegisteredModule {
  label: string;
  memory: () => WebAssembly.Memory | null;
  heap?: () => WasmHeapStats | undefined;
}

const registered: RegisteredModule[] = [];

/** The memory of every wasm module instantiated in this thread, by label. */
export function wasmModuleMemories(): WasmModuleMemory[] {
  const out: WasmModuleMemory[] = [];
  for (const module of registered) {
    const memory = module.memory();
    if (!memory) continue;
    out.push({ label: module.label, bufferBytes: memory.buffer.byteLength, ...module.heap?.() });
  }
  return out;
}

function exportedMemory(output: unknown): WebAssembly.Memory | null {
  const memory = (output as { memory?: unknown } | null | undefined)?.memory;
  return typeof WebAssembly !== 'undefined' && memory instanceof WebAssembly.Memory ? memory : null;
}

export function createWasmModuleState(options: {
  label: string;
  preloadName: string;
  assetUrl: () => URL;
  initAsync: (input: { module_or_path: WasmAsyncInput | Promise<WasmAsyncInput> }) => Promise<unknown>;
  initSync: (input: { module: WasmSyncInput }) => unknown;
  shareModule?: boolean;
  /** Reads the module's allocator counters, if it counts; called only once initialized. */
  heap?: () => WasmHeapStats | undefined;
}): WasmModuleState {
  let initialized = false;
  let pending: Promise<void> | undefined;
  let compiled: Promise<WebAssembly.Module> | undefined;
  let memory: WebAssembly.Memory | null = null;
  registered.push({
    label: options.label,
    memory: () => (initialized ? memory : null),
    ...(options.heap ? { heap: options.heap } : {}),
  });

  function module(): Promise<WebAssembly.Module> {
    if (!compiled) {
      const compiling = compileWasm(defaultAsyncInput(options.assetUrl())).catch((error) => {
        if (compiled === compiling) compiled = undefined;
        throw error;
      });
      compiled = compiling;
    }
    return compiled;
  }

  function initialize(input: WasmAsyncInput | Promise<WasmAsyncInput>): Promise<void> {
    pending = options.initAsync({ module_or_path: input }).then(
      (output) => {
        memory = exportedMemory(output);
        initialized = true;
      },
      (error: unknown) => {
        pending = undefined;
        compiled = undefined;
        throw error instanceof Error ? error : new Error(String(error));
      }
    );
    return pending;
  }

  return {
    module,
    sharedModule(): Promise<WebAssembly.Module | null> {
      if (compiled) return compiled;
      if (initialized || pending) return Promise.resolve(null);
      return module();
    },
    preload(input?: WasmAsyncInput): Promise<void> {
      if (initialized) return Promise.resolve();
      if (pending) return pending;
      if (typeof WebAssembly !== 'undefined' && input instanceof WebAssembly.Module) {
        compiled = Promise.resolve(input);
      }
      return initialize(
        input ?? (options.shareModule ? module() : defaultAsyncInput(options.assetUrl()))
      );
    },
    preloadFrom(source: Promise<WebAssembly.Module | null>): Promise<void> {
      if (initialized) return Promise.resolve();
      if (pending) return pending;
      return initialize(
        source.catch(() => null).then((input) => {
          if (input) {
            compiled = Promise.resolve(input);
            return input;
          }
          return options.shareModule ? module() : defaultAsyncInput(options.assetUrl());
        })
      );
    },
    ensure(): void {
      if (initialized) return;
      const bytes = readWasmSync(options.assetUrl());
      if (bytes) {
        memory = exportedMemory(options.initSync({ module: bytes }));
        initialized = true;
        return;
      }
      throw new Error(
        `${options.label} wasm is not initialized; await ${options.preloadName}() before first use ` +
          '(browser builds load the external wasm asset asynchronously)'
      );
    },
  };
}
