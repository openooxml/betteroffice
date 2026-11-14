// @betteroffice/docx declaration wasm/loadWasmAsset.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Shared init-state machinery for the generated wasm modules (opc / edit / parse).
 * The binaries are external assets (gitignored,
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
export declare function compileWasm(input: WasmAsyncInput): Promise<WebAssembly.Module>;
/** Read a `file:` wasm asset synchronously from disk; undefined off-Node. */
export declare function readWasmSync(url: URL): Uint8Array | undefined;
export interface WasmModuleState {
    /** Ready after init; loading during init or a shared compile, including before instantiation. */
    status(): 'idle' | 'loading' | 'ready';
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
/**
 * One instantiated engine module's memory in this thread; heap counters only for modules that
 * keep them. A native engine that counts its heap is listed too, with no linear memory.
 */
export interface WasmModuleMemory extends Partial<WasmHeapStats> {
    label: string;
    /**
     * Size of its linear memory. It only grows, so it is also the module's
     * high-water mark. A native engine has no linear memory: this is its peak
     * heap bytes, its high-water mark, and `maximumBytes` is absent.
     */
    bufferBytes: number;
    maximumBytes?: number;
    /** Listed by a native engine, which has no linear memory and no 4 GiB limit. */
    native?: true;
}
/** The address space of a wasm32 memory: 4 GiB. */
export declare const WASM32_MEMORY_LIMIT_BYTES: number;
/**
 * Starts counting the heap of every native engine in this process, so that
 * {@link wasmModuleMemories} lists it. A native engine does not count by default:
 * counting costs a little on every allocation. It starts at zero and stays on for
 * the life of the process, for every thread and session in it: one caller starting
 * it changes what every session in the process reports, and it cannot be stopped.
 * Each thread's figures stay its own. Call it before opening documents: memory allocated
 * earlier is subtracted when freed without having been counted (saturating at
 * zero), so live and peak bytes can read low. Pressure events on a native engine
 * also need a host memory budget. WebAssembly engines always count, each module
 * instance its own heap, from the start, and cannot be stopped either. Returns
 * whether every engine module that can count now does.
 */
export declare function startHeapCounting(): boolean;
/** The memory of every engine module instantiated in this thread, by label. */
export declare function wasmModuleMemories(): WasmModuleMemory[];
export declare function createWasmModuleState(options: {
    label: string;
    preloadName: string;
    assetUrl: () => URL;
    engineKind?: 'wasm' | 'native';
    initAsync: (input: {
        module_or_path: WasmAsyncInput | Promise<WasmAsyncInput>;
    }) => Promise<unknown>;
    initSync: (input: {
        module: WasmSyncInput;
    }) => unknown;
    shareModule?: boolean;
    /** Reads the module's allocator counters, if it counts; called only once initialized. */
    heap?: () => WasmHeapStats | undefined;
    /** Starts a native engine's heap counting; returns whether it counts. */
    startHeapCounting?: () => boolean;
}): WasmModuleState;
