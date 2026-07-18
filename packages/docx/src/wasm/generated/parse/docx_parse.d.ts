/* tslint:disable */
/* eslint-disable */

/**
 * Wasm control-plane entry: safe ZIP -> bounded XML -> typed relationships.
 */
export function parse_docx_relationships(data: Uint8Array): string;

/**
 * Legacy staged Rust S2 entry retained for ABI compatibility.
 */
export function parse_docx_s2(data: Uint8Array): string;

/**
 * Legacy staged Rust S3 entry retained for ABI compatibility.
 */
export function parse_docx_s3(data: Uint8Array): string;

/**
 * Legacy staged Rust S4 entry retained for ABI compatibility.
 */
export function parse_docx_s4(data: Uint8Array): string;

/**
 * Legacy staged Rust S5 entry retained for ABI compatibility.
 */
export function parse_docx_s5(data: Uint8Array): string;

/**
 * Legacy staged Rust S6 entry retained for ABI compatibility.
 */
export function parse_docx_s6(data: Uint8Array): string;

/**
 * Legacy staged Rust S7 entry retained for ABI compatibility.
 */
export function parse_docx_s7(data: Uint8Array): string;

/**
 * Legacy staged Rust S8 entry retained for ABI compatibility.
 */
export function parse_docx_s8(data: Uint8Array): string;

/**
 * S9 production read facade: one safe package pass to the full Document wire.
 */
export function parse_docx_s9(data: Uint8Array, options_json: string): string;

/**
 * Focused wasm leaf used by hostile-input and facade tests.
 */
export function parse_relationships_xml(xml: Uint8Array, part_path: string): string;

/**
 * Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
 */
export function rezip_docx(entries: any): Uint8Array;

/**
 * Legacy staged Rust S10 serializer entry retained for ABI compatibility.
 */
export function serialize_docx_s10(request_json: string): string;

/**
 * Legacy staged Rust S11 serializer entry retained for ABI compatibility.
 */
export function serialize_docx_s11(request_json: string): string;

/**
 * Legacy staged Rust S12 serializer entry retained for ABI compatibility.
 */
export function serialize_docx_s12(request_json: string): string;

/**
 * Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
 */
export function unzip_docx(data: Uint8Array): any;

/**
 * S13 production-capable package writer: typed model + original package -> DOCX.
 */
export function write_docx_s13_wasm(request_json: string, original_docx: Uint8Array): Uint8Array;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly parse_docx_relationships: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s2: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s3: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s4: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s5: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s6: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s7: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s8: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s9: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly parse_relationships_xml: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly serialize_docx_s10: (a: number, b: number) => [number, number, number, number];
    readonly serialize_docx_s11: (a: number, b: number) => [number, number, number, number];
    readonly serialize_docx_s12: (a: number, b: number) => [number, number, number, number];
    readonly write_docx_s13_wasm: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly rezip_docx: (a: any) => [number, number, number, number];
    readonly unzip_docx: (a: number, b: number) => [number, number, number];
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
