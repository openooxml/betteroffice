/* tslint:disable */
/* eslint-disable */

/**
 * wasm wrapper over [`session::close_display_list`]: drop a handle so its
 * parsed display list is freed. Idempotent.
 */
export function close_display_list(handle: number): void;

/**
 * wasm wrapper over [`session::hit_test_regions_by_handle`]: region-aware hit
 * test against a stored display list. `Err` on an unknown/closed handle so the
 * caller can fall back to [`hit_test_regions_json`].
 */
export function hit_test_regions_by_handle(handle: number, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
 * `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,"noteId"?,`
 * `"pos":n|null,"target":"text"|"image"|"none"}` (or `"null"` for an
 * out-of-range page). The plain `hit_test_json` export stays body-only.
 */
export function hit_test_regions_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`session::open_display_list`]: parse a display list once
 * and return a handle the by-handle query exports reuse (no per-query
 * re-parse). The caller frees it with [`close_display_list`]. `Err` on
 * malformed JSON — the caller then stays on the JSON-arg path.
 */
export function open_display_list(display_list: string): number;

/**
 * wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
 * stored display list. `Err` on an unknown/closed handle so the caller can
 * fall back to [`range_rects_json`].
 */
export function range_rects_by_handle(handle: number, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_json`]: display-list JSON + document range
 * in, JSON array of page-local rects out.
 */
export function range_rects_json(display_list: string, from: number, to: number): string;

/**
 * @internal
 */
export function range_rects_on_pages_by_handle(handle: number, from: number, to: number, first_page: number, last_page: number): string;

/**
 * @internal
 */
export function range_rects_on_pages_json(display_list: string, from: number, to: number, first_page: number, last_page: number): string;

/**
 * wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
 * range rects against a stored display list. `region` is
 * `"body" | "header" | "footer" | "footnote" | "endnote"`; `part_id` scopes
 * header/footer to one HF part and names the note id for a note region.
 * `Err` on an unknown/closed handle so the caller can fall back to
 * [`range_rects_region_json`].
 */
export function range_rects_region_by_handle(handle: number, region: string, part_id: string, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
 * `region` is `"body" | "header" | "footer" | "footnote" | "endnote"`;
 * `part_id` scopes a header/footer to one HF part (empty for body / match-any)
 * and names the note id for a note region. The `from`/`to` refer to that
 * region's doc. The plain `range_rects_json` export stays body-only.
 */
export function range_rects_region_json(display_list: string, region: string, part_id: string, from: number, to: number): string;

/**
 * wasm wrapper over [`session::update_display_list`]: apply a page-delta
 * update to a stored display list so an incremental rebuild re-parses only
 * its changed pages. `Err` closes the handle first, so the caller's fallback
 * (a fresh [`open_display_list`]) can never race a half-updated list.
 */
export function update_display_list(handle: number, update: string): void;

export function vertical_move_by_handle(handle: number, position: number, direction: string, goal_x: number): string;

export function vertical_move_json(display_list: string, position: number, direction: string, goal_x: number): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_start: () => void;
    readonly close_display_list: (a: number) => void;
    readonly hit_test_regions_by_handle: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly hit_test_regions_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly memory: WebAssembly.Memory;
    readonly open_display_list: (a: number, b: number) => [number, number, number];
    readonly range_rects_by_handle: (a: number, b: number, c: number) => [number, number, number, number];
    readonly range_rects_json: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly range_rects_on_pages_by_handle: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly range_rects_on_pages_json: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly range_rects_region_by_handle: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly range_rects_region_json: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number, number];
    readonly update_display_list: (a: number, b: number, c: number) => [number, number];
    readonly vertical_move_by_handle: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly vertical_move_json: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
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
