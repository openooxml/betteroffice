/* tslint:disable */
/* eslint-disable */

/**
 * wasm compatibility wrapper. Resident engine users call
 * [`build_display_list_value`] and keep the typed result.
 */
export function build_display_list_json(input: string): string;

/**
 * Drop every registered measurement font (ids restart at 0). Callers must
 * re-register before the next `measure_paragraph_json`.
 */
export function clear_measure_fonts(): void;

/**
 * wasm wrapper over [`session::close_display_list`]: drop a handle so its
 * parsed display list is freed. Idempotent.
 */
export function close_display_list(handle: number): void;

/**
 * wasm wrapper over [`hit::hit_test_json`]: display-list JSON + page-local
 * point in, PM position (or `null`) as JSON out.
 */
export function hit_test_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`session::hit_test_regions_by_handle`]: region-aware hit
 * test against a stored display list. `Err` on an unknown/closed handle so the
 * caller can fall back to [`hit_test_regions_json`].
 */
export function hit_test_regions_by_handle(handle: number, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
 * `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` (or `"null"` for
 * an out-of-range page). The legacy `hit_test_json` export stays body-only.
 */
export function hit_test_regions_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`layout_to_json`].
 */
export function layout_document_json(input: string): string;

/**
 * wasm wrapper over [`ooxml_text::measure_paragraph_json`]: measurement
 * input JSON in, `ParagraphExtent` JSON out. An `Err` whose message starts
 * with `"UNSUPPORTED"` means the caller must fall back to browser
 * measurement for that block.
 */
export function measure_paragraph_json(input: string): string;

/**
 * wasm wrapper over [`session::open_display_list`]: parse a display list once
 * and return a handle the by-handle query exports reuse (no per-query
 * re-parse). The caller frees it with [`close_display_list`]. `Err` on
 * malformed JSON — the caller then stays on the JSON-arg path.
 */
export function open_display_list(display_list: string): number;

/**
 * wasm wrapper over [`ooxml_text::FontStore::outline_glyph_json`]: the outline
 * of a registered font's glyph, in font design units, as JSON:
 * `{"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L","x":..,"y":..},
 * {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
 * {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}`.
 * The canvas caches this per `(fontId, glyphId)` and scales by `size/upem`,
 * flipping y at draw time. `cmds` is empty for a blank glyph (space).
 */
export function outline_glyph_json(font_id: number, glyph_id: number): string;

/**
 * wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
 * stored display list. `Err` on an unknown/closed handle so the caller can
 * fall back to [`range_rects_json`].
 */
export function range_rects_by_handle(handle: number, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_json`]: display-list JSON + PM range
 * in, JSON array of page-local rects out.
 */
export function range_rects_json(display_list: string, from: number, to: number): string;

/**
 * wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
 * range rects against a stored display list. `region` is
 * `"body" | "header" | "footer"`; `r_id` scopes header/footer to one HF part.
 * `Err` on an unknown/closed handle so the caller can fall back to
 * [`range_rects_region_json`].
 */
export function range_rects_region_by_handle(handle: number, region: string, r_id: string, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
 * `region` is `"body" | "header" | "footer"`; `r_id` scopes a header/footer to
 * one HF part (empty for body / match-any). The `from`/`to` refer to that
 * region's PM doc. The legacy `range_rects_json` export stays body-only.
 */
export function range_rects_region_json(display_list: string, region: string, r_id: string, from: number, to: number): string;

/**
 * Register a font for measurement from raw sfnt bytes; returns the font id
 * that `measure_paragraph_json` inputs reference in their `fontChains`.
 * Malformed bytes (attacker-controlled embedded fonts) are rejected as an
 * error at this boundary, mirroring `FontStore::register`.
 */
export function register_measure_font(bytes: Uint8Array): number;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly build_display_list_json: (a: number, b: number) => [number, number, number, number];
    readonly hit_test_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly hit_test_regions_by_handle: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly hit_test_regions_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly layout_document_json: (a: number, b: number) => [number, number, number, number];
    readonly measure_paragraph_json: (a: number, b: number) => [number, number, number, number];
    readonly open_display_list: (a: number, b: number) => [number, number, number];
    readonly outline_glyph_json: (a: number, b: number) => [number, number, number, number];
    readonly range_rects_by_handle: (a: number, b: number, c: number) => [number, number, number, number];
    readonly range_rects_json: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly range_rects_region_by_handle: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly range_rects_region_json: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number, number];
    readonly register_measure_font: (a: number, b: number) => [number, number, number];
    readonly close_display_list: (a: number) => void;
    readonly clear_measure_fonts: () => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
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
