/**
 * Loader for the docx-layout wasm (the Rust display-list builder and the
 * ooxml-text measure surface). External-asset pattern — see ./loadWasmAsset.ts
 * for the init contract and URL-geometry invariant.
 *
 * IMPORTANT: reach this module only via dynamic `import()` — its consumers are
 * the canvas renderer and the Rust measurement source, and those dynamic-import
 * seams await {@link preloadLayoutWasm} before first use.
 *
 * Requires a CSP that permits wasm compilation (`wasm-unsafe-eval`).
 */

// vendored wasm-bindgen glue (typed by ./docx_layout.d.ts)
import wasmInit, {
  initSync,
  build_display_list_json,
  clear_measure_fonts,
  hit_test_regions_json,
  layout_document_json,
  measure_paragraph_json,
  range_rects_json,
  register_measure_font,
} from './generated/layout/docx_layout.js';
// Resolve optional exports dynamically when generated declarations omit them.
import * as layoutGlue from './generated/layout/docx_layout.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const state = createWasmModuleState({
  label: 'docx-layout',
  preloadName: 'preloadLayoutWasm',
  assetUrl: () => new URL('./generated/layout/docx_layout_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

/** Load + instantiate the layout wasm (browser path). Idempotent. */
export function preloadLayoutWasm(input?: WasmAsyncInput): Promise<void> {
  return state.preload(input);
}

/**
 * Region-aware hit test over a display list: page-local point in,
 * `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` (or `"null"` for
 * an out-of-range page) JSON out. Header/footer hits identify the HF doc
 * by `rId`; their `pos` refers to that doc, not the body doc.
 */
export function hitTestRegionsJson(
  displayList: string,
  pageIndex: number,
  x: number,
  y: number
): string {
  state.ensure();
  return hit_test_regions_json(displayList, pageIndex, x, y);
}

/**
 * Highlight rects for a body document range over a display list: JSON array of
 * `{pageIndex, x, y, width, height}` in page-local px. Body-only — HF doc
 * positions live in different docs and are never consulted.
 */
export function rangeRectsJson(displayList: string, from: number, to: number): string {
  state.ensure();
  return range_rects_json(displayList, from, to);
}

// Optional session exports avoid repeated display-list serialization.

type OpenDisplayListExport = (json: string) => number;
type CloseDisplayListExport = (handle: number) => void;
type UpdateDisplayListExport = (handle: number, update: string) => void;
type HitTestRegionsByHandleExport = (
  handle: number,
  pageIndex: number,
  x: number,
  y: number
) => string;
type VerticalMoveJsonExport = (
  displayList: string,
  position: number,
  direction: string,
  goalX: number
) => string;
type VerticalMoveByHandleExport = (
  handle: number,
  position: number,
  direction: string,
  goalX: number
) => string;
type RangeRectsByHandleExport = (handle: number, from: number, to: number) => string;
type RangeRectsRegionJsonExport = (
  displayList: string,
  region: string,
  rId: string,
  from: number,
  to: number
) => string;
type RangeRectsRegionByHandleExport = (
  handle: number,
  region: string,
  rId: string,
  from: number,
  to: number
) => string;

function glueExport<T>(name: string): T | undefined {
  const fn = (layoutGlue as unknown as Record<string, unknown>)[name];
  return typeof fn === 'function' ? (fn as T) : undefined;
}

/**
 * True when the embedded layout wasm carries the session-handle query exports.
 * Callers gate the handle path on this and otherwise use the JSON-arg exports.
 */
export function hasDisplayListSession(): boolean {
  state.ensure();
  return glueExport<OpenDisplayListExport>('open_display_list') !== undefined;
}

/** Reports whether region-aware range rectangles are available. */
export function hasRangeRectsRegion(): boolean {
  state.ensure();
  return glueExport<RangeRectsRegionJsonExport>('range_rects_region_json') !== undefined;
}

/** Opens a reusable display-list handle. */
export function openDisplayList(displayList: string): number {
  state.ensure();
  const open = glueExport<OpenDisplayListExport>('open_display_list');
  if (!open) throw new Error('open_display_list is not available in the embedded layout wasm yet');
  return open(displayList);
}

/** Drop a display-list handle so its parsed list is freed. No-op when the export is absent. */
export function closeDisplayList(handle: number): void {
  state.ensure();
  glueExport<CloseDisplayListExport>('close_display_list')?.(handle);
}

/**
 * Apply a page-delta update to a stored display list, re-parsing only its
 * changed pages. Throws when the export is absent or the update is
 * inconsistent — the Rust side closes the handle on failure, so the facade's
 * fallback (a fresh `openDisplayList`) is always safe.
 */
export function updateDisplayList(handle: number, update: string): void {
  state.ensure();
  const apply = glueExport<UpdateDisplayListExport>('update_display_list');
  if (!apply) {
    throw new Error('update_display_list is not available in the embedded layout wasm yet');
  }
  apply(handle, update);
}

/** True when the embedded layout wasm carries the page-delta update export. */
export function hasDisplayListUpdate(): boolean {
  state.ensure();
  return glueExport<UpdateDisplayListExport>('update_display_list') !== undefined;
}

/**
 * Region-aware hit test against a stored display list (by handle). Throws when
 * the export is absent or the handle is unknown/closed — the facade catches it
 * and falls back to `hitTestRegionsJson`.
 */
export function hitTestRegionsByHandle(
  handle: number,
  pageIndex: number,
  x: number,
  y: number
): string {
  state.ensure();
  const query = glueExport<HitTestRegionsByHandleExport>('hit_test_regions_by_handle');
  if (!query)
    throw new Error('hit_test_regions_by_handle is not available in the embedded layout wasm yet');
  return query(handle, pageIndex, x, y);
}

/** Resolve the closest caret position on the adjacent visual line. */
export function verticalMoveJson(
  displayList: string,
  position: number,
  direction: 'up' | 'down',
  goalX: number
): string {
  state.ensure();
  const query = glueExport<VerticalMoveJsonExport>('vertical_move_json');
  if (!query) throw new Error('vertical_move_json is not available in the embedded layout wasm');
  return query(displayList, position, direction, goalX);
}

/** Resolve the closest caret position on the adjacent visual line by handle. */
export function verticalMoveByHandle(
  handle: number,
  position: number,
  direction: 'up' | 'down',
  goalX: number
): string {
  state.ensure();
  const query = glueExport<VerticalMoveByHandleExport>('vertical_move_by_handle');
  if (!query)
    throw new Error('vertical_move_by_handle is not available in the embedded layout wasm');
  return query(handle, position, direction, goalX);
}

/**
 * Highlight rects for a body document range against a stored display list (by handle).
 * Throws when the export is absent or the handle is unknown/closed.
 */
export function rangeRectsByHandle(handle: number, from: number, to: number): string {
  state.ensure();
  const query = glueExport<RangeRectsByHandleExport>('range_rects_by_handle');
  if (!query)
    throw new Error('range_rects_by_handle is not available in the embedded layout wasm yet');
  return query(handle, from, to);
}

/** Returns region range rects or throws when the optional export is unavailable. */
export function rangeRectsRegionJson(
  displayList: string,
  region: string,
  rId: string,
  from: number,
  to: number
): string {
  state.ensure();
  const query = glueExport<RangeRectsRegionJsonExport>('range_rects_region_json');
  if (!query)
    throw new Error('range_rects_region_json is not available in the embedded layout wasm yet');
  return query(displayList, region, rId, from, to);
}

/**
 * Region-aware highlight rects against a stored display list (by handle). Throws
 * when the export is absent or the handle is unknown/closed.
 */
export function rangeRectsRegionByHandle(
  handle: number,
  region: string,
  rId: string,
  from: number,
  to: number
): string {
  state.ensure();
  const query = glueExport<RangeRectsRegionByHandleExport>('range_rects_region_by_handle');
  if (!query)
    throw new Error(
      'range_rects_region_by_handle is not available in the embedded layout wasm yet'
    );
  return query(handle, region, rId, from, to);
}

/** `{ measured, options, layout }` JSON in, `DisplayList` JSON out. Throws on inputs the Rust builder rejects. */
export function buildDisplayListJson(input: string): string {
  state.ensure();
  return build_display_list_json(input);
}

/**
 * Pagination: `{ measured, options }` JSON in (the golden-fixture envelope),
 * `Layout` JSON out. Throws for inputs the Rust layout kernel rejects — the
 * layout pipeline surfaces the error and keeps the previous committed
 * layout; there is no fallback engine.
 */
export function layoutDocumentJson(input: string): string {
  state.ensure();
  return layout_document_json(input);
}

/**
 * Register raw sfnt bytes with the measurement FontStore; returns the font id
 * that `measureParagraphJson` inputs reference in their `fontChains`. Throws
 * on malformed bytes (embedded fonts are attacker-controlled — the engine
 * rejects unparseable input at this boundary).
 */
export function registerMeasureFont(bytes: Uint8Array): number {
  state.ensure();
  return register_measure_font(bytes);
}

/** Drop every registered measurement font (ids restart at 0). Callers must re-register before the next `measureParagraphJson`. */
export function clearMeasureFonts(): void {
  state.ensure();
  clear_measure_fonts();
}

/**
 * Measurement input JSON in, `ParagraphExtent` JSON out. Throws with a
 * message starting `"UNSUPPORTED"` for blocks the Rust engine cannot measure
 * yet — the caller must fall back to browser measurement for that block.
 */
export function measureParagraphJson(input: string): string {
  state.ensure();
  return measure_paragraph_json(input);
}

type OutlineGlyphExport = (fontId: number, glyphId: number) => string;

// Missing outline exports fall back to fillText.
function resolveOutlineGlyphExport(): OutlineGlyphExport | undefined {
  const fn = (layoutGlue as unknown as Record<string, unknown>).outline_glyph_json;
  return typeof fn === 'function' ? (fn as OutlineGlyphExport) : undefined;
}

/** Returns a glyph outline or throws when it is unavailable. */
export function outlineGlyphJson(fontId: number, glyphId: number): string {
  state.ensure();
  const outlineExport = resolveOutlineGlyphExport();
  if (!outlineExport) {
    throw new Error('outline_glyph_json is not available in the embedded layout wasm yet');
  }
  return outlineExport(fontId, glyphId);
}
