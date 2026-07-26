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
// Namespace view of the same glue. Exports the generated .d.ts does not
// declare are resolved by name through this, since a named import of an
// undeclared export would not typecheck.
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

// ---------------------------------------------------------------------------
// session-handle query surface
//
// The JSON-arg queries above re-send the whole display list on every call and
// Rust re-parses it each time, which dominates the cost of click and drag
// paths. The session exports parse the list ONCE (`open_display_list` → a
// handle) and answer many queries by handle, reusing the same hit and range
// logic, so results are byte-identical and only the cost differs.
//
// These exports are OPTIONAL: the embedded wasm may not carry them. Every
// wrapper below throws when its export is missing, so callers feature-detect
// with `hasDisplayListSession()` / `hasDisplayListUpdate()` /
// `hasRangeRectsRegion()` and fall back to the JSON-arg path.
// ---------------------------------------------------------------------------

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

/**
 * True when the embedded layout wasm carries the region-aware range-rect
 * exports (`range_rects_region_json` and its by-handle twin). Callers gate
 * header/footer selection and caret geometry on this and otherwise return
 * `[]`; probing avoids invoking — and so dropping the session handle over — an
 * absent export.
 */
export function hasRangeRectsRegion(): boolean {
  state.ensure();
  return glueExport<RangeRectsRegionJsonExport>('range_rects_region_json') !== undefined;
}

/**
 * Parse a display list once and return the handle the by-handle queries reuse.
 * The caller owns it and must `closeDisplayList` it. Throws when the export is
 * absent or the JSON is malformed; the query facade catches that and stays on
 * the JSON-arg path.
 */
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

/**
 * Closest caret position on the adjacent visual line: `{"position","goalX"}`
 * JSON, or `"null"` when there is no line in that direction. `goalX` is the
 * page-local x the caret holds across successive moves.
 */
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

/**
 * Region-aware highlight rects: `region` is `"body" | "header" | "footer"`,
 * and `rId` scopes a header/footer to one part (empty string for body, or to
 * match any). `from`/`to` are positions in THAT region's document. Throws when
 * the export is absent, so the facade falls back to `[]`.
 */
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

// Optional like the session exports: when the embedded wasm has no
// `outline_glyph_json` this resolves to undefined, `outlineGlyphJson` throws,
// and the canvas backend paints the run with `fillText` instead.
function resolveOutlineGlyphExport(): OutlineGlyphExport | undefined {
  const fn = (layoutGlue as unknown as Record<string, unknown>).outline_glyph_json;
  return typeof fn === 'function' ? (fn as OutlineGlyphExport) : undefined;
}

/**
 * Glyph outline for `(fontId, glyphId)` as JSON:
 * `{"upem":2048,"cmds":[{"t":"M",…},{"t":"L",…},{"t":"Q",…},{"t":"C",…},{"t":"Z"}]}`
 * — commands in font units, y-up per the font convention. `fontId` is a
 * measurement FontStore id from {@link registerMeasureFont}, `glyphId` a glyph
 * index from shaping. Throws when the export is absent or the glyph cannot be
 * extracted; the canvas backend catches that and paints the run via
 * `fillText`, so text is never invisible.
 */
export function outlineGlyphJson(fontId: number, glyphId: number): string {
  state.ensure();
  const outlineExport = resolveOutlineGlyphExport();
  if (!outlineExport) {
    throw new Error('outline_glyph_json is not available in the embedded layout wasm yet');
  }
  return outlineExport(fontId, glyphId);
}
