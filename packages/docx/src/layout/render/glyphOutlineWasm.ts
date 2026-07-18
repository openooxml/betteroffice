/**
 * Lazy accessor for the embedded wasm glyph-outline provider — the production
 * source a canvas host feeds to {@link GlyphCache}.
 *
 * Reached through the SAME dynamic `import('../wasm/index')` specifier the
 * display-list builder (`buildRustDisplayList`) and the measure engine
 * (`getRustTextEngine`) use, so all three resolve to ONE wasm module instance —
 * i.e. one measurement `FontStore`. That shared identity is load-bearing:
 * `outlineGlyphJson(fontId, glyphId)` must read glyph outlines from the very
 * store `registerMeasureFont` populated during measurement, otherwise the ids
 * in a GlyphRun would resolve against an empty store and every glyph would
 * rasterize as `.notdef`. Keeping the import dynamic also keeps the ~800KB
 * inlined wasm out of the default bundle (see `layout/wasm/index.ts`).
 *
 * @experimental part of the rust-canvas-engine glyph pipeline; shape may evolve.
 */

import type { GlyphOutlineProvider } from './glyphCache';

/**
 * Resolve the wasm-backed {@link GlyphOutlineProvider} (`outlineGlyphJson`).
 * The returned function is synchronous and initializes the wasm instance on
 * first call; it throws when the embedded module lacks the `outline_glyph_json`
 * export, which the canvas backend catches and repaints the run via `fillText`
 * so text is never invisible. Shares the module promise with the display-list
 * builder, so on the canvas path the wasm is already resolved by the time the
 * first display list lands.
 */
export function loadGlyphOutlineProvider(): Promise<GlyphOutlineProvider> {
  return import('../wasm/index').then(async (m) => {
    await m.preloadLayoutWasm();
    return m.outlineGlyphJson;
  });
}
