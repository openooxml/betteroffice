/**
 * Measurement entry point — the Rust (wasm `docx-text`) measurement source,
 * its font registry contracts, and the float-aware block pipeline. There is
 * no browser text measurement: every extent comes from the engine (real font
 * bytes, Word metrics) or, for degenerate input, from deterministic
 * synthetic arithmetic.
 * @packageDocumentation
 * @public
 */

// Float-zone geometry (pure math shared by measurement and pagination)
export {
  clampFloatingWrapMargins,
  rectsToFloatingZones,
  getFloatingMargins,
  type FloatingImageZone,
  type FloatingExclusionRect,
  type FloatingLineSegmentZone,
} from './floatingZones';

export {
  measureBlocksWithFloats,
  type MeasureBlockFn,
  type FloatPageGeometry,
} from './measureBlocksPipeline';

// The Rust measurement source (wasm docx-text engine) — the sole path
export {
  createRustMeasureSource,
  getRustTextEngine,
  type FontStyle,
  type RustMeasureSource,
  type RustMeasureStats,
  type RustTextEngine,
} from './rustMeasureSource';
export type { BundledFontProvider, EmbeddedFaceInput, FontScript } from './fontRegistry';

// Paragraph identity hash (the Rust measure memo's key)
export { hashParagraphBlock } from './paragraphHash';
