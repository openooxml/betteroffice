/**
 * Measurement entry point — the Rust (wasm `docx-text`) measurement source,
 * its font registry contracts, and the float-aware block pipeline. There is
 * no browser text measurement: every extent comes from the engine (real font
 * bytes, Word metrics) or, for degenerate input, from deterministic
 * synthetic arithmetic.
 * @packageDocumentation
 * @public
 */

export {
  createRustMeasureSource,
  getRustTextEngine,
  type RustMeasureSource,
  type RustTextEngine,
  type ResidentFontRequirement,
  type ResidentMeasurementConfig,
} from './rustMeasureSource';
export type { BundledFontProvider, EmbeddedFaceInput, FontScript } from './fontRegistry';
