/**
 * Measurement entry point — the Rust (wasm `docx-text`) measurement source,
 * its font registry contracts, and the float-aware block pipeline. There is
 * no browser text measurement: every extent comes from the engine (real font
 * bytes, Word metrics) or, for degenerate input, from deterministic
 * synthetic arithmetic.
 *
 * Font bytes come from the injected `BundledFontProvider` when a host supplies
 * one, otherwise from `@betteroffice/fonts` via an optional dynamic import
 * ({@link configureDefaultFonts} points that at a CDN). With neither, families
 * the document does not embed fall through to synthetic metrics — real bytes
 * are worth 15.4 points of exact page-count accuracy against Word.
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
export type {
  BundledFontProvider,
  BundledFontProviderSource,
  EmbeddedFaceInput,
  FontScript,
} from './fontRegistry';
export {
  configureDefaultFonts,
  resolveDefaultFontProvider,
  type DefaultFontOptions,
} from './defaultFontProvider';
