/**
 * Rust measurement and deterministic font-provider contracts.
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
