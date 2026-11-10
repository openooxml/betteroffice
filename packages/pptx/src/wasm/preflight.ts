import wasmInit, { font_preflight, engineKind } from './enginePreflight';
import { wasmAssetUrl } from './assetPreflight';
import { createPreflightState, type FontPreflight, type PreflightWasmInput } from '../../../../shared/office-session/fontPreflight';

export type { FontPreflight, FontPreflightFace } from '../../../../shared/office-session/fontPreflight';

export const preflightWasmState = createPreflightState({
  assetUrl: wasmAssetUrl,
  engineKind,
  initAsync: wasmInit,
});

/** @experimental Load the font preflight wasm on demand. Idempotent. */
export function preloadPreflightWasm(input?: PreflightWasmInput): Promise<void> {
  return preflightWasmState.preload(input);
}

/** @experimental Font faces for a zero-based slide index. Defaults to the first slide. */
export async function fontPreflight(bytes: Uint8Array, activeSlide = 0): Promise<FontPreflight> {
  if (!Number.isInteger(activeSlide) || activeSlide < 0 || activeSlide > 0xffffffff) {
    throw new RangeError('activeSlide must be an unsigned slide index');
  }
  await preflightWasmState.preload();
  return JSON.parse(font_preflight(bytes, activeSlide)) as FontPreflight;
}
