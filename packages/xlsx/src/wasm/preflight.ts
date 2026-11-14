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

/** @experimental Font faces for the active sheet's A1:T50 region, without opening the engine. */
export async function fontPreflight(bytes: Uint8Array): Promise<FontPreflight> {
  await preflightWasmState.preload();
  return JSON.parse(font_preflight(bytes)) as FontPreflight;
}
