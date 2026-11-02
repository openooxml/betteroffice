import wasmInit, { font_preflight, initSync, engineKind } from './enginePreflight';
import { wasmAssetUrl } from './assetPreflight';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

export const preflightWasmState = createWasmModuleState({
  label: 'docx-preflight',
  preloadName: 'preloadPreflightWasm',
  assetUrl: wasmAssetUrl,
  engineKind,
  initAsync: wasmInit,
  initSync,
});

/** @experimental Load the font preflight wasm on demand. Idempotent. */
export function preloadPreflightWasm(input?: WasmAsyncInput): Promise<void> {
  return preflightWasmState.preload(input);
}

/** @experimental */
export interface FontPreflightFace {
  family: string;
  bold: boolean;
  italic: boolean;
  scripts: string[];
}

/** @experimental */
export interface FontPreflight {
  faces: FontPreflightFace[];
  /** False when some font rules could not be resolved. */
  complete: boolean;
}

/** @experimental Font faces needed by DOCX bytes, without opening the editing engine. */
export async function fontPreflight(bytes: Uint8Array): Promise<FontPreflight> {
  await preflightWasmState.preload();
  return JSON.parse(font_preflight(bytes)) as FontPreflight;
}
