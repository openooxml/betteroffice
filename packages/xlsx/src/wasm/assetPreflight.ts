export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/preflight/xlsx_preflight_bg.wasm', import.meta.url);
}
