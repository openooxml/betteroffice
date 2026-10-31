export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/preflight/docx_preflight_bg.wasm', import.meta.url);
}
