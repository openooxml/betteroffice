export function wasmAssetUrl(): URL {
  return new URL('./generated/pptx_wasm_bg.wasm', import.meta.url);
}
