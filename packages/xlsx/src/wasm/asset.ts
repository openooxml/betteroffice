export function wasmAssetUrl(): URL {
  return new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url);
}
