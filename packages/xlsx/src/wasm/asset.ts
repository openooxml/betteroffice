export function wasmAssetUrl(png = true): URL {
  return png
    ? new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url)
    : new URL('./generated/xlsx_wasm_lean_bg.wasm', import.meta.url);
}
