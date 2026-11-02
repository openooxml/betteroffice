export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/parse/docx_parse_bg.wasm', import.meta.url);
}
