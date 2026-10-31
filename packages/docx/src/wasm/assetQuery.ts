export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/query/docx_query_bg.wasm', import.meta.url);
}
