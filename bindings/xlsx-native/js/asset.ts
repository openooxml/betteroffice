export function wasmAssetUrl(png = true): URL {
  return new URL(png ? 'betteroffice-native:xlsx' : 'betteroffice-native:xlsx-lean');
}

export function rasterWasmAssetUrl(): URL {
  throw new Error('@betteroffice/xlsx-native renders PNG in the native engine and ships no raster wasm');
}
