export function wasmAssetUrl(png = true): URL {
  return new URL(png ? 'betteroffice-native:xlsx' : 'betteroffice-native:xlsx-lean');
}
