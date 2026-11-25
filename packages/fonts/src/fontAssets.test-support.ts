import { file } from 'bun';
import baseLatinManifest from '../assets/base-latin.json';
import { getFontCdnConfiguration, implicitFontCdnBases } from '../../../shared/office-session/fontCdnRuntime';
import { createFontRuntime } from './fontRuntime';
import { fontAssetResponse as baseFontAssetResponse } from './test-assets';

export function fontRuntimeWithoutPacks(base = new URL('../assets/', import.meta.url)) {
  return createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'base' ? Object.fromEntries(baseLatinManifest.faces.map(({ file }) =>
      [file, () => new URL(file, base)])) : undefined,
  }));
}

export async function fontAssetResponse(input: RequestInfo | URL): Promise<Response> {
  const path = new URL(String(input)).pathname;
  if (path.includes('/fonts-cjk@')) {
    const asset = path.split('/assets/')[1]!;
    return new Response(await file(new URL(`../../fonts-cjk/assets/${asset}`, import.meta.url)).arrayBuffer());
  }
  return baseFontAssetResponse(input);
}

export function withBrowserFontGlobals(page = true): () => void {
  implicitFontCdnBases().clear();
  const key = Symbol.for('betteroffice.fonts.cdn.test-globals');
  const scope = globalThis as Record<symbol, unknown>;
  const previous = scope[key];
  scope[key] = page ? { window: { document: {} } } : {};
  return () => {
    implicitFontCdnBases().clear();
    if (previous === undefined) delete scope[key];
    else scope[key] = previous;
  };
}
