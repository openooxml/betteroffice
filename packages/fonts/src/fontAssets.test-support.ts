import { file } from 'bun';
import baseLatinManifest from '../assets/base-latin.json';
import { getFontCdnConfiguration, implicitFontCdnBases } from '../../../shared/office-session/fontCdn';
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
  const document = page ? globalThis.document ?? {} : undefined;
  const globals = { window: page ? { document } : undefined, document, process: undefined, Deno: undefined, Bun: undefined };
  const previous = Object.keys(globals).map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  for (const [key, value] of Object.entries(globals)) {
    Object.defineProperty(globalThis, key, { configurable: true, writable: true, value });
  }
  return () => {
    implicitFontCdnBases().clear();
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  };
}
