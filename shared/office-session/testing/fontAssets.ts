import baseLatinManifest from '../../../packages/fonts/assets/base-latin.json';
import { createFontRuntime } from '../../../packages/fonts/src/fontRuntime';
import { fontAssetResponse as baseFontAssetResponse } from '../../../packages/fonts/src/test-assets';
import { getFontCdnConfiguration } from '../fontCdn';

export function fontRuntimeWithoutPacks(base = new URL('../../../packages/fonts/assets/', import.meta.url)) {
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
    return new Response(await Bun.file(new URL(`../../../packages/fonts-cjk/assets/${asset}`, import.meta.url)).arrayBuffer());
  }
  return baseFontAssetResponse(input);
}

export function withBrowserFontGlobals(page = true): () => void {
  const document = page ? globalThis.document ?? {} : undefined;
  const globals = { window: page ? { document } : undefined, document, process: undefined, Deno: undefined };
  const previous = Object.keys(globals).map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  for (const [key, value] of Object.entries(globals)) {
    Object.defineProperty(globalThis, key, { configurable: true, writable: true, value });
  }
  return () => {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  };
}
