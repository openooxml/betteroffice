import latinManifest from '../../fonts-latin/assets/manifest.json';
import cjkManifest from '../../fonts-cjk/assets/manifest.json';
import baseLatinManifest from '../assets/base-latin.json';

const latinFiles = new Set(latinManifest.faces.map((face) => face.file));
const cjkFiles = new Set(cjkManifest.faces.map((face) => face.file));
const baseLatinFiles = new Set(baseLatinManifest.faces.map((face) => face.file));

export async function fontAssetResponse(input: RequestInfo | URL): Promise<Response> {
  const path = new URL(String(input)).pathname;
  if (path.endsWith('/decoder/fontpack.wasm')) {
    return new Response(await Bun.file(new URL('../dist/decoder/fontpack.wasm', import.meta.url)).arrayBuffer());
  }
  const file = path.slice(path.lastIndexOf('/') + 1);
  if (baseLatinFiles.has(file)) {
    return new Response(await Bun.file(new URL(`../assets/${file}`, import.meta.url)).arrayBuffer());
  }
  const pack = latinFiles.has(file) ? 'fonts-latin' : cjkFiles.has(file) ? 'fonts-cjk' : 'fonts-intl';
  return new Response(await Bun.file(new URL(`../../${pack}/assets/${file}`, import.meta.url)).arrayBuffer());
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
