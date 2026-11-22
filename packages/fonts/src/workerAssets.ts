import { workerFontAssets } from '../../../shared/office-session/optionalFontPacks';
import { readWasmSync } from '../../../shared/office-session/readWasmAsset';

export { setWorkerFontAssets, workerFontAssets } from '../../../shared/office-session/optionalFontPacks';

export async function bundledDecoderUrl(): Promise<URL | undefined> {
  const href = (await workerFontAssets('decoder'))?.urls?.['fontpack.wasm'];
  return href === undefined ? undefined : new URL(href);
}

export async function readWasm(url: URL): Promise<Uint8Array> {
  if (url.protocol === 'file:') {
    const { Deno, Bun } = globalThis as unknown as {
      Deno?: { readFile(url: URL): Promise<Uint8Array> };
      Bun?: { file(url: URL): { arrayBuffer(): Promise<ArrayBuffer> } };
    };
    try {
      const bytes = readWasmSync(url, true);
      if (bytes) return bytes;
      if (Deno) return await Deno.readFile(url);
      if (Bun) return new Uint8Array(await Bun.file(url).arrayBuffer());
    } catch (error) {
      if (error instanceof Error && error.name === 'NotFound') Object.assign(error, { code: 'ENOENT' });
      throw error;
    }
  }
  const response = await fetch(url, { signal: AbortSignal.timeout(30_000) });
  if (!response.ok) {
    throw Object.assign(new Error(`Cannot load font decoder wasm: HTTP ${response.status}`), { decoderStatus: response.status });
  }
  return new Uint8Array(await response.arrayBuffer());
}

export function absoluteHref(href: string): string {
  return new URL(href).href;
}
