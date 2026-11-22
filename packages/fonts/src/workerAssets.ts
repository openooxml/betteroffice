import type { FontPackSource } from '../../../shared/office-session/optionalFontPacks';
import { readWasmSync } from '../../../shared/office-session/readWasmAsset';

let source: FontPackSource = async () => undefined;
let timeout: ((url: URL) => number | undefined) | undefined;

export function setWorkerFontAssets(next: FontPackSource, nextTimeout?: typeof timeout): void {
  source = next;
  timeout = nextTimeout;
}

export const workerFontAssets: FontPackSource = (pack) => source(pack);
export const workerFontTimeout = (url: URL) => timeout?.(url);

export async function bundledDecoderUrl(): Promise<URL | undefined> {
  const href = (await source('decoder'))?.['fontpack.wasm'];
  return href === undefined ? undefined : new URL(href);
}

export async function readWasm(url: URL): Promise<Uint8Array> {
  const bytes = readWasmSync(url);
  if (bytes) return bytes;
  const response = await fetch(url, { signal: AbortSignal.timeout(30_000) });
  if (!response.ok) throw new Error(`Cannot load font decoder wasm: HTTP ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

export function absoluteHref(href: string): string {
  return new URL(href).href;
}
