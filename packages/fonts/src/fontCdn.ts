import { version as latinVersion } from '../../fonts-latin/package.json';
import { version as intlVersion } from '../../fonts-intl/package.json';
import { version as cjkVersion } from '../../fonts-cjk/package.json';
import { isImplicitFontCdnError, type FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
import {
  getFontCdnConfiguration, JSDELIVR, setFontCdnConfiguration, STALL_MS, type FontCdnConfiguration,
} from '../../../shared/office-session/fontCdn';

export { fontCdnCacheKey, isBrowserFontPage, workerFontCdnConfiguration, type FontCdnConfiguration } from '../../../shared/office-session/fontCdn';

type Pack = FontPackDiagnostic['pack'];
export type FontRequestPolicy = number | Pack;

class ImplicitFontCdnError extends Error {
  readonly diagnostic: FontPackDiagnostic;

  constructor(pack: Pack, cause: unknown) {
    const reason = cause instanceof Error ? cause.message : String(cause);
    const detail = `Cannot load @betteroffice/fonts-${pack} from cdn.jsdelivr.net: ${reason}. To opt out, install @betteroffice/fonts-${pack} or call configureFontCdn({ baseUrl: false }).`;
    super(detail);
    this.name = 'ImplicitFontCdnError';
    this.diagnostic = { pack, reason: 'cdn-unavailable', detail };
  }
}

export function fontAssetError(request: FontRequestPolicy | undefined, message: string): Error {
  return typeof request === 'string' ? new ImplicitFontCdnError(request, message) : new Error(message);
}

export async function decodeFontAsset<T>(request: FontRequestPolicy | undefined, decode: () => Promise<T>): Promise<T> {
  try { return await decode(); }
  catch (error) {
    if (typeof request === 'string' && !isImplicitFontCdnError(error)) throw new ImplicitFontCdnError(request, error);
    throw error;
  }
}

export async function readImplicitFontBytes(file: string, url: URL, pack: Pack): Promise<ArrayBuffer> {
  const controller = new AbortController();
  const { signal } = controller;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  let rejectStall!: (reason: unknown) => void;
  const stalled = new Promise<never>((_, reject) => { rejectStall = reject; });
  const arm = () => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      const error = new DOMException(`No font data received for ${STALL_MS} ms`, 'TimeoutError');
      controller.abort(error);
      rejectStall(error);
      if (reader) void reader.cancel(error).catch(() => {});
    }, STALL_MS);
  };
  const read = async () => {
    const response = await fetch(url, { signal, credentials: 'omit', referrerPolicy: 'no-referrer' });
    if (signal.aborted) throw signal.reason;
    if (!response.ok && response.type !== 'opaque') {
      throw new Error(`Failed to fetch bundled font ${file}: HTTP ${response.status}`);
    }
    arm();
    const body = response.body;
    if (typeof body?.getReader !== 'function') return response.arrayBuffer();
    const streamReader = body.getReader();
    reader = streamReader;
    const chunks: Uint8Array[] = [];
    let byteLength = 0;
    try {
      for (;;) {
        const { done, value } = await streamReader.read();
        if (signal.aborted) throw signal.reason;
        if (done) break;
        if (value.byteLength === 0) continue;
        chunks.push(value);
        byteLength += value.byteLength;
        arm();
      }
    } finally {
      streamReader.releaseLock();
      reader = undefined;
    }
    const bytes = new Uint8Array(byteLength);
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return bytes.buffer;
  };
  arm();
  try {
    return await Promise.race([read(), stalled]);
  } catch (error) {
    throw new ImplicitFontCdnError(pack, error);
  } finally {
    clearTimeout(timer);
  }
}

export interface FontCdnOptions {
  baseUrl: string | URL | false;
  timeoutMs?: number;
}

const packageVersions: Record<string, string> = { latin: latinVersion, intl: intlVersion, cjk: cjkVersion };

/** @internal */
export function fontCdnAssetUrl(pack: { pack: string; version: string }, file: string, baseUrl?: string): URL {
  if (baseUrl !== undefined) return new URL(`${pack.pack}/${pack.version}/${file}`, baseUrl);
  const version = packageVersions[pack.pack];
  if (version === undefined) throw new Error(`Unknown font pack: ${pack.pack}`);
  return new URL(`${JSDELIVR}@betteroffice/fonts-${pack.pack}@${version}/assets/${file}`);
}

/** Configures the optional fallback for fonts absent from installed packs. */
export function configureFontCdn(options: FontCdnOptions): void {
  const timeoutMs = options.timeoutMs ?? 30_000;
  if (!Number.isInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > 4_294_967_295) {
    throw new RangeError('Font CDN timeoutMs must be a positive 32-bit integer');
  }
  let baseUrl: string | false = false;
  if (options.baseUrl !== false) {
    const value = String(options.baseUrl);
    const url = new URL(value.endsWith('/') ? value : `${value}/`,
      typeof location === 'undefined' ? undefined : location.href);
    if (url.protocol !== 'https:' && url.protocol !== 'http:') {
      throw new TypeError('Font CDN baseUrl must use HTTP or HTTPS');
    }
    baseUrl = url.href;
  }
  setFontCdnConfiguration({ baseUrl, timeoutMs }, true);
}

/** @internal */
export function fontCdnConfiguration(): FontCdnConfiguration {
  return getFontCdnConfiguration() ?? { baseUrl: false, timeoutMs: 30_000 };
}
