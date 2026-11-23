import { version as latinVersion } from '../../fonts-latin/package.json';
import { version as intlVersion } from '../../fonts-intl/package.json';
import { version as cjkVersion } from '../../fonts-cjk/package.json';
import type { FontCdnDiagnostic } from '../../../shared/fontCdnDiagnostics';
import {
  getFontCdnConfiguration, implicitFontCdnBases, JSDELIVR, setFontCdnConfiguration, STALL_MS,
  type FontCdnConfiguration, type ImplicitFontCdnBase,
} from '../../../shared/office-session/fontCdnRuntime';

export { fontCdnCacheKey, isBrowserFontPage, workerFontCdnConfiguration, type FontCdnConfiguration } from '../../../shared/office-session/fontCdnRuntime';

type Pack = FontCdnDiagnostic['pack'];

class ImplicitFontCdnError extends Error {
  readonly diagnostic: FontCdnDiagnostic;

  constructor(pack: Pack, cause: unknown) {
    const reason = cause instanceof Error ? cause.message : String(cause);
    const detail = `Cannot load @betteroffice/fonts-${pack} from cdn.jsdelivr.net: ${reason}. To opt out, install @betteroffice/fonts-${pack} or call configureFontCdn({ baseUrl: false }).`;
    super(detail);
    this.name = 'ImplicitFontCdnError';
    this.diagnostic = { pack, reason: 'cdn-unavailable', detail };
  }
}

const COOLDOWN_MS = 60_000;

class HttpFontError extends Error {
  constructor(file: string, readonly status: number) {
    super(`Failed to fetch bundled font ${file}: HTTP ${status}`);
  }
}

type CheckFontAsset = <T>(check: () => T | Promise<T>) => Promise<T>;

interface ImplicitFontRequest {
  check(url: URL): void;
  load<T>(
    file: string, url: URL, decode: (bytes: ArrayBuffer, check: CheckFontAsset) => Promise<T>, checkDependencies?: () => void,
  ): Promise<T>;
}

export type FontRequestPolicy = number | ImplicitFontRequest;

export function implicitFontRequest(pack: Pack): ImplicitFontRequest {
  return {
    check: checkImplicitUrlCooldown,
    load: (file, url, decode, checkDependencies) => loadImplicitFontAsset(file, url, pack, decode, checkDependencies),
  };
}

function failureScope(cause: unknown, invalidBytes: boolean): 'base' | 'url' {
  if (invalidBytes || (cause instanceof HttpFontError && (cause.status === 404 || cause.status === 410))) return 'url';
  return 'base';
}

function baseKey(url: URL): string {
  return url.origin + url.pathname.split('/assets/', 1)[0];
}

function coolingFailure(base: ImplicitFontCdnBase, href: string): Error | undefined {
  const now = Date.now();
  if (base.unavailable && now < base.unavailable.until) return base.unavailable.error;
  const invalid = base.invalid.get(href);
  if (invalid && now < invalid.until) return invalid.error;
  if (invalid) base.invalid.delete(href);
}

export function checkImplicitUrlCooldown(url: URL): void {
  const base = implicitFontCdnBases().get(baseKey(url));
  const failure = base && coolingFailure(base, url.href);
  if (failure) throw failure;
}

export function checkImplicitSliceCooldown(index: URL): void {
  const base = implicitFontCdnBases().get(baseKey(index));
  if (!base) return;
  const indexFailure = coolingFailure(base, index.href);
  if (indexFailure) throw indexFailure;
  const root = new URL('./', index).href;
  for (const href of base.invalid.keys()) {
    if (!href.startsWith(root)) continue;
    const failure = coolingFailure(base, href);
    if (failure) throw failure;
  }
}

export function implicitAssetCooldown(url: URL, sliceIndex?: URL): number | undefined {
  const base = implicitFontCdnBases().get(baseKey(url));
  if (!base) return undefined;
  let until = Math.max(base.unavailable?.until ?? 0, base.invalid.get(url.href)?.until ?? 0);
  if (sliceIndex) {
    const root = new URL('./', sliceIndex).href;
    for (const [href, failure] of base.invalid) {
      if (href.startsWith(root)) until = Math.max(until, failure.until);
    }
  }
  return until > Date.now() ? until : undefined;
}

export async function loadImplicitFontAsset<T>(
  file: string, url: URL, pack: Pack, load: (bytes: ArrayBuffer, check: CheckFontAsset) => Promise<T>,
  checkDependencies?: () => void,
): Promise<T> {
  checkDependencies?.();
  const implicitBases = implicitFontCdnBases();
  const key = baseKey(url);
  let state = implicitBases.get(key);
  if (!state) {
    state = { transfer: Promise.resolve(), invalid: new Map(), warnedUntil: 0 };
    implicitBases.set(key, state);
  }
  const base = state;
  const cooling = coolingFailure(base, url.href);
  if (cooling) throw cooling;
  const fail = (cause: unknown, invalidBytes = false): ImplicitFontCdnError => {
    const now = Date.now();
    const until = now + COOLDOWN_MS;
    const error = new ImplicitFontCdnError(pack, cause);
    if (implicitBases.get(key) !== base) return error;
    if (failureScope(cause, invalidBytes) === 'base') base.unavailable = { error, until };
    else base.invalid.set(url.href, { error, until });
    if (base.warnedUntil <= now) {
      base.warnedUntil = until;
      console.warn(error.diagnostic.detail);
    }
    return error;
  };

  const pending = base.transfer.then(async () => {
    const cooling = coolingFailure(base, url.href);
    if (cooling) throw cooling;
    checkDependencies?.();
    try {
      const bytes = await readImplicitFontBytes(file, url);
      return { bytes, unavailable: base.unavailable };
    } catch (error) {
      throw fail(error);
    }
  });
  base.transfer = pending.then(() => {}, () => {});

  const { bytes, unavailable } = await pending;
  const data = await load(bytes, async (check) => {
    try { return await check(); }
    catch (error) { throw fail(error, true); }
  });
  if (base.unavailable === unavailable) {
    base.unavailable = undefined;
    const now = Date.now();
    if ([...base.invalid.values()].every(({ until }) => until <= now)) base.warnedUntil = 0;
  }
  return data;
}

async function readImplicitFontBytes(file: string, url: URL): Promise<ArrayBuffer> {
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
    try {
      if (signal.aborted) throw signal.reason;
      if (response.type === 'opaque' || response.type === 'error') throw new Error(`No readable response for bundled font ${file}`);
      if (!response.ok) throw new HttpFontError(file, response.status);
    } catch (error) {
      void response.body?.cancel().catch(() => {});
      throw error;
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
  implicitFontCdnBases().clear();
  setFontCdnConfiguration({ baseUrl, timeoutMs }, true);
}

/** @internal */
export function fontCdnConfiguration(): FontCdnConfiguration {
  return getFontCdnConfiguration() ?? { baseUrl: false, timeoutMs: 30_000 };
}
