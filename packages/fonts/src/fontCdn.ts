import { version as latinVersion } from '../../fonts-latin/package.json';
import { version as intlVersion } from '../../fonts-intl/package.json';
import { version as cjkVersion } from '../../fonts-cjk/package.json';
import type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';

export interface FontCdnConfiguration {
  baseUrl: string | false;
  timeoutMs: number;
  implicit?: true;
}

type FontGlobals = {
  window?: { document?: unknown };
  process?: { versions?: { node?: string } };
  Deno?: unknown;
};

const JSDELIVR = 'https://cdn.jsdelivr.net/npm/';
const IMPLICIT_FONT_CDN_TIMEOUT_MS = 3000;
const CONFIGURATION = Symbol.for('betteroffice.fonts.cdn');
const CONFIGURED_HERE = Symbol.for('betteroffice.fonts.cdn.configured-here');
const UNAVAILABLE = Symbol.for('betteroffice.fonts.cdn-unavailable');

type Pack = FontPackDiagnostic['pack'];
type Failure = { error: Error; diagnostic: FontPackDiagnostic };
type CdnScope = {
  [CONFIGURATION]?: FontCdnConfiguration;
  [CONFIGURED_HERE]?: true;
  [UNAVAILABLE]?: Map<Pack, Failure>;
};

const scope = globalThis as CdnScope & FontGlobals;

function isServer(globals: FontGlobals): boolean {
  return globals.process?.versions?.node !== undefined || globals.Deno !== undefined;
}

export function isBrowserFontPage(globals: FontGlobals): boolean {
  return globals.window?.document != null && !isServer(globals);
}

function configuredFontCdn(configuration: FontCdnConfiguration | undefined): FontCdnConfiguration | undefined {
  if (configuration && !(configuration.implicit && isServer(scope))) return { ...configuration };
}

export function effectiveFontCdnConfiguration(): FontCdnConfiguration | undefined {
  const configured = scope[CONFIGURATION];
  if (configured) return configuredFontCdn(configured);
  if (isBrowserFontPage(scope)) return { baseUrl: JSDELIVR, timeoutMs: IMPLICIT_FONT_CDN_TIMEOUT_MS, implicit: true };
}

export function setFontCdnConfiguration(configuration: FontCdnConfiguration | undefined): void {
  if (configuration === undefined) delete scope[CONFIGURATION];
  else scope[CONFIGURATION] = { ...configuration };
  delete scope[CONFIGURED_HERE];
}

export function workerFontCdnConfiguration(reply: { cdn?: FontCdnConfiguration } | undefined): FontCdnConfiguration | undefined {
  return configuredFontCdn(scope[CONFIGURED_HERE] || reply === undefined ? scope[CONFIGURATION] : reply.cdn);
}

export function fontCdnCacheKey(): string | undefined {
  if (!isServer(scope) && !isBrowserFontPage(scope)) return undefined;
  const configuration = effectiveFontCdnConfiguration();
  return configuration ? JSON.stringify([configuration.baseUrl, configuration.timeoutMs, configuration.implicit]) : '';
}

export function implicitCdnFailure(pack: Pack): Error | undefined {
  return scope[UNAVAILABLE]?.get(pack)?.error;
}

export function implicitCdnDiagnostics(): FontPackDiagnostic[] {
  return [...scope[UNAVAILABLE]?.values() ?? []].map(({ diagnostic }) => diagnostic);
}

export function markImplicitCdnUnavailable(pack: Pack, cause: unknown): Error {
  const unavailable = scope[UNAVAILABLE] ??= new Map();
  const failed = unavailable.get(pack);
  if (failed) return failed.error;
  const reason = cause instanceof Error ? cause.message : String(cause);
  const detail = `Cannot load @betteroffice/fonts-${pack} from cdn.jsdelivr.net: ${reason}. To opt out, install @betteroffice/fonts-${pack} or call configureFontCdn({ baseUrl: false }).`;
  const error = new Error(detail);
  unavailable.set(pack, { error, diagnostic: { pack, reason: 'cdn-unavailable', detail } });
  console.warn(detail);
  return error;
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
  setFontCdnConfiguration({ baseUrl, timeoutMs });
  scope[CONFIGURED_HERE] = true;
}

/** @internal */
export function fontCdnConfiguration(): FontCdnConfiguration {
  return effectiveFontCdnConfiguration() ?? { baseUrl: false, timeoutMs: 30_000 };
}
