import { version as latinVersion } from '../../fonts-latin/package.json';
import { version as intlVersion } from '../../fonts-intl/package.json';
import { version as cjkVersion } from '../../fonts-cjk/package.json';
import { getFontCdnConfiguration, setFontCdnConfiguration, type FontCdnConfiguration } from '../../../shared/office-session/fontCdn';

export type { FontCdnConfiguration } from '../../../shared/office-session/fontCdn';

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
  return new URL(`https://cdn.jsdelivr.net/npm/@betteroffice/fonts-${pack.pack}@${version}/assets/${file}`);
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
}

/** @internal */
export function fontCdnConfiguration(): FontCdnConfiguration {
  return getFontCdnConfiguration() ?? { baseUrl: false, timeoutMs: 30_000 };
}
