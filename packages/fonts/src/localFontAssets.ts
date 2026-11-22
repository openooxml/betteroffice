import { BASE_LATIN_FONT_ASSET_URLS } from './base-latin';
import { fontCdnAssetUrl, fontCdnConfiguration } from './fontCdn';
import type { BundledFontFace } from './manifest';
import { isBaseLatinFont, packFor } from './packed';

export const localFontAssets = {
  baseLatin: BASE_LATIN_FONT_ASSET_URLS,
  resolveBase(baseUrl: string | URL): URL {
    const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
    const base = href.endsWith('/') ? href : `${href}/`;
    try {
      return new URL(base, globalThis.location?.href);
    } catch {
      throw new TypeError(`Font baseUrl must be absolute when no browser location exists: ${base}`);
    }
  },
  cdn() {
    const { baseUrl, timeoutMs } = fontCdnConfiguration();
    return {
      base: (face: BundledFontFace) => baseUrl ? fontCdnAssetUrl(packFor(face), '', baseUrl).href : '',
      url(face: BundledFontFace, file: string): URL | undefined {
        if (baseUrl === false || isBaseLatinFont(face.file)) return undefined;
        const pack = packFor(face);
        if (!pack.faces.some((entry) => entry.file === face.file)) return undefined;
        return fontCdnAssetUrl(pack, file, baseUrl);
      },
      timeoutMs: baseUrl ? timeoutMs : undefined,
    };
  },
};
