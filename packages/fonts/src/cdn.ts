import { version as fontsVersion } from '../package.json';
import { version as fontsCjkVersion } from '../../fonts-cjk/package.json';
import { version as fontsLatinVersion } from '../../fonts-latin/package.json';
import { version as fontsIntlVersion } from '../../fonts-intl/package.json';
import { loadFontBytes } from './bytes';
import type { FontDecoder } from './decode';
import { loadPackedFontBytes, isLatinFont } from './packed';
import { fontProvider, type BundledFontSource } from './provider';

export interface CdnFontOptions {
  baseUrl?: string | URL;
  cjkBaseUrl?: string | URL;
}

const BASE_URL = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-latin@${fontsLatinVersion}/assets/`;
const CJK_BASE_URL = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${fontsCjkVersion}/assets/`;
const INTL_BASE_URL = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-intl@${fontsIntlVersion}/assets/`;
const DECODER_URL = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${fontsVersion}/dist/decoder/fontpack.wasm`;

function assetBase(value: string | URL): URL {
  const href = String(value);
  const base = href.endsWith('/') ? href : `${href}/`;
  return new URL(
    base,
    typeof location === 'undefined' ? undefined : location.href,
  );
}

/** Loads pinned CDN fonts without bundling font assets or add-ons. */
export function createFontProvider(
  options: CdnFontOptions = {},
  runtime?: { decoder?: FontDecoder },
): BundledFontSource {
  const base = assetBase(options.baseUrl ?? BASE_URL);
  const cjkBase = assetBase(
    options.cjkBaseUrl ?? options.baseUrl ?? CJK_BASE_URL,
  );
  const intlBase = assetBase(options.baseUrl ?? INTL_BASE_URL);
  const decoderUrl = options.baseUrl === undefined
    ? new URL(DECODER_URL)
    : new URL('decoder/fontpack.wasm', base);
  return fontProvider((face) => {
    if (face.file.endsWith('.ttf.zst')) {
      return loadPackedFontBytes(
        face,
        async (dependency) => new URL(dependency.file, isLatinFont(face.file) ? base : intlBase),
        30_000,
        decoderUrl,
        undefined,
        runtime,
      );
    }
    return loadFontBytes(
      face,
      new URL(face.file, face.script?.startsWith('cjk-') ? cjkBase : base),
      30_000,
    );
  });
}
