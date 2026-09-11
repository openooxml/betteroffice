import { loadFontBytes } from './bytes';
import { fontProvider, type BundledFontSource } from './provider';

export interface CdnFontOptions {
  baseUrl?: string | URL;
  cjkBaseUrl?: string | URL;
}

const BASE_URL =
  'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.1.0/assets/';
const CJK_BASE_URL =
  'https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@0.1.0/assets/';

function assetBase(value: string | URL): URL {
  const href = String(value);
  const base = href.endsWith('/') ? href : `${href}/`;
  return new URL(
    base,
    typeof location === 'undefined' ? undefined : location.href,
  );
}

/** Loads pinned CDN fonts without bundling font assets or the CJK package. */
export function createFontProvider(
  options: CdnFontOptions = {},
): BundledFontSource {
  const base = assetBase(options.baseUrl ?? BASE_URL);
  const cjkBase = assetBase(
    options.cjkBaseUrl ?? options.baseUrl ?? CJK_BASE_URL,
  );
  return fontProvider((face) =>
    loadFontBytes(
      face,
      new URL(face.file, face.script?.startsWith('cjk-') ? cjkBase : base),
      30_000,
    ),
  );
}
