/** Optional static CJK assets for `@betteroffice/fonts`, split to reduce install size. */

/** Literal asset URLs required for webpack and Turbopack emission. */
export const CJK_FONT_ASSET_URLS: Record<string, () => URL> = {
  'NotoSansJP-Regular.otf': () => new URL('../assets/NotoSansJP-Regular.otf', import.meta.url),
  'NotoSansKR-Regular.otf': () => new URL('../assets/NotoSansKR-Regular.otf', import.meta.url),
  'NotoSansSC-Regular.otf': () => new URL('../assets/NotoSansSC-Regular.otf', import.meta.url),
  'NotoSansTC-Regular.otf': () => new URL('../assets/NotoSansTC-Regular.otf', import.meta.url),
  'NotoSerifSC-Regular.otf': () => new URL('../assets/NotoSerifSC-Regular.otf', import.meta.url),
};
