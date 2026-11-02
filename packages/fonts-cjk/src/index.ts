export const CJK_FONT_ASSET_URLS: Record<string, () => URL> = {
  'NotoSansJP-Regular.otf.zst': () => new URL('../assets/NotoSansJP-Regular.otf.zst', import.meta.url),
  'NotoSansKR-Regular.otf.zst': () => new URL('../assets/NotoSansKR-Regular.otf.zst', import.meta.url),
  'NotoSansSC-Regular.otf.zst': () => new URL('../assets/NotoSansSC-Regular.otf.zst', import.meta.url),
  'NotoSansTC-Regular.otf.zst': () => new URL('../assets/NotoSansTC-Regular.otf.zst', import.meta.url),
  'NotoSerifSC-Regular.otf.zst': () => new URL('../assets/NotoSerifSC-Regular.otf.zst', import.meta.url),
};
