export const BASE_LATIN_FONT_ASSET_URLS: Record<string, () => URL> = {
  'BetterOfficeBaseLatin-Regular.ttf.zst': () => new URL('../assets/BetterOfficeBaseLatin-Regular.ttf.zst', import.meta.url),
  'BetterOfficeBaseLatin-Bold.ttf.zst': () => new URL('../assets/BetterOfficeBaseLatin-Bold.ttf.zst', import.meta.url),
  'BetterOfficeBaseLatin-Italic.ttf.zst': () => new URL('../assets/BetterOfficeBaseLatin-Italic.ttf.zst', import.meta.url),
  'BetterOfficeBaseLatin-BoldItalic.ttf.zst': () => new URL('../assets/BetterOfficeBaseLatin-BoldItalic.ttf.zst', import.meta.url),
};
