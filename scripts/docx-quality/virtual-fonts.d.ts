declare module 'virtual:docx-quality-fonts' {
  export type BundledFontScript = import('../../packages/fonts/src/manifest').BundledFontScript;
  export type LastResortOffice = import('../../packages/fonts/src/manifest').LastResortOffice;
  export const createFontProvider: typeof import('../../packages/fonts/src/cdn').createFontProvider;
  export const BUNDLED_FONTS: typeof import('../../packages/fonts/src/manifest').BUNDLED_FONTS;
  export const BASE_LATIN_FONTS: import('../../packages/fonts/src/manifest').BundledFontFace[];
  export const resolveGlyphFallbackFace: typeof import('../../packages/fonts/src/manifest').resolveGlyphFallbackFace;
  export const resolveLastResortChoice: typeof import('../../packages/fonts/src/manifest').resolveLastResortChoice;
  export const resolveMetricCompatFace: typeof import('../../packages/fonts/src/manifest').resolveMetricCompatFace;
  export const resolveScriptFallbackFace: typeof import('../../packages/fonts/src/manifest').resolveScriptFallbackFace;
}
