/**
 * Bundled fonts for the BetterOffice engine: metric-compatible Latin faces
 * plus script-coverage faces for CJK and RTL text.
 *
 * This package ships the Latin and RTL binaries (25 faces, 7.9 MB). The five
 * CJK faces are 33 MB on their own and live in the optional
 * `@betteroffice/fonts-cjk` add-on; their manifest entries stay here (policy
 * belongs in one place) and {@link loadBundledFontBytes} reaches the bytes
 * through an optional dynamic import. Without the add-on installed, a CJK
 * loader rejects and the caller falls back — nothing else changes.
 *
 * Metric-compatible set — the open fonts the LibreOffice/ChromeOS ecosystem
 * uses as drop-in metric replacements for the MS core fonts (same advance
 * widths, so line breaks and pagination match even where glyph outlines
 * differ slightly):
 *
 * - Carlito          <-> Calibri
 * - Caladea          <-> Cambria
 * - Liberation Sans  <-> Arial (and the Helvetica alias)
 * - Liberation Serif <-> Times New Roman (and the Times alias)
 * - Liberation Mono  <-> Courier New (and the Courier alias)
 *
 * Script-coverage set — Noto faces that give the Rust text engine (and the
 * browser) real glyphs for scripts the Latin faces cannot cover:
 *
 * - Noto Sans SC / Noto Serif SC  <- SimHei/YaHei/DengXian, SimSun/FangSong/KaiTi, ...
 * - Noto Sans TC                  <- Microsoft JhengHei, PMingLiU, MingLiU, ...
 * - Noto Sans JP                  <- MS Gothic/Mincho, Meiryo, Yu Gothic/Mincho, ...
 * - Noto Sans KR                  <- Malgun Gothic, Gulim, Dotum, Batang, ...
 * - Noto Sans Hebrew / Noto Sans Arabic / Noto Naskh Arabic (script fallbacks
 *   only — Latin families like Arial/Times keep their Liberation mapping and
 *   the per-script fallback chain supplies Hebrew/Arabic glyph coverage)
 *
 * IMPORTANT — the CJK mappings are **coverage fallbacks first, metric
 * approximations second**. Unlike Carlito/Calibri, the Noto CJK faces do NOT
 * share advance widths with SimSun/MS Gothic/Malgun Gothic etc. (fullwidth
 * ideographs are uniformly 1.0 em everywhere, but proportional Latin runs and
 * line heights differ), so pagination of CJK documents approximates Word
 * rather than matching it. The CJK faces ship as **Regular only** — a bold
 * CJK request resolves to the Regular face and bold falls back through the
 * font chain (the engine may synthesize; advances stay Regular).
 *
 * The CJK binaries are static CFF (OTTO) Regulars from noto-cjk's SubsetOTF
 * distribution, NOT the google/fonts variable TTFs: those VFs default to the
 * Thin (wght=100) instance, and the Rust engine reads default-instance
 * advances while the browser applies wght=400 — same bytes, different
 * numbers. The static Regulars keep both sides identical (the Rust FontStore
 * parses CFF via skrifa).
 *
 * All other binaries are raw TTF (sfnt). Nothing is woff2, because the same
 * bytes must be consumed by two sides at once: the browser (registered via
 * `FontFace`, so DOM measurement uses these exact bytes) and the Rust/WASM
 * `FontStore` (which parses raw sfnt). Both consumers require identical bytes.
 *
 * Fonts are fetched lazily and same-origin by default (`new URL(...,
 * import.meta.url)` so bundlers emit the assets); a host that wants a CDN
 * passes an explicit `baseUrl`. Importing this module performs no network
 * activity; zero-click fetches on package import are forbidden (see the
 * security section in the repo security guidelines).
 */

/** Script bucket a bundled face provides glyph coverage for. */
export type BundledFontScript = 'cjk-sc' | 'cjk-tc' | 'cjk-jp' | 'cjk-kr' | 'arabic' | 'hebrew';

/** One bundled font binary and the Word font(s) it stands in for. */
export interface BundledFontFace {
  /** Family name as it appears in the font's own name table, e.g. "Carlito". */
  family: string;
  /**
   * The Word font this face substitutes, e.g. "Calibri". For the Latin set
   * this is a true metric match; for the CJK set it is a coverage fallback
   * (see the module doc). Absent on the pure script-fallback faces (RTL).
   */
  metricCompatWith?: string;
  weight: 400 | 700;
  style: 'normal' | 'italic';
  /** Asset filename under this package's `assets/` directory. */
  file: string;
  /** Present on faces that serve as per-script coverage fallbacks. */
  script?: BundledFontScript;
}

function familyFaces(
  family: string,
  metricCompatWith: string,
  fileBase: string
): BundledFontFace[] {
  return [
    { family, metricCompatWith, weight: 400, style: 'normal', file: `${fileBase}-Regular.ttf` },
    { family, metricCompatWith, weight: 700, style: 'normal', file: `${fileBase}-Bold.ttf` },
    { family, metricCompatWith, weight: 400, style: 'italic', file: `${fileBase}-Italic.ttf` },
    { family, metricCompatWith, weight: 700, style: 'italic', file: `${fileBase}-BoldItalic.ttf` },
  ];
}

/**
 * The complete manifest of bundled faces. Single source of truth: the
 * metric-compat and script-fallback resolution below is derived from this
 * list, never duplicated.
 *
 * Order matters within a script bucket: `resolveScriptFallbackFace` prefers
 * earlier entries on ties, so the sans face of each script comes first.
 */
export const BUNDLED_FONTS: BundledFontFace[] = [
  ...familyFaces('Carlito', 'Calibri', 'Carlito'),
  ...familyFaces('Caladea', 'Cambria', 'Caladea'),
  ...familyFaces('Liberation Sans', 'Arial', 'LiberationSans'),
  ...familyFaces('Liberation Serif', 'Times New Roman', 'LiberationSerif'),
  ...familyFaces('Liberation Mono', 'Courier New', 'LiberationMono'),

  // RTL script fallbacks. No metricCompatWith: Hebrew/Arabic documents mostly
  // name Latin families (Arial, Times New Roman, ...) whose mapping stays with
  // the Liberation faces; these faces ride the per-script fallback chain.
  {
    family: 'Noto Sans Hebrew',
    weight: 400,
    style: 'normal',
    file: 'NotoSansHebrew-Regular.ttf',
    script: 'hebrew',
  },
  {
    family: 'Noto Sans Hebrew',
    weight: 700,
    style: 'normal',
    file: 'NotoSansHebrew-Bold.ttf',
    script: 'hebrew',
  },
  {
    family: 'Noto Sans Arabic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansArabic-Regular.ttf',
    script: 'arabic',
  },
  {
    family: 'Noto Sans Arabic',
    weight: 700,
    style: 'normal',
    file: 'NotoSansArabic-Bold.ttf',
    script: 'arabic',
  },
  {
    family: 'Noto Naskh Arabic',
    weight: 400,
    style: 'normal',
    file: 'NotoNaskhArabic-Regular.ttf',
    script: 'arabic',
  },

  // CJK coverage faces (Regular-only statics; see the module doc). The sans
  // face precedes the serif face of the same script bucket on purpose.
  {
    family: 'Noto Sans SC',
    metricCompatWith: 'Microsoft YaHei',
    weight: 400,
    style: 'normal',
    file: 'NotoSansSC-Regular.otf',
    script: 'cjk-sc',
  },
  {
    family: 'Noto Serif SC',
    metricCompatWith: 'SimSun',
    weight: 400,
    style: 'normal',
    file: 'NotoSerifSC-Regular.otf',
    script: 'cjk-sc',
  },
  {
    family: 'Noto Sans TC',
    metricCompatWith: 'Microsoft JhengHei',
    weight: 400,
    style: 'normal',
    file: 'NotoSansTC-Regular.otf',
    script: 'cjk-tc',
  },
  {
    family: 'Noto Sans JP',
    metricCompatWith: 'MS Gothic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansJP-Regular.otf',
    script: 'cjk-jp',
  },
  {
    family: 'Noto Sans KR',
    metricCompatWith: 'Malgun Gothic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansKR-Regular.otf',
    script: 'cjk-kr',
  },
];

/**
 * Alternate Word font names that resolve to the same bundled face as a
 * covered Word family (keys and values lowercase). Kept separate from the
 * manifest: these are aliases of the *Word-side* name, resolved through
 * `BUNDLED_FONTS`.
 *
 * The CJK alias set mirrors the CJK table in core's `utils/fontResolver.ts`
 * (both romanized and native spellings; native full-width Latin lowercases
 * too, e.g. `ＭＳ ゴシック` -> `ｍｓ ゴシック`). Where fontResolver picks a
 * serif Noto family this package does not vendor (Noto Serif TC/JP/KR), the
 * alias points at the vendored sans face of the same region — coverage
 * first.
 */
const WORD_FAMILY_ALIASES: Record<string, string> = {
  helvetica: 'arial',
  times: 'times new roman',
  courier: 'courier new',

  // Simplified Chinese — sans
  simhei: 'microsoft yahei',
  dengxian: 'microsoft yahei',
  微软雅黑: 'microsoft yahei',
  黑体: 'microsoft yahei',
  等线: 'microsoft yahei',
  // Simplified Chinese — serif
  nsimsun: 'simsun',
  fangsong: 'simsun',
  kaiti: 'simsun',
  宋体: 'simsun',
  仿宋: 'simsun',
  楷体: 'simsun',
  // Traditional Chinese (the Ming/Kai serif families map to the sans face —
  // Noto Serif TC is not vendored)
  微軟正黑體: 'microsoft jhenghei',
  pmingliu: 'microsoft jhenghei',
  mingliu: 'microsoft jhenghei',
  'dfkai-sb': 'microsoft jhenghei',
  新細明體: 'microsoft jhenghei',
  細明體: 'microsoft jhenghei',
  標楷體: 'microsoft jhenghei',
  // Japanese (the Mincho serif families map to the sans face — Noto Serif JP
  // is not vendored)
  'ms pgothic': 'ms gothic',
  meiryo: 'ms gothic',
  'yu gothic': 'ms gothic',
  'ｍｓ ゴシック': 'ms gothic',
  'ｍｓ ｐゴシック': 'ms gothic',
  メイリオ: 'ms gothic',
  游ゴシック: 'ms gothic',
  'ms mincho': 'ms gothic',
  'ms pmincho': 'ms gothic',
  'yu mincho': 'ms gothic',
  'ｍｓ 明朝': 'ms gothic',
  'ｍｓ ｐ明朝': 'ms gothic',
  游明朝: 'ms gothic',
  // Korean (Batang/Gungsuh serif map to the sans face — Noto Serif KR is not
  // vendored)
  '맑은 고딕': 'malgun gothic',
  gulim: 'malgun gothic',
  dotum: 'malgun gothic',
  batang: 'malgun gothic',
  gungsuh: 'malgun gothic',
  굴림: 'malgun gothic',
  돋움: 'malgun gothic',
  바탕: 'malgun gothic',
  궁서: 'malgun gothic',
};

const metricCompatByWordFamily = new Map<string, string>();
for (const face of BUNDLED_FONTS) {
  if (face.metricCompatWith !== undefined) {
    metricCompatByWordFamily.set(face.metricCompatWith.toLowerCase(), face.family);
  }
}

/**
 * Resolve a Word font name (case-insensitive) to the bundled substitute
 * family, e.g. `"calibri"` -> `"Carlito"`, `"SimSun"` -> `"Noto Serif SC"`.
 * Returns `undefined` when no bundled font covers the name.
 */
export function resolveMetricCompatFamily(wordFamily: string): string | undefined {
  const key = wordFamily.trim().toLowerCase();
  return metricCompatByWordFamily.get(WORD_FAMILY_ALIASES[key] ?? key);
}

/**
 * Resolve a Word font name plus style request to a concrete bundled face.
 * Exact (weight, style) match first; families that only ship a Regular (the
 * CJK set) fall back to it — bold then falls back through the font chain,
 * mirroring how the measurement registry treats embedded faces.
 */
export function resolveMetricCompatFace(
  wordFamily: string,
  bold: boolean,
  italic: boolean
): BundledFontFace | undefined {
  const family = resolveMetricCompatFamily(wordFamily);
  if (!family) return undefined;
  const faces = BUNDLED_FONTS.filter((f) => f.family === family);
  const weight = bold ? 700 : 400;
  const style = italic ? 'italic' : 'normal';
  return (
    faces.find((f) => f.weight === weight && f.style === style) ??
    faces.find((f) => f.weight === 400 && f.style === 'normal')
  );
}

/**
 * Pick the bundled face that provides glyph coverage for a script bucket.
 * Preference order: exact (weight, style) -> same weight upright -> the
 * script's Regular -> the first face of the bucket. Ties resolve to the
 * earlier manifest entry, i.e. the sans face (Noto Naskh Arabic is reachable
 * by requesting it as a family, not through the script fallback).
 */
export function resolveScriptFallbackFace(
  script: BundledFontScript,
  bold: boolean,
  italic: boolean
): BundledFontFace | undefined {
  const faces = BUNDLED_FONTS.filter((f) => f.script === script);
  if (faces.length === 0) return undefined;
  const weight = bold ? 700 : 400;
  const style = italic ? 'italic' : 'normal';
  return (
    faces.find((f) => f.weight === weight && f.style === style) ??
    faces.find((f) => f.weight === weight && f.style === 'normal') ??
    faces.find((f) => f.weight === 400 && f.style === 'normal') ??
    faces[0]
  );
}

/**
 * Whether a Word family name reads as a serif — decides only which
 * always-available base face measures a truly-unknown font (Liberation Serif
 * vs Liberation Sans). Mirrors the serif branch of core's `detectFontCategory`
 * (`utils/fontResolver.ts`) so an unmapped serif name lands on a serif base.
 * Deliberately coarse: this feeds the last-resort face pick, nothing else.
 */
function looksSerif(family: string): boolean {
  const lower = family.toLowerCase();
  return (
    lower.includes('times') ||
    lower.includes('georgia') ||
    lower.includes('garamond') ||
    lower.includes('palatino') ||
    lower.includes('baskerville') ||
    lower.includes('bodoni') ||
    lower.includes('cambria') ||
    lower.includes('minion') ||
    lower.includes('mincho') ||
    lower.includes('明朝') ||
    lower.includes('明體') ||
    lower.includes('宋') ||
    lower.includes('ming') ||
    lower.includes('song') ||
    lower.includes('serif')
  );
}

/**
 * The always-available last-resort base face for ANY Word family. Broad-
 * coverage Latin: Liberation Serif for serif-looking names, Liberation Sans
 * otherwise. Both bundled Liberation families ship the full
 * Regular/Bold/Italic/BoldItalic set, so the exact (bold, italic) style always
 * resolves and this NEVER returns undefined.
 *
 * This is the terminal link of the measurement font chain (see the font
 * registry's chain contract): appended after the embedded and metric-compatible
 * faces so a run whose family has no embedded/bundled match still has real font
 * bytes to measure with, keeping it on the native (Rust) measurement path
 * instead of routing the whole block to browser `measureText`. The measured
 * metrics are Liberation's, not the requested font's — an accepted width
 * divergence for a truly-unknown font, in exchange for staying native. Latin
 * coverage only; per-script coverage (CJK/RTL) rides
 * {@link resolveScriptFallbackFace}, appended separately by the registry.
 */
export function resolveLastResortFace(
  family: string,
  bold: boolean,
  italic: boolean
): BundledFontFace {
  // Liberation Sans/Serif always ship the full four-face set, so the
  // metric-compat resolution is guaranteed to return a face here.
  const base = looksSerif(family) ? 'Times New Roman' : 'Arial';
  return resolveMetricCompatFace(base, bold, italic)!;
}


// Per-file LITERAL asset URLs. Bundlers only statically resolve `new URL()`
// when the specifier is a string literal — a template expression works under
// Vite's directory glob but collapses to a single (wrong) asset under
// webpack/Turbopack. Every face this package ships must have a row here; the
// CJK faces resolve through `@betteroffice/fonts-cjk` instead.
const FONT_ASSET_URLS: Record<string, () => URL> = {
  'Caladea-Bold.ttf': () => new URL('../assets/Caladea-Bold.ttf', import.meta.url),
  'Caladea-BoldItalic.ttf': () => new URL('../assets/Caladea-BoldItalic.ttf', import.meta.url),
  'Caladea-Italic.ttf': () => new URL('../assets/Caladea-Italic.ttf', import.meta.url),
  'Caladea-Regular.ttf': () => new URL('../assets/Caladea-Regular.ttf', import.meta.url),
  'Carlito-Bold.ttf': () => new URL('../assets/Carlito-Bold.ttf', import.meta.url),
  'Carlito-BoldItalic.ttf': () => new URL('../assets/Carlito-BoldItalic.ttf', import.meta.url),
  'Carlito-Italic.ttf': () => new URL('../assets/Carlito-Italic.ttf', import.meta.url),
  'Carlito-Regular.ttf': () => new URL('../assets/Carlito-Regular.ttf', import.meta.url),
  'LiberationMono-Bold.ttf': () => new URL('../assets/LiberationMono-Bold.ttf', import.meta.url),
  'LiberationMono-BoldItalic.ttf': () => new URL('../assets/LiberationMono-BoldItalic.ttf', import.meta.url),
  'LiberationMono-Italic.ttf': () => new URL('../assets/LiberationMono-Italic.ttf', import.meta.url),
  'LiberationMono-Regular.ttf': () => new URL('../assets/LiberationMono-Regular.ttf', import.meta.url),
  'LiberationSans-Bold.ttf': () => new URL('../assets/LiberationSans-Bold.ttf', import.meta.url),
  'LiberationSans-BoldItalic.ttf': () => new URL('../assets/LiberationSans-BoldItalic.ttf', import.meta.url),
  'LiberationSans-Italic.ttf': () => new URL('../assets/LiberationSans-Italic.ttf', import.meta.url),
  'LiberationSans-Regular.ttf': () => new URL('../assets/LiberationSans-Regular.ttf', import.meta.url),
  'LiberationSerif-Bold.ttf': () => new URL('../assets/LiberationSerif-Bold.ttf', import.meta.url),
  'LiberationSerif-BoldItalic.ttf': () => new URL('../assets/LiberationSerif-BoldItalic.ttf', import.meta.url),
  'LiberationSerif-Italic.ttf': () => new URL('../assets/LiberationSerif-Italic.ttf', import.meta.url),
  'LiberationSerif-Regular.ttf': () => new URL('../assets/LiberationSerif-Regular.ttf', import.meta.url),
  'NotoNaskhArabic-Regular.ttf': () => new URL('../assets/NotoNaskhArabic-Regular.ttf', import.meta.url),
  'NotoSansArabic-Bold.ttf': () => new URL('../assets/NotoSansArabic-Bold.ttf', import.meta.url),
  'NotoSansArabic-Regular.ttf': () => new URL('../assets/NotoSansArabic-Regular.ttf', import.meta.url),
  'NotoSansHebrew-Bold.ttf': () => new URL('../assets/NotoSansHebrew-Bold.ttf', import.meta.url),
  'NotoSansHebrew-Regular.ttf': () => new URL('../assets/NotoSansHebrew-Regular.ttf', import.meta.url),
};

/** Where a face's bytes are served from. */
export interface FontAssetOptions {
  /**
   * Serve the faces from here instead of this package's own assets. Same-origin
   * (the package's own assets) is the default on purpose: a CDN default would
   * leak document-font usage to a third party and break offline and strict-CSP
   * deployments. A trailing slash is optional; a relative base resolves against
   * the document.
   */
  baseUrl?: string | URL;
}

let cjkAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;

/**
 * Asset URLs of the optional CJK add-on, or `undefined` when it is not
 * installed. Memoized including the miss, so an absent package costs one
 * failed import per session.
 */
function loadCjkAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  cjkAssetUrls ??= import('@betteroffice/fonts-cjk').then(
    (module) => module.CJK_FONT_ASSET_URLS,
    () => undefined
  );
  return cjkAssetUrls;
}

function assetBase(baseUrl: string | URL): string {
  const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
  return href.endsWith('/') ? href : `${href}/`;
}

async function assetUrl(file: string, baseUrl: string | URL | undefined): Promise<URL> {
  if (baseUrl !== undefined) {
    const base = assetBase(baseUrl);
    return new URL(file, typeof location === 'undefined' ? base : new URL(base, location.href));
  }
  const local = FONT_ASSET_URLS[file];
  if (local) return local();
  const cjk = await loadCjkAssetUrls();
  const resolveCjk = cjk?.[file];
  if (resolveCjk) return resolveCjk();
  if (!cjk) {
    throw new Error(
      `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`
    );
  }
  throw new Error(`Unknown bundled font asset: ${file}`);
}

const bytesCache = new Map<string, Promise<ArrayBuffer>>();

/**
 * Lazily fetch the raw sfnt bytes for a face. The fetch is same-origin by
 * default: the asset URL is derived with `new URL(..., import.meta.url)` so
 * bundlers (Vite) emit the file and serve it alongside the module. Pass
 * `baseUrl` to serve the faces from a CDN instead.
 *
 * Results are cached per (face, base URL) — the same promise is returned for
 * concurrent callers, and the same `ArrayBuffer` instance is handed to every
 * consumer, so byte-identity lets registries deduplicate registrations. A
 * failed fetch is evicted so it can be retried.
 */
export function loadBundledFontBytes(
  face: BundledFontFace,
  options?: FontAssetOptions
): Promise<ArrayBuffer> {
  const key = `${options?.baseUrl === undefined ? '' : assetBase(options.baseUrl)}\n${face.file}`;
  const cached = bytesCache.get(key);
  if (cached) return cached;
  const promise = assetUrl(face.file, options?.baseUrl)
    .then((url) => fetch(url))
    .then((response) => {
      if (!response.ok) {
        throw new Error(`Failed to fetch bundled font ${face.file}: HTTP ${response.status}`);
      }
      return response.arrayBuffer();
    });
  promise.catch(() => {
    if (bytesCache.get(key) === promise) bytesCache.delete(key);
  });
  bytesCache.set(key, promise);
  return promise;
}

const registeredFaces = new Map<string, Promise<void>>();

/**
 * Register a face with the DOM via the `FontFace` API under an explicit CSS
 * family name (defaults to the face's real family), so browser measurement
 * uses the SAME bytes the wasm-side `FontStore` receives. Idempotent per
 * (cssFamily, weight, style); a failed registration is evicted so it can be
 * retried. Resolves as a no-op in non-DOM environments.
 */
export function registerBundledFontFace(
  face: BundledFontFace,
  cssFamily?: string,
  options?: FontAssetOptions
): Promise<void> {
  if (
    typeof document === 'undefined' ||
    typeof FontFace === 'undefined' ||
    document.fonts === undefined
  ) {
    return Promise.resolve();
  }
  const family = cssFamily ?? face.family;
  const key = `${family}|${face.weight}|${face.style}`;
  const existing = registeredFaces.get(key);
  if (existing) return existing;
  const promise = (async () => {
    const bytes = await loadBundledFontBytes(face, options);
    // The family name goes through the FontFace API as a value, never
    // interpolated into a CSS string, so there is no CSS-injection sink here.
    const fontFace = new FontFace(family, bytes, {
      weight: String(face.weight),
      style: face.style,
    });
    await fontFace.load();
    document.fonts.add(fontFace);
  })();
  promise.catch(() => {
    if (registeredFaces.get(key) === promise) registeredFaces.delete(key);
  });
  registeredFaces.set(key, promise);
  return promise;
}

/**
 * A resolver of bundled font bytes, shaped exactly like the measurement
 * engine's `BundledFontProvider` (`@betteroffice/docx`'s
 * `layout/measure/fontRegistry`). Declared structurally rather than imported
 * so this package stays independent of the engine packages.
 */
export interface BundledFontSource {
  /** Bundled metric-compatible face for a Word family, or undefined. */
  resolve(family: string, bold: boolean, italic: boolean): (() => Promise<ArrayBuffer>) | undefined;
  /** Per-script coverage face (CJK/RTL), or undefined. */
  resolveScriptFallback(
    script: BundledFontScript,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
  /** The always-available last-resort base face — never undefined. */
  resolveLastResort(family: string, bold: boolean, italic: boolean): () => Promise<ArrayBuffer>;
}

/**
 * Build the provider the measurement engine consumes. This is what
 * `@betteroffice/docx` instantiates when a host injects no provider of its
 * own, and what a host should use when it wants the bundled faces served from
 * somewhere other than its own origin.
 *
 * Loaders are lazy: nothing is fetched until a document actually needs a face.
 */
export function createFontProvider(options?: FontAssetOptions): BundledFontSource {
  const load = (face: BundledFontFace) => () => loadBundledFontBytes(face, options);
  return {
    resolve(family, bold, italic) {
      const face = resolveMetricCompatFace(family, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveScriptFallback(script, bold, italic) {
      const face = resolveScriptFallbackFace(script, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveLastResort(family, bold, italic) {
      return load(resolveLastResortFace(family, bold, italic));
    },
  };
}
