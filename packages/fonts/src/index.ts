/** Deterministic bundled-font resolution with lazy loading and optional CJK assets. */

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
  byteLength: number;
  /** Present on faces that serve as per-script coverage fallbacks. */
  script?: BundledFontScript;
}

function familyFaces(
  family: string,
  metricCompatWith: string,
  fileBase: string,
  byteLengths: readonly [number, number, number, number]
): BundledFontFace[] {
  return [
    {
      family,
      metricCompatWith,
      weight: 400,
      style: 'normal',
      file: `${fileBase}-Regular.ttf`,
      byteLength: byteLengths[0],
    },
    {
      family,
      metricCompatWith,
      weight: 700,
      style: 'normal',
      file: `${fileBase}-Bold.ttf`,
      byteLength: byteLengths[1],
    },
    {
      family,
      metricCompatWith,
      weight: 400,
      style: 'italic',
      file: `${fileBase}-Italic.ttf`,
      byteLength: byteLengths[2],
    },
    {
      family,
      metricCompatWith,
      weight: 700,
      style: 'italic',
      file: `${fileBase}-BoldItalic.ttf`,
      byteLength: byteLengths[3],
    },
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
  ...familyFaces('Carlito', 'Calibri', 'Carlito', [628032, 682468, 615236, 808508]),
  ...familyFaces('Caladea', 'Cambria', 'Caladea', [81600, 84492, 83780, 83356]),
  ...familyFaces('Liberation Sans', 'Arial', 'LiberationSans', [
    410712, 414456, 415816, 408996,
  ]),
  ...familyFaces('Liberation Serif', 'Times New Roman', 'LiberationSerif', [
    393576, 370096, 375632, 376772,
  ]),
  ...familyFaces('Liberation Mono', 'Courier New', 'LiberationMono', [
    319508, 307996, 281536, 284068,
  ]),

  // RTL script fallbacks. No metricCompatWith: Hebrew/Arabic documents mostly
  // name Latin families (Arial, Times New Roman, ...) whose mapping stays with
  // the Liberation faces; these faces ride the per-script fallback chain.
  {
    family: 'Noto Sans Hebrew',
    weight: 400,
    style: 'normal',
    file: 'NotoSansHebrew-Regular.ttf',
    byteLength: 26860,
    script: 'hebrew',
  },
  {
    family: 'Noto Sans Hebrew',
    weight: 700,
    style: 'normal',
    file: 'NotoSansHebrew-Bold.ttf',
    byteLength: 26860,
    script: 'hebrew',
  },
  {
    family: 'Noto Sans Arabic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansArabic-Regular.ttf',
    byteLength: 234892,
    script: 'arabic',
  },
  {
    family: 'Noto Sans Arabic',
    weight: 700,
    style: 'normal',
    file: 'NotoSansArabic-Bold.ttf',
    byteLength: 261460,
    script: 'arabic',
  },
  {
    family: 'Noto Naskh Arabic',
    weight: 400,
    style: 'normal',
    file: 'NotoNaskhArabic-Regular.ttf',
    byteLength: 247336,
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
    byteLength: 8331336,
    script: 'cjk-sc',
  },
  {
    family: 'Noto Serif SC',
    metricCompatWith: 'SimSun',
    weight: 400,
    style: 'normal',
    file: 'NotoSerifSC-Regular.otf',
    byteLength: 11625800,
    script: 'cjk-sc',
  },
  {
    family: 'Noto Sans TC',
    metricCompatWith: 'Microsoft JhengHei',
    weight: 400,
    style: 'normal',
    file: 'NotoSansTC-Regular.otf',
    byteLength: 5683368,
    script: 'cjk-tc',
  },
  {
    family: 'Noto Sans JP',
    metricCompatWith: 'MS Gothic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansJP-Regular.otf',
    byteLength: 4533028,
    script: 'cjk-jp',
  },
  {
    family: 'Noto Sans KR',
    metricCompatWith: 'Malgun Gothic',
    weight: 400,
    style: 'normal',
    file: 'NotoSansKR-Regular.otf',
    byteLength: 4644748,
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

export interface FontAssetOptions {
  /**
   * Asset root. The default stays same-origin for privacy, offline use, and
   * strict CSP. Relative roots pin to the current document; server roots must
   * be absolute.
   */
  baseUrl?: string | URL;
}

let cjkAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;

async function importCjkAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  // Keep the SYNTACTIC try/catch with the await as its direct body. Rewriting
  // this as `import(…).catch()` or a two-argument `.then()` makes webpack (and
  // so `next build`) fail hard on the absent optional peer, and esbuild starts
  // resolving the specifier eagerly the moment it stops being that direct body.
  try {
    return (await import('@betteroffice/fonts-cjk')).CJK_FONT_ASSET_URLS;
  } catch {
    return undefined;
  }
}

/** Shares in-flight or successful CJK imports while leaving misses retryable. */
function loadCjkAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  if (cjkAssetUrls === undefined) {
    const promise = importCjkAssetUrls();
    promise.then((urls) => {
      if (urls === undefined && cjkAssetUrls === promise) cjkAssetUrls = undefined;
    });
    cjkAssetUrls = promise;
  }
  return cjkAssetUrls;
}

function assetBase(baseUrl: string | URL): string {
  const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
  return href.endsWith('/') ? href : `${href}/`;
}

function resolvedAssetBase(baseUrl: string | URL): URL {
  const base = assetBase(baseUrl);
  try {
    return typeof location === 'undefined' ? new URL(base) : new URL(base, location.href);
  } catch {
    throw new TypeError(`Font baseUrl must be absolute when no browser location exists: ${base}`);
  }
}

async function assetUrl(file: string, baseUrl: URL | undefined): Promise<URL> {
  if (baseUrl !== undefined) return new URL(file, baseUrl);
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

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}
interface NodeUrlLike {
  fileURLToPath(url: string): string;
}

function builtinModule<T>(name: string): T | undefined {
  const proc = (
    globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }
  ).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

/** Node fetch cannot read file: package assets, so server imports use built-in I/O. */
function readFileAsset(url: URL): ArrayBuffer | undefined {
  if (url.protocol !== 'file:') return undefined;
  const fs = builtinModule<NodeFsLike>('node:fs');
  const nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) return undefined;
  try {
    // `.href` and not the URL object: fileURLToPath brand-checks its argument.
    const bytes = fs.readFileSync(nodeUrl.fileURLToPath(url.href));
    return bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength
    ) as ArrayBuffer;
  } catch {
    return undefined;
  }
}

const bytesCache = new Map<string, Promise<ArrayBuffer>>();

function validateByteLength(face: BundledFontFace, bytes: ArrayBuffer): ArrayBuffer {
  if (bytes.byteLength !== face.byteLength) {
    throw new Error(
      `Bundled font ${face.file} has ${bytes.byteLength} bytes; expected ${face.byteLength}`
    );
  }
  return bytes;
}

/**
 * Lazily loads raw sfnt bytes from package assets or `baseUrl`. Literal
 * `import.meta.url` assets keep defaults same-origin; loads share buffer
 * identity per face/base and failures remain retryable.
 */
export function loadBundledFontBytes(
  face: BundledFontFace,
  options?: FontAssetOptions
): Promise<ArrayBuffer> {
  const baseUrl = options?.baseUrl === undefined ? undefined : resolvedAssetBase(options.baseUrl);
  const key = `${baseUrl?.href ?? ''}\n${face.file}`;
  const cached = bytesCache.get(key);
  if (cached) return cached;
  const promise = assetUrl(face.file, baseUrl).then(async (url) => {
    const onDisk = readFileAsset(url);
    if (onDisk) return validateByteLength(face, onDisk);
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to fetch bundled font ${face.file}: HTTP ${response.status}`);
    }
    return validateByteLength(face, await response.arrayBuffer());
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

/** Structural provider contract keeps this package independent of the engine. */
export interface BundledFontSource {
  resolve(family: string, bold: boolean, italic: boolean): (() => Promise<ArrayBuffer>) | undefined;
  resolveScriptFallback(
    script: BundledFontScript,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
  resolveLastResort(family: string, bold: boolean, italic: boolean): () => Promise<ArrayBuffer>;
}

/** Creates a lazy provider, optionally serving assets from a custom base URL. */
export function createFontProvider(options?: FontAssetOptions): BundledFontSource {
  const resolvedOptions =
    options?.baseUrl === undefined ? undefined : { baseUrl: resolvedAssetBase(options.baseUrl) };
  const load = (face: BundledFontFace) => () => loadBundledFontBytes(face, resolvedOptions);
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
