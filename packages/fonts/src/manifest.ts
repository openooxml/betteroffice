/** Deterministic bundled-font resolution with lazy loading and optional CJK assets. */

/** Script bucket a bundled face provides glyph coverage for. */
export type BundledFontScript =
  | 'cjk-sc'
  | 'cjk-tc'
  | 'cjk-jp'
  | 'cjk-kr'
  | 'arabic'
  | 'hebrew';

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
  byteLengths: readonly [number, number, number, number],
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
  ...familyFaces(
    'Carlito',
    'Calibri',
    'Carlito',
    [628032, 682468, 615236, 808508],
  ),
  ...familyFaces('Caladea', 'Cambria', 'Caladea', [81600, 84492, 83780, 83356]),
  ...familyFaces(
    'Liberation Sans',
    'Arial',
    'LiberationSans',
    [410712, 414456, 415816, 408996],
  ),
  ...familyFaces(
    'Liberation Serif',
    'Times New Roman',
    'LiberationSerif',
    [393576, 370096, 375632, 376772],
  ),
  ...familyFaces(
    'Liberation Mono',
    'Courier New',
    'LiberationMono',
    [319508, 307996, 281536, 284068],
  ),

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
    metricCompatByWordFamily.set(
      face.metricCompatWith.toLowerCase(),
      face.family,
    );
  }
}

/**
 * Resolve a Word font name (case-insensitive) to the bundled substitute
 * family, e.g. `"calibri"` -> `"Carlito"`, `"SimSun"` -> `"Noto Serif SC"`.
 * Returns `undefined` when no bundled font covers the name.
 */
export function resolveMetricCompatFamily(
  wordFamily: string,
): string | undefined {
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
  italic: boolean,
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
  italic: boolean,
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
  italic: boolean,
): BundledFontFace {
  // Liberation Sans/Serif always ship the full four-face set, so the
  // metric-compat resolution is guaranteed to return a face here.
  const base = looksSerif(family) ? 'Times New Roman' : 'Arial';
  return resolveMetricCompatFace(base, bold, italic)!;
}
