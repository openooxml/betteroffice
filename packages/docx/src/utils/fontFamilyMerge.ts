import type { TextFormatting } from '../types/document';

type FontFamily = NonNullable<TextFormatting['fontFamily']>;

/**
 * Merge two `<w:rFonts>` references per ECMA-376 §17.3.2.27.
 *
 * Each of the four font slots (`ascii`, `hAnsi`, `eastAsia`, `cs`) has an
 * explicit name and a paired theme reference (`asciiTheme`, etc.). When a
 * source style overrides a slot, both members of that pair must travel
 * together — otherwise an inherited theme attr can leak through and silently
 * override an explicit name (the `Theme` attr wins at render time, so a
 * stale `asciiTheme="minorHAnsi"` from `docDefaults` overrides an explicit
 * `ascii="Arial"` from the parent style and resolves back to `Calibri`).
 *
 * This is the rule that fixes #387.
 */
export function mergeFontFamily(target: FontFamily | undefined, source: FontFamily): FontFamily {
  const result: Record<string, unknown> = target ? { ...target } : {};
  const src = source as Record<string, unknown>;
  const pairs: Array<[explicit: string, theme: string]> = [
    ['ascii', 'asciiTheme'],
    ['hAnsi', 'hAnsiTheme'],
    ['eastAsia', 'eastAsiaTheme'],
    ['cs', 'csTheme'],
  ];
  const pairKeys = new Set<string>();
  // For each explicit/theme pair: if source touches either side, replace
  // both sides from source — never blend with target's pair.
  for (const [explicit, theme] of pairs) {
    pairKeys.add(explicit);
    pairKeys.add(theme);
    if (src[explicit] !== undefined || src[theme] !== undefined) {
      delete result[explicit];
      delete result[theme];
      if (src[explicit] !== undefined) result[explicit] = src[explicit];
      if (src[theme] !== undefined) result[theme] = src[theme];
    }
  }
  // Pass through any non-paired fontFamily keys (e.g. `hint`).
  for (const key of Object.keys(src)) {
    if (!pairKeys.has(key) && src[key] !== undefined) {
      result[key] = src[key];
    }
  }
  return result as FontFamily;
}
