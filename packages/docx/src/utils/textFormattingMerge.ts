import type { ColorValue, TextFormatting } from '../types/document';
import { mergeFontFamily } from './fontFamilyMerge';

/**
 * Merge two `TextFormatting` objects (source overrides target).
 *
 * Used everywhere OOXML rPr inheritance is resolved — basedOn chains in
 * `styleParser`, paragraph-style merge in `styleResolver`, and pPr/rPr
 * onto style rPr in `toProseDoc`. Keeping one implementation prevents the
 * three diverging again (#391, #394).
 *
 * Merge rules per ECMA-376 §17.3.2:
 * - `fontFamily` — per-slot merge (§17.3.2.27): each ascii/hAnsi/eastAsia/cs
 *   slot and its theme pair travel together; source only replaces slots it
 *   actually sets. See `mergeFontFamily`.
 * - `color` — `w:val="auto"` (§17.3.2.6) means "inherit" and does not
 *   override an explicit color from the chain unless source also names an
 *   explicit color or theme reference.
 * - other object-shaped fields (`underline`, `borders`, etc.) — shallow
 *   merge so a child rPr that only sets one property does not wipe out
 *   sibling properties from the parent.
 * - primitives — replace.
 */
export function mergeTextFormatting(
  target: TextFormatting | undefined,
  source: TextFormatting | undefined
): TextFormatting | undefined {
  if (!source && !target) return undefined;
  if (!source) return target;
  if (!target) return { ...source };

  const result: Record<string, unknown> = { ...target };

  for (const key of Object.keys(source) as (keyof TextFormatting)[]) {
    const value = source[key];
    if (value === undefined) continue;

    if (key === 'fontFamily' && typeof value === 'object' && value !== null) {
      result.fontFamily = mergeFontFamily(
        target.fontFamily,
        value as NonNullable<TextFormatting['fontFamily']>
      );
      continue;
    }

    if (key === 'color' && typeof value === 'object' && value !== null) {
      const c = value as ColorValue;
      const hasExplicit = !!(c.rgb || c.themeColor || c.themeTint || c.themeShade);
      if (!c.auto || hasExplicit) result.color = c;
      continue;
    }

    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      result[key] = {
        ...((target[key] as Record<string, unknown> | undefined) ?? {}),
        ...(value as Record<string, unknown>),
      };
      continue;
    }

    result[key] = value;
  }

  return result as TextFormatting;
}
