/**
 * Discover the fonts a document actually references and that the browser can
 * render, so they can be offered in the font picker under a "Document fonts"
 * group. A font qualifies when it is embedded in the file (already loaded via
 * {@link loadEmbeddedFonts}) or when the host system can render it
 * ({@link canRenderFont}). Fonts that would only fall back to a substitute are
 * left out, so the selector never lists a face it cannot actually show.
 */

import { extractFonts } from './fontExtractor';
import { resolveFontFamily } from './fontResolver';
import { canRenderFont } from './fontLoader';
import type { FontOption } from './fontOptions';
import type { Document } from '../types';

const GENERIC_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
]);

export interface RenderableFontOptions {
  /** Font families already loaded from the document (embedded faces). */
  embeddedFamilies?: ReadonlySet<string>;
  /** Names already present in the picker (built-in / configured) to skip. */
  exclude?: Iterable<string>;
  /** Override the system-font probe. Defaults to {@link canRenderFont}. */
  canRender?: (name: string) => boolean;
}

/**
 * Filter a list of referenced font names down to renderable {@link FontOption}s.
 * Pure and DOM-free when `canRender` is injected.
 *
 * @public
 */
export function selectRenderableFonts(
  names: readonly string[],
  options: RenderableFontOptions = {}
): FontOption[] {
  const canRender = options.canRender ?? canRenderFont;
  const exclude = new Set<string>();
  for (const name of options.exclude ?? []) exclude.add(name.trim().toLowerCase());
  const embedded = new Set<string>();
  for (const name of options.embeddedFamilies ?? []) embedded.add(name.trim().toLowerCase());

  const out: FontOption[] = [];
  const seen = new Set<string>();
  for (const raw of names) {
    const name = raw.trim();
    const lower = name.toLowerCase();
    if (!name || GENERIC_FAMILIES.has(lower)) continue;
    if (exclude.has(lower) || seen.has(lower)) continue;
    if (!embedded.has(lower) && !canRender(name)) continue;

    seen.add(lower);
    out.push({
      name,
      fontFamily: resolveFontFamily(name).cssFallback,
      category: 'other',
    });
  }
  return out;
}

/**
 * Walk a parsed document for the fonts it references and return those the
 * browser can render (embedded or system-resolved) as picker options.
 *
 * @public
 */
export function getRenderableDocumentFonts(
  doc: Document,
  options: RenderableFontOptions = {}
): FontOption[] {
  return selectRenderableFonts(extractFonts(doc), options);
}

/**
 * Drop fonts whose names already appear in `existingNames` (case-insensitive),
 * also deduping the input. Used by both adapters' pickers to render the
 * "Document fonts" group without repeating a font the built-in list covers.
 *
 * @public
 */
export function excludeFontsByName(
  fonts: readonly FontOption[] | undefined,
  existingNames: Iterable<string>
): FontOption[] {
  if (!fonts || fonts.length === 0) return [];
  const existing = new Set<string>();
  for (const name of existingNames) existing.add(name.trim().toLowerCase());
  const seen = new Set<string>();
  const out: FontOption[] = [];
  for (const f of fonts) {
    const key = f.name.trim().toLowerCase();
    if (existing.has(key) || seen.has(key)) continue;
    seen.add(key);
    out.push(f);
  }
  return out;
}
