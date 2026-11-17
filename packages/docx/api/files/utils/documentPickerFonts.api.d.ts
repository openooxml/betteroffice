// @betteroffice/docx declaration utils/documentPickerFonts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Discover the fonts a document actually references and that the browser can
 * render, so they can be offered in the font picker under a "Document fonts"
 * group. A font qualifies when it is embedded in the file (already loaded via
 * {@link loadEmbeddedFonts}) or when the host system can render it
 * ({@link canRenderFont}). Fonts that would only fall back to a substitute are
 * left out, so the selector never lists a face it cannot actually show.
 */
import type { FontOption } from './fontOptions';
import type { Document } from '../types';
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
export declare function selectRenderableFonts(names: readonly string[], options?: RenderableFontOptions): FontOption[];
/**
 * Walk a parsed document for the fonts it references and return those the
 * browser can render (embedded or system-resolved) as picker options.
 *
 * @public
 */
export declare function getRenderableDocumentFonts(doc: Document, options?: RenderableFontOptions): FontOption[];
/**
 * Drop fonts whose names already appear in `existingNames` (case-insensitive),
 * also deduping the input. Used by both adapters' pickers to render the
 * "Document fonts" group without repeating a font the built-in list covers.
 *
 * @public
 */
export declare function excludeFontsByName(fonts: readonly FontOption[] | undefined, existingNames: Iterable<string>): FontOption[];
