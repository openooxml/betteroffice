// @betteroffice/docx declaration utils/embeddedFonts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Embedded font loading.
 *
 * Turns the obfuscated `word/fonts/*.odttf` binaries a DOCX carries into live
 * `@font-face` registrations so documents render in their authored fonts
 * instead of falling back to a metric-compatible substitute.
 *
 * The pure half ({@link getEmbeddedFontFaces}) resolves + de-obfuscates the
 * faces; {@link loadEmbeddedFonts} registers them with the browser (no-op
 * outside a DOM). Round-trip preservation is unaffected: the serializer copies
 * `word/fonts/*` and `fontTable.xml` from the original package untouched.
 */
import { type FontLoadScope } from './fontLoader';
import type { FontTable } from '../types/styles';
import type { Document } from '../types/document';
/**
 * Names of the fonts a table embeds at least one face for. Used by the picker
 * to surface embedded fonts even when the canvas probe is unreliable for a
 * subsetted face.
 *
 * @public
 */
export declare function getEmbeddedFontFamilies(fontTable: FontTable | undefined): Set<string>;
/** A single de-obfuscated embedded font face, ready for `loadFontFromBuffer`. */
export interface EmbeddedFontFace {
    /** Word font name to register the face under. */
    family: string;
    /** CSS `font-weight` the face maps to (`embedBold*` → `'bold'`). */
    weight: 'normal' | 'bold';
    /** CSS `font-style` the face maps to (`embed*Italic` → `'italic'`). */
    style: 'normal' | 'italic';
    /** De-obfuscated OpenType/TrueType bytes. */
    data: ArrayBuffer;
    /** Whether the source face was subsetted (`w:subsetted`). */
    subsetted: boolean;
}
/** @internal */
export declare function retainEmbeddedFontFaces(document: Document, faces: EmbeddedFontFace[]): void;
/**
 * Resolve and de-obfuscate every embedded font face declared in a font table.
 * Pure: does not touch the DOM. Faces whose relationship or binary is missing,
 * or whose key is unusable, are skipped.
 *
 * @param fontTable - Parsed `fontTable.xml` (`pkg.fontTable`).
 * @param rawFonts - Unzipped font binaries keyed by package path.
 * @param fontTableRelsXml - Raw `word/_rels/fontTable.xml.rels` XML.
 * @public
 */
export declare function getEmbeddedFontFaces(fontTable: FontTable | undefined, rawFonts: ReadonlyMap<string, ArrayBuffer>, fontTableRelsXml: string | null | undefined): EmbeddedFontFace[];
/**
 * Resolve + de-obfuscate the embedded faces of an already-parsed Document by
 * re-reading its `originalBuffer` (parseDocx registers the faces with the
 * browser during load but does not retain the raw bytes on the model).
 * Returns `[]` when the document declares no embeds, carries no original
 * buffer, or the parts cannot be resolved. The zip pass goes through the
 * hardened `unzipDocx` container facade.
 *
 * Used by measurement, which needs the byte
 * form of every embedded face for the wasm FontStore.
 *
 * @public
 */
export declare function extractEmbeddedFontFaces(document: Document): Promise<EmbeddedFontFace[]>;
/**
 * Register every embedded font face with the browser via `@font-face`. No-op
 * outside a DOM (headless/SSR). Resolves to the set of font family names that
 * were registered (deduped), so callers can surface them in the font picker.
 *
 * @public
 */
export declare function loadEmbeddedFonts(fontTable: FontTable | undefined, rawFonts: ReadonlyMap<string, ArrayBuffer>, fontTableRelsXml: string | null | undefined): Promise<Set<string>>;
/**
 * {@link loadEmbeddedFonts}, resolving to the CSS family each embedded family
 * was registered under: its own name, or an alias when another live document
 * registered different faces under that name (see {@link registerDocumentFaces}).
 * With a scope, the faces are held until its next document or its disposal,
 * and a call that a later document load into the scope overtook resolves to
 * an empty map.
 *
 * @public
 */
export declare function loadEmbeddedFontFamilies(fontTable: FontTable | undefined, rawFonts: ReadonlyMap<string, ArrayBuffer>, fontTableRelsXml: string | null | undefined, scope?: FontLoadScope): Promise<Map<string, string>>;
