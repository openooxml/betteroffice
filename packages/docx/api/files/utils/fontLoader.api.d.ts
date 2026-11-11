// @betteroffice/docx declaration utils/fontLoader.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Browser Font Loader
 *
 * Dynamically loads managed and app-configured fonts with:
 * - Loading state tracking
 * - Duplicate prevention
 * - Callback notifications
 * - Font availability detection
 */
/**
 * The loads and notifications of one editor instance; see
 * {@link createFontLoadScope}.
 *
 * @public
 */
export interface FontLoadScope {
    /** Hears the loads this scope asked for and module-level loads, not other scopes' loads. */
    onFontsLoaded(callback: (fonts: string[]) => void): () => void;
    /** Hears the failures of this scope's loads and of module-level loads. */
    onFontError(callback: (error: Error) => void): () => void;
    /** {@link loadDocumentFonts} on behalf of this scope. */
    loadDocumentFonts(document: unknown): Promise<void>;
    /** {@link loadFontsWithMapping} on behalf of this scope. */
    loadFontsWithMapping(families: string[]): Promise<void>;
    /** {@link loadFontDefinitions} on behalf of this scope. */
    loadFontDefinitions(defs: ReadonlyArray<FontDefinition> | undefined): Promise<void>;
    /** Drops the listeners and releases the embedded faces this scope holds; a later listener revives it. */
    dispose(): void;
    /** Whether it is disposed and not revived since. */
    readonly disposed: boolean;
}
/**
 * Load a font from the default font provider
 *
 * @param fontFamily - The font family name to load
 * @param options - Optional configuration
 * @returns Promise resolving to true if font loaded successfully, false otherwise
 */
export declare function loadFont(fontFamily: string, options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
}): Promise<boolean>;
/**
 * Load multiple fonts from the default font provider
 *
 * @param families - Array of font family names to load
 * @param options - Optional configuration
 * @returns Promise resolving when all fonts are loaded (or failed)
 */
export declare function loadFonts(families: string[], options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
}): Promise<void>;
/**
 * Check if a font is loaded
 *
 * @param fontFamily - The font family name to check
 * @returns true if the font is loaded, false otherwise
 */
export declare function isFontLoaded(fontFamily: string): boolean;
/**
 * Enable or disable the editor's automatic Google Fonts lookup.
 *
 * Defaults to enabled. Set to `false` in no-egress embedders (strict CSP,
 * offline) so `loadFont` / `loadFontWithMapping` never inject a
 * `fonts.googleapis.com` stylesheet `<link>`. This gates ONLY the implicit
 * Google Fonts lookup — font URLs you register yourself (the `fonts` prop /
 * `loadFontFromUrl`) are still fetched by the browser, and embedded blobs
 * (`loadFontFromBuffer`) and system fonts still resolve.
 *
 * The flag is page-global (module-level), not per-editor: with multiple
 * editors on one page the last caller wins. Call it before loading documents;
 * disabling does not cancel fetches already in flight. A document font that
 * is neither local nor registered then resolves `loadFont` to `false`
 * silently (no `onFontError`) and renders via its CSS fallback stack.
 *
 * @see isGoogleFontsEnabled
 * @see loadFontFromBuffer
 * @see loadFontFromUrl
 *
 * @public
 */
export declare function setGoogleFontsEnabled(enabled: boolean): void;
/**
 * Whether the automatic Google Fonts lookup is currently enabled.
 *
 * @see setGoogleFontsEnabled
 *
 * @public
 */
export declare function isGoogleFontsEnabled(): boolean;
/**
 * Check if any fonts are currently loading
 *
 * @returns true if any fonts are loading, false otherwise
 */
export declare function isLoading(): boolean;
/**
 * Get list of all loaded fonts
 *
 * @returns Array of loaded font family names
 */
export declare function getLoadedFonts(): string[];
/**
 * Register a callback to be notified when fonts are loaded
 *
 * @param callback - Function to call when fonts are loaded
 * @returns Cleanup function to remove the callback
 */
export declare function onFontsLoaded(callback: (fonts: string[]) => void): () => void;
/**
 * Register a callback to be notified when a font fails to load.
 *
 * Adapters subscribe and forward to their `onError` prop. Returns the unsub.
 *
 * @public
 */
export declare function onFontError(callback: (error: Error) => void): () => void;
/**
 * Check if a font is available on the system using canvas measurement
 *
 * Compares text width with the target font vs a known fallback font (and the
 * opposite fallback, for names that collide with the browser defaults). If
 * the widths differ, the font is available. Reuses one shared canvas; the
 * fallback widths are session constants and memoized.
 *
 * @param fontFamily - The font family name to check
 * @param fallbackFont - Fallback font to compare against
 * @returns true if font is available, false otherwise
 */
export declare function canRenderFont(fontFamily: string, fallbackFont?: string): boolean;
/**
 * Load a font from a raw buffer (e.g., embedded in DOCX)
 *
 * Call before loading the document when possible: registration marks the
 * family as ours, which keeps the managed coverage fallback
 * in play for subsetted faces during the document's font resolution pass.
 * Late registration still self-heals on the next pass.
 *
 * @param fontFamily - The font family name
 * @param buffer - Font file buffer (TTF, OTF, WOFF, WOFF2)
 * @param options - Font options
 * @returns Promise resolving when font is loaded
 */
export declare function loadFontFromBuffer(fontFamily: string, buffer: ArrayBuffer, options?: {
    weight?: number | string;
    style?: 'normal' | 'italic';
}): Promise<boolean>;
/** One face of a font file's bytes, as {@link registerDocumentFaces} takes it. */
export interface BufferFaceInput {
    family: string;
    data: ArrayBuffer;
    weight?: number | string;
    style?: 'normal' | 'italic';
}
/**
 * Claims `scope` for a document's faces now, so that of two documents loading
 * into one scope the later claim wins, whichever faces arrive first. The
 * returned registration resolves to null when a later claim or disposal made
 * it obsolete.
 */
export declare function claimDocumentFaces(scope?: FontLoadScope): (faces: readonly BufferFaceInput[] | PromiseLike<readonly BufferFaceInput[]>) => Promise<Map<string, string> | null>;
/**
 * Registers one document's font faces from their bytes. Each family takes its
 * own name, as {@link loadFontFromBuffer} registers it, unless a live
 * registration holds different bytes under that name; then all its faces take
 * an alias, so both documents keep their own glyphs. Faces with the same bytes
 * are registered once. With a scope, the faces are held until the scope loads
 * its next document or is disposed; without one, for the page's lifetime.
 * `faces` may be a promise: the scope's next document is the one whose call
 * came last, whichever faces arrive first. Resolves to the CSS family each
 * family was registered under, for the families with at least one face that
 * loaded, or to an empty map once a later document or disposal replaced it.
 *
 * @public
 */
export declare function registerDocumentFaces(faces: readonly BufferFaceInput[] | PromiseLike<readonly BufferFaceInput[]>, scope?: FontLoadScope): Promise<Map<string, string>>;
/**
 * Load a font face from a URL (woff2, woff, ttf, otf).
 *
 * Injects an `@font-face` rule pointing at the URL. Multiple weights of the
 * same family can be registered independently. Families registered here are
 * treated as potentially subsetted: the editor also loads their managed
 * equivalent as a glyph-coverage fallback.
 *
 * @param fontFamily - CSS font-family name to expose
 * @param src - URL to the font file
 * @param options - Optional weight
 * @returns Promise resolving to true if the face became available
 *
 * @public
 */
export declare function loadFontFromUrl(fontFamily: string, src: string, options?: {
    weight?: number | string;
}): Promise<boolean>;
/**
 * Declarative description of a single font face to register with the editor.
 *
 * Each entry injects one `@font-face` rule pointing at a URL. Multiple
 * entries can share `family` to register distinct weights as separate faces.
 *
 * For managed fonts, call `loadFont(family)` directly; the `fonts` prop is
 * for fonts the consumer hosts themselves. For raw bytes already in memory
 * (DOCX-embedded fonts, user uploads), call `loadFontFromBuffer(family, buf)`.
 *
 * @public
 */
export interface FontDefinition {
    /**
     * CSS `font-family` name to expose. Match the family name your documents
     * reference; the browser uses this to look up glyphs when text is rendered.
     */
    family: string;
    /**
     * URL to the font file (woff2, woff, ttf, or otf). The loader injects an
     * `@font-face` rule and lets the browser fetch on demand.
     */
    src: string;
    /**
     * CSS `font-weight` for this face. Defaults to `'normal'` (≈400). Pass a
     * number (`400`, `700`) or a CSS keyword (`'bold'`). Required when one
     * `family` registers multiple weights as separate entries.
     */
    weight?: number | string;
}
/**
 * Register a list of custom font faces. Used by the `fonts` prop on
 * `<DocxEditor>` (React + Vue). Idempotent — safe to call on every render.
 *
 * @public
 */
export declare function loadFontDefinitions(defs: ReadonlyArray<FontDefinition> | undefined): Promise<void>;
/**
 * Mapping from common Office/system fonts to Google Fonts equivalents
 *
 * Google Fonts doesn't have exact matches for many Microsoft fonts,
 * but these are close alternatives that work well for document rendering.
 */
export declare const FONT_MAPPING: Record<string, string>;
/**
 * Get the Google Fonts equivalent for a font name
 *
 * @param fontName - The original font name from the document
 * @returns The Google Fonts equivalent, or the original name if no mapping exists
 */
export declare function getGoogleFontEquivalent(fontName: string): string;
/**
 * Load a font, automatically mapping to Google Fonts equivalent if needed.
 * If the font needs mapping, also creates a CSS alias so the original font
 * name works in stylesheets.
 *
 * @param fontFamily - The font family name (may be an Office font)
 * @returns Promise resolving to true if font loaded
 */
export declare function loadFontWithMapping(fontFamily: string): Promise<boolean>;
/**
 * Load multiple fonts with automatic mapping to Google Fonts equivalents
 *
 * @param families - Array of font family names
 * @returns Promise resolving when all fonts are loaded
 */
export declare function loadFontsWithMapping(families: string[]): Promise<void>;
/**
 * Preload a list of common document fonts
 *
 * This preloads fonts commonly used in DOCX documents that have
 * Google Fonts equivalents.
 */
export declare function preloadCommonFonts(): Promise<void>;
/**
 * Extract all font families used in a document
 *
 * Uses loose typing to handle any document-like structure.
 *
 * @param document - The parsed document
 * @returns Set of unique font family names
 */
export declare function extractFontsFromDocument(document: unknown): Set<string>;
/**
 * Extract fonts from a document and load them from Google Fonts
 *
 * @param document - The parsed document
 * @returns Promise resolving when fonts are loaded
 */
export declare function loadDocumentFonts(document: unknown): Promise<void>;
/**
 * A scope for one editor instance's font loads. Loads made through it notify
 * its own `onFontsLoaded` / `onFontError` listeners and the module-level ones,
 * never another scope's; module-level loads still reach every scope, so a
 * single editor hears what it heard before. Dedupe and caches stay shared.
 *
 * @public
 */
export declare function createFontLoadScope(): FontLoadScope;
