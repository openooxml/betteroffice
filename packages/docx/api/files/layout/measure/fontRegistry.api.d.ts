// @betteroffice/docx declaration layout/measure/fontRegistry.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Where a sliced face's files are, as `@betteroffice/fonts` loaders carry it; opaque to the registry. */
export type SliceLocation = object;
/** What the registry uses of a face `@betteroffice/fonts/progressive` opens. */
export interface ProgressiveFontFace {
    bytes(): ArrayBuffer;
}
import { type FontCoverage } from './fontCoverage';
/** A lazy face loader with measurement and fallback options. */
export type BundledFontLoader = (() => Promise<ArrayBuffer>) & {
    /** Word measurement family; defaults to the requested family. */
    measureAs?: string;
    /** Only stands in for a primary face that failed to load. */
    onlyIfPrimaryFails?: boolean;
    /** Locates a sliced face whose outlines load as they are needed. */
    slices?: SliceLocation;
    /** @internal Precomputed base Latin coverage bits of the face. */
    coverage?: string;
    /** @internal The face's decoded bytes, shared with the font cache: read them, never change or transfer them. */
    readOnlyBytes?: () => Promise<ArrayBuffer>;
};
/** Byte sink of the wasm text engine (`crates/docx-text` FontStore). */
export interface TextEngineFontSink {
    /**
     * Register raw sfnt bytes with the engine; returns its font id.
     * Throws on unparseable bytes.
     */
    registerFont(bytes: Uint8Array, face?: {
        family: string;
        bold: boolean;
        italic: boolean;
    }): number;
    /**
     * Register a measurement view of `id` carrying the vertical metrics and
     * advance pitch Word measures `family` with, and return its id — the
     * engine's answer to a face this host had to substitute. Returns `id` for a
     * family whose metrics the engine does not know. Optional so partial sinks
     * (tests, older hosts) keep working; without it a substitute measures with
     * its own metrics.
     */
    registerSubstituteFont?(id: number, family: string): number;
}
/**
 * Script bucket a fallback face provides glyph coverage for. Han text is
 * region-resolved by the caller (SC/TC/JP/KR — see `rustMeasureSource.ts`);
 * kana, Hangul, Arabic and Hebrew detect directly.
 *
 * @public
 */
export type FontScript = 'cjk-sc' | 'cjk-tc' | 'cjk-jp' | 'cjk-kr' | 'arabic' | 'hebrew';
/**
 * Resolver for the bundled metric-compatible set (Carlito↔Calibri,
 * Liberation↔Arial/Times/Courier, …). Returns a lazy byte loader so the
 * (lazily fetched, same-origin) binary is only downloaded when a document
 * actually needs the face, or `undefined` when the family has no bundled
 * substitute.
 *
 * @public
 */
export interface BundledFontProvider {
    /** @internal */
    resolveBaseLatin?(bold: boolean, italic: boolean): BundledFontLoader | undefined;
    /** Resolve a Word family to bundled metric-compatible face byte loaders, or undefined. */
    resolve(family: string, bold: boolean, italic: boolean): BundledFontLoader | undefined;
    /** Like {@link BundledFontProvider.resolve}, but also accepts bundled family names such as `"Gelasio"`, and leaves missing styles for browser synthesis. */
    resolveFamily?(family: string, bold: boolean, italic: boolean): (() => Promise<ArrayBuffer>) | undefined;
    /**
     * Optional per-script coverage fallback (Noto CJK/RTL faces). Same loader
     * contract as {@link BundledFontProvider.resolve}; providers without
     * script faces simply omit the method. These faces are coverage fallbacks
     * first, metric approximations second — CJK metric compatibility is much
     * weaker than the Carlito/Calibri class of substitutes. `family` is the
     * family the text requests, which can select a closer face of the script.
     */
    resolveScriptFallback?(script: FontScript, bold: boolean, italic: boolean, family?: string): (() => Promise<ArrayBuffer>) | undefined;
    /** Coverage for characters absent from a bundled substitute. */
    resolveGlyphFallback?(family: string, bold: boolean, italic: boolean, office?: 'word' | 'powerpoint'): (() => Promise<ArrayBuffer>) | undefined;
    /**
     * The always-available last-resort base face (broad-coverage Latin —
     * Liberation Sans/Serif per serif-ness). Unlike {@link
     * BundledFontProvider.resolve}, a conforming provider (e.g.
     * `@betteroffice/fonts`'s `resolveLastResortFace`) returns a loader for
     * EVERY family, so the chain this registry builds is guaranteed non-empty
     * and a run never routes to the browser measurer for want of font bytes. The
     * face's metrics are the base font's, not the requested family's — the
     * accepted divergence for a truly-unknown font, in exchange for staying on
     * the native measurement path.
     *
     * Optional so mock/partial providers can omit it; when absent (or returning
     * undefined) an unmapped family still yields an empty chain and the caller
     * browser-falls-back that run — the pre-policy behavior. Same lazy loader
     * contract as {@link BundledFontProvider.resolve}. This registry passes
     * `office: 'word'` so the pick follows Word's substitution.
     */
    resolveLastResort?(family: string, bold: boolean, italic: boolean, office?: 'word' | 'powerpoint'): BundledFontLoader | undefined;
}
/**
 * One embedded face extracted from the open document. Structurally identical
 * to `EmbeddedFontFace` (`utils/embeddedFonts.ts`), so the output of
 * `getEmbeddedFontFaces` feeds straight in; duplicated here so this module
 * stays decoupled from the loader.
 *
 * @public
 */
export interface EmbeddedFaceInput {
    /** Word font name the face is registered under (attacker-controlled). */
    family: string;
    /** `'bold'` for the embedBold/embedBoldItalic slots, else `'normal'`. */
    weight: 'normal' | 'bold';
    /** `'italic'` for the embedItalic/embedBoldItalic slots, else `'normal'`. */
    style: 'normal' | 'italic';
    /** De-obfuscated OpenType/TrueType bytes. */
    data: ArrayBuffer;
    /** Whether the source face was subsetted (`w:subsetted`). */
    subsetted: boolean;
}
/** A provider or async factory; `undefined` means no bundled bytes. @public */
export type BundledFontProviderSource = BundledFontProvider | (() => Promise<BundledFontProvider | undefined>);
export declare class TextMeasureFontRegistry {
    private coverage;
    /** Ids whose coverage is a known lower bound, with the loader that yields their exact coverage. */
    private inexact;
    private baseLatinResults;
    private readonly baseLatinSource;
    private baseLatinLoaders;
    private readonly sink;
    private readonly bundledSource;
    private bundledPromise;
    /** Per-registry so misconfigured editors warn independently. */
    private warnedFallback;
    /** Normalized family → embedded faces of the current document. */
    private facesByFamily;
    private appFacesByFamily;
    /**
     * Per-face registration memo, keyed by face object identity so re-feeding
     * the same faces (or sharing one face across several chains) never
     * re-registers its bytes. Failed registrations are evicted. Weak so faces of
     * a closed document can be collected.
     */
    private faceIds;
    private bundledIds;
    /**
     * Last-resort base-face registration memo, keyed by chain key. Kept separate
     * from `bundledIds` so a family's metric-compat face and its (possibly
     * different) last-resort base face don't clobber each other's registration.
     * Document-independent like `bundledIds` — survives `setEmbeddedFaces`, reset
     * by `clear()`. Failed registrations are evicted.
     */
    private lastResortIds;
    /**
     * Byte-identity dedupe across every registration path. Rejections are
     * evicted. Weak so buffers of a cleared registry can be collected.
     */
    private bufferIds;
    private loaderIds;
    private scriptIds;
    /** Settled per-script, per-requested-family results for the synchronous view. */
    private scriptResults;
    /** Chain memo — concurrent `getFontIdChain` calls share one resolution. */
    private chains;
    /** Settled chains for the synchronous view. */
    private chainResults;
    /** Bumped on invalidation so stale in-flight resolutions can't repopulate caches. */
    private generation;
    /** Script-fallback faces registered from their skeleton and core, by font id. */
    private progressiveIds;
    private readonly progressive?;
    /** `progressive` opens sliced faces, which register from their skeleton and core. */
    constructor(sink: TextEngineFontSink, opts?: {
        bundled?: BundledFontProviderSource;
        baseLatin?: BundledFontProviderSource;
        progressive?: (slices: SliceLocation) => Promise<ProgressiveFontFace>;
    });
    /** The sliced faces registered from their skeleton and core, by font id. */
    progressiveFaces(): ReadonlyMap<number, ProgressiveFontFace>;
    /** Shares in-flight resolution but evicts misses for retry. */
    private bundled;
    /**
     * Feed the embedded faces extracted from the open document, replacing any
     * previous set and invalidating every cached chain. Font ids already issued
     * by the sink stay valid (the engine keeps the bytes); faces passed again
     * by object identity keep their registration. Script-fallback
     * registrations are document-independent and survive.
     */
    setEmbeddedFaces(faces: EmbeddedFaceInput[], appFaces?: EmbeddedFaceInput[]): void;
    /**
     * Lazily resolves and registers a font chain. Concurrent calls share work;
     * failed registrations remain retryable.
     */
    getFontIdChain(family: string, bold: boolean, italic: boolean): Promise<number[]>;
    /** Synchronous settled view; retryable results are evicted after one read. */
    getCachedFontIdChain(family: string, bold: boolean, italic: boolean, consumeRetryable?: boolean): readonly number[] | undefined;
    coverageOf(id: number): FontCoverage | undefined;
    /** Characters the base Latin tier settled for `key`, without those whose face failed to load. */
    settledBaseLatinCharacters(key: string): Set<string>;
    getCachedBaseLatinIds(family: string, bold: boolean, italic: boolean, characters?: readonly string[], earlier?: readonly number[], key?: string): readonly number[] | undefined;
    prepareBaseLatin(family: string, bold: boolean, italic: boolean, characters: readonly string[], earlier: readonly number[], key?: string): Promise<boolean>;
    /** Resolves unique script fallbacks in order; stable misses cache and failures retry. */
    getScriptFallbackIds(scripts: FontScript[], family?: string): Promise<number[]>;
    /** Synchronous settled view; retryable results are evicted after one pass. */
    getCachedScriptFallbackIds(scripts: FontScript[], consumeRetryable?: boolean, family?: string): readonly number[] | undefined;
    /**
     * Forget every cached chain, face registration, bundled registration and
     * script-fallback registration. Call when the engine/FontStore is
     * recreated — previously issued font ids are invalid then. The embedded
     * face set is retained; pass a new set via `setEmbeddedFaces` when the
     * document changes.
     */
    clear(): void;
    private warnFailure;
    private resolveChain;
    /**
     * An empty chain forces browser measurement. Distinguish absent, incomplete,
     * and broken providers so the one warning gives applicable remediation.
     */
    private warnFallback;
    /**
     * Exact (weight, style) match first; otherwise the family's regular face —
     * the engine synthesizes bold/italic from regular outlines when asked for a
     * style the document did not embed.
     */
    private pickEmbeddedFace;
    /**
     * The id that measures `base`'s bytes the way Word measures `family` — what
     * a chain head must be for a substituted face, since line metrics come from
     * the head. Falls back to `base` when the sink or the engine has no metrics
     * for the family, which is how every face measured before this existed.
     */
    private measuredAs;
    /** Register raw bytes exactly once per buffer identity (see `bufferIds`). */
    private registerBuffer;
    /** Replaces a known lower bound by the face's exact coverage. */
    private exactCoverage;
    /** Registers a sliced face's skeleton and core when progressive loading is on, else the whole face. */
    private registerFaceLoader;
    private registerLoader;
    private registerFace;
    private registerBundled;
    /**
     * Register the last-resort base face. Mirrors {@link registerBundled} but
     * uses its own memo so it never collides with metric-compatible registration.
     */
    private registerLastResort;
    private resolveScript;
}
