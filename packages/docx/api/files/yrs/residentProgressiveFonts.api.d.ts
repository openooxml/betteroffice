// @betteroffice/docx declaration yrs/residentProgressiveFonts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** A sliced face's pack files, as `sliceSourceOf` from `@betteroffice/fonts/progressive` describes them. */
export interface ProgressiveFontSource {
    ref: {
        format: number;
        index: string;
        indexSha256: string;
    };
    /** Absolute URL of the index, skeleton, core and every chunk by manifest path. */
    urls: Record<string, string>;
    timeoutMs?: number;
}
/** A command font slot registered from a sliced face's skeleton and core. */
export interface ProgressiveFontSpec {
    slot: number;
    source: ProgressiveFontSource;
}
/** What the worker uses of a face `@betteroffice/fonts/progressive` opens. */
interface OutlineFace {
    readonly revision: number;
    bytes(): ArrayBuffer;
    ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
    complete(): boolean;
}
/** How long a frame waits for missing outlines before it paints without them and repaints later. */
export declare const PROGRESSIVE_OUTLINE_WAIT_MS = 10000;
interface OutlineEngine {
    updateFontOutlines(fontId: number, bytes: Uint8Array): void;
}
/** Glyph ids per font id drawn by a display page: its own runs and those of its regions and note areas. */
export declare function glyphsByFont(page: unknown, into?: Map<number, Set<number>>): Map<number, Set<number>>;
/** What the worker uses of `@betteroffice/fonts/progressive`. */
interface ProgressiveModule {
    openProgressiveFace(ref: ProgressiveFontSource['ref'], transport: {
        read(file: string): Promise<ArrayBuffer>;
        decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
    }): Promise<OutlineFace>;
    readFontAsset(file: string, url: URL, timeoutMs?: number): Promise<ArrayBuffer>;
}
/**
 * Sliced CJK faces registered from their skeleton and core. Before a page is
 * rasterized, the outlines of the glyphs it draws are fetched (each chunk
 * once, shared across frames) and installed in place; layout never changes.
 */
export declare class ResidentProgressiveFonts {
    private readonly engine;
    private readonly log;
    private readonly owners;
    private readonly ownerOf;
    private readonly pending;
    private readonly lateListeners;
    private readonly refusalListeners;
    private warned;
    private late;
    private disposed;
    /** Outline installs that arrived after the frame that needed them had painted. */
    get lateInstalls(): number;
    constructor(specs: readonly ProgressiveFontSpec[], ids: readonly (number | null)[], substituteOf: readonly (number | undefined)[], engine: OutlineEngine, decode: (encoded: ArrayBuffer, maxOutput: number) => Promise<ArrayBuffer>, log?: (message: string) => void, loadModule?: () => Promise<ProgressiveModule>);
    /** Stops installing outlines, for fonts that were cleared or replaced; pending loads finish unused. */
    dispose(): void;
    get size(): number;
    /** Every font id drawing with the bytes of `owners`: the owners and their substitute views. */
    fontIdsOf(owners: ReadonlySet<number>): Set<number>;
    /** Calls `listener` when outlines a frame stopped waiting for are installed. */
    onLateInstall(listener: () => void): () => void;
    /**
     * Installs the outlines of `glyphs` (per font id) before a raster. Resolves
     * with the owner font ids whose outlines changed once they are installed,
     * or after `waitMs` with those installed so far; the rest install later.
     */
    ensure(glyphs: Map<number, Set<number>>, waitMs?: number): Promise<Set<number>>;
    /** Settles once every outline requested so far is installed or has failed. */
    whenInstalled(): Promise<void>;
    private track;
    /**
     * Loads the chunks of `ids`, retrying a chunk that fails. Chunks that keep
     * failing leave their glyphs blank; the next frame drawing them asks again.
     */
    private load;
    /** Installs the face's current outlines; on refusal stops progressive loading for it. True when outlines changed. */
    private install;
    /** Stops progressive loading for `owner` and asks the host to reload it whole. */
    private refuse;
    /** Calls `listener` with a font id that cannot load progressively here (the engine refused its outlines or the module is unavailable); the host reloads it whole. */
    onRefused(listener: (owner: number) => void): () => void;
}
export {};
