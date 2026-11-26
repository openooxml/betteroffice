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
    /** Load every outline of the face once the editor is idle, so any character in it draws without waiting. */
    fill?: boolean;
}
/** What the worker uses of a face `@betteroffice/fonts/progressive` opens. */
interface OutlineFace {
    readonly revision: number;
    bytes(): ArrayBuffer;
    ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
    complete(): boolean;
    /** Absent from a `@betteroffice/fonts` release older than filling; such a face only loads as frames need it. */
    chunks?(): readonly number[];
    loadedChunks?(): ReadonlySet<number>;
    ensureChunks?(chunks: Iterable<number>): Promise<number>;
    chunksOfGlyphs?(glyphIds: Iterable<number>): readonly number[];
}
/**
 * Opened faces by identity, kept by the worker across font environments: a face a fill loaded, or one frames
 * loaded, is not fetched and decoded again when the host sends the fonts anew.
 */
export type SharedOutlineFaces = Map<string, Promise<OutlineFace>>;
/** How long a frame waits for missing outlines before it paints without them and repaints later. */
export declare const PROGRESSIVE_OUTLINE_WAIT_MS = 10000;
/** A fill starts this long after the fonts land, so opening and the first frames go first. */
export declare const FILL_DELAY_MS = 1000;
/** A fill loads its next chunks only once no frame has drawn for this long. */
export declare const FILL_QUIET_MS = 300;
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
    private readonly loadModule;
    private readonly owners;
    private readonly ownerOf;
    private readonly pending;
    private readonly lateListeners;
    private readonly refusalListeners;
    private readonly fillListeners;
    private readonly trapListeners;
    private module?;
    private warned;
    private late;
    private disposed;
    /** When a frame last asked for outlines; a fill waits for a quiet stretch after it. */
    private lastFrame;
    private fillTimer;
    private filling;
    /** Outline installs that arrived after the frame that needed them had painted. */
    get lateInstalls(): number;
    constructor(specs: readonly ProgressiveFontSpec[], ids: readonly (number | null)[], substituteOf: readonly (number | undefined)[], engine: OutlineEngine, decode: (encoded: ArrayBuffer, maxOutput: number) => Promise<ArrayBuffer>, log?: (message: string) => void, loadModule?: () => Promise<ProgressiveModule>, shared?: SharedOutlineFaces);
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
     * or after `waitMs` of waiting for outlines once the module loads.
     */
    ensure(glyphs: Map<number, Set<number>>, waitMs?: number): Promise<Set<number>>;
    private ensureOutlines;
    /** Settles once every outline requested so far is installed or has failed. */
    whenInstalled(): Promise<void>;
    /** Calls `listener` with the owner font ids a fill installed; glyph images cached from their earlier bytes go stale. */
    onFilled(listener: (owners: ReadonlySet<number>) => void): () => void;
    /**
     * Loads every outline of the faces marked `fill`, a few chunks at a time while no frame draws, and installs each
     * face once it is whole, so later text in it never waits for a chunk. Starts `FILL_DELAY_MS` after the first call;
     * later calls while one is scheduled or running do nothing. A face that cannot fill keeps loading as frames need it.
     */
    scheduleFill(): void;
    /** Calls `listener` with an engine trap a fill hit while installing outlines; the engine cannot go on. */
    onTrapped(listener: (error: Error) => void): () => void;
    /** Settles once a scheduled fill has run, or at once without one. */
    whenFilled(): Promise<void>;
    private fill;
    /** Resolves once no frame has asked for outlines for `FILL_QUIET_MS` and no frame's outlines are loading. */
    private quiet;
    /** The progressive font module, loaded once; after a failed load the next caller tries again. */
    private progressiveModule;
    private track;
    /**
     * Loads the chunks of `ids`, retrying a chunk that fails. Chunks that keep
     * failing leave their glyphs blank; the next frame drawing them asks again.
     */
    private load;
    /** Whether the engine's bytes already hold every outline these glyph ids draw from; unknown counts as no. */
    private drawsInstalled;
    /** Installs the face's current outlines; on refusal stops progressive loading for it. True when outlines changed. */
    private install;
    /** Stops progressive loading for `owner` and asks the host to reload it whole. */
    private refuse;
    /** Calls `listener` with a font id that cannot load progressively here (the engine refused its outlines or the module is unavailable); the host reloads it whole. */
    onRefused(listener: (owner: number) => void): () => void;
}
export {};
