// @betteroffice/docx declaration layout/render/glyphCache.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** outline path commands in font units, y-up (font convention). */
export type OutlineCommand = {
    t: 'M';
    x: number;
    y: number;
} | {
    t: 'L';
    x: number;
    y: number;
} | {
    t: 'Q';
    cx: number;
    cy: number;
    x: number;
    y: number;
} | {
    t: 'C';
    c1x: number;
    c1y: number;
    c2x: number;
    c2y: number;
    x: number;
    y: number;
} | {
    t: 'Z';
};
/** a cached glyph: its outline path (font units) plus the font's units-per-em. */
export interface GlyphOutline {
    /**
     * the glyph contours as a Path2D in font units, or `null` when the glyph has
     * no visible contours (space/whitespace) — the backend skips null paths.
     */
    path: Path2D | null;
    /** the font's units-per-em; the backend scales the path by `size`/upem. */
    upem: number;
}
/** fetch a glyph's outline as JSON — `outlineGlyphJson` in production, a fake in tests. */
export type GlyphOutlineProvider = (fontId: number, glyphId: number) => string;
export interface GlyphCacheOptions {
    /** outline source; throws/invalid JSON propagate to the caller (fallback). */
    provider: GlyphOutlineProvider;
    /** max cached glyphs before LRU eviction (default 4096). */
    maxEntries?: number;
    /**
     * Path2D factory — injectable so tests (and non-DOM environments) can supply
     * a recording double. Defaults to the browser `Path2D` constructor.
     */
    createPath?: () => Path2D;
}
/**
 * LRU cache keyed `(fontId, glyphId)` → `GlyphOutline`. `Map` preserves
 * insertion order, so the oldest key is `keys().next()` and touching an entry
 * is delete-then-set (reinsert at the tail).
 */
export declare class GlyphCache {
    private readonly provider;
    private readonly maxEntries;
    private readonly createPath;
    private readonly entries;
    constructor(options: GlyphCacheOptions);
    /** number of cached glyphs (for tests / eviction assertions). */
    get size(): number;
    /** Drops the cached outlines of these fonts, whose outline data changed. */
    invalidateFonts(fontIds: ReadonlySet<number>): void;
    /**
     * Outline for one glyph, building and caching it on first request. Throws if
     * the provider throws or returns unparseable JSON — the caller falls back.
     */
    get(fontId: number, glyphId: number): GlyphOutline;
    private build;
}
/**
 * Parse the outline-JSON envelope. Throws on unparseable text (an empty string
 * or provider failure lands here) or a non-positive upem, so the backend's
 * whole-run `fillText` fallback engages.
 */
export declare function parseOutlineJson(json: string): {
    upem: number;
    cmds: OutlineCommand[];
};
/**
 * Replay outline commands into a Path2D in FONT UNITS (no scaling here — the
 * backend scales at draw). M→moveTo, L→lineTo, Q→quadraticCurveTo,
 * C→bezierCurveTo, Z→closePath.
 */
export declare function buildGlyphPath(cmds: OutlineCommand[], createPath: () => Path2D): Path2D;
