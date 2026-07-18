// LRU cache of glyph outlines for the canvas backend. Each entry is a
// size-independent `Path2D` built ONCE from a glyph's contours in FONT UNITS
// (y-up, the font convention) — the backend scales it per draw (`size`/upem)
// and flips y, so one cached path serves every size the glyph is painted at.
//
// The outline source is injected (`provider`): production passes
// `outlineGlyphJson` from `../wasm`, tests pass a fake. The provider returns
// the outline as JSON:
//   {"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L",..},
//     {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
//     {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}
// A provider that throws or returns unparseable text propagates the error to
// the caller (the canvas backend catches it and falls back to `fillText` for
// the whole run) — a failed glyph is never cached, so a later working provider
// (e.g. once the wasm export lands) retries cleanly.

/** outline path commands in font units, y-up (font convention). */
export type OutlineCommand =
  | { t: 'M'; x: number; y: number }
  | { t: 'L'; x: number; y: number }
  | { t: 'Q'; cx: number; cy: number; x: number; y: number }
  | { t: 'C'; c1x: number; c1y: number; c2x: number; c2y: number; x: number; y: number }
  | { t: 'Z' };

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

const DEFAULT_MAX_ENTRIES = 4096;

/**
 * LRU cache keyed `(fontId, glyphId)` → `GlyphOutline`. `Map` preserves
 * insertion order, so the oldest key is `keys().next()` and touching an entry
 * is delete-then-set (reinsert at the tail).
 */
export class GlyphCache {
  private readonly provider: GlyphOutlineProvider;
  private readonly maxEntries: number;
  private readonly createPath: () => Path2D;
  private readonly entries = new Map<string, GlyphOutline>();

  constructor(options: GlyphCacheOptions) {
    this.provider = options.provider;
    this.maxEntries = Math.max(1, options.maxEntries ?? DEFAULT_MAX_ENTRIES);
    this.createPath = options.createPath ?? (() => new Path2D());
  }

  /** number of cached glyphs (for tests / eviction assertions). */
  get size(): number {
    return this.entries.size;
  }

  /**
   * Outline for one glyph, building and caching it on first request. Throws if
   * the provider throws or returns unparseable JSON — the caller falls back.
   */
  get(fontId: number, glyphId: number): GlyphOutline {
    const key = `${fontId}:${glyphId}`;
    const cached = this.entries.get(key);
    if (cached !== undefined) {
      // LRU touch: reinsert so the most-recently-used entry moves to the tail
      this.entries.delete(key);
      this.entries.set(key, cached);
      return cached;
    }

    // build BEFORE inserting so a throwing/invalid provider never poisons the
    // cache — the entry is only stored once the outline is successfully built
    const outline = this.build(fontId, glyphId);
    this.entries.set(key, outline);
    if (this.entries.size > this.maxEntries) {
      const oldest = this.entries.keys().next().value;
      if (oldest !== undefined) this.entries.delete(oldest);
    }
    return outline;
  }

  private build(fontId: number, glyphId: number): GlyphOutline {
    const { upem, cmds } = parseOutlineJson(this.provider(fontId, glyphId));
    // empty-contour glyphs (spaces) cache as a null path so the backend skips
    // them without re-fetching
    if (cmds.length === 0) return { path: null, upem };
    return { path: buildGlyphPath(cmds, this.createPath), upem };
  }
}

/**
 * Parse the outline-JSON envelope. Throws on unparseable text (an empty string
 * or provider failure lands here) or a non-positive upem, so the backend's
 * whole-run `fillText` fallback engages.
 */
export function parseOutlineJson(json: string): { upem: number; cmds: OutlineCommand[] } {
  const data = JSON.parse(json) as { upem?: unknown; cmds?: unknown };
  const upem = typeof data.upem === 'number' && data.upem > 0 ? data.upem : NaN;
  if (!Number.isFinite(upem)) {
    throw new Error('glyph outline: missing or non-positive upem');
  }
  const cmds = Array.isArray(data.cmds) ? (data.cmds as OutlineCommand[]) : [];
  return { upem, cmds };
}

/**
 * Replay outline commands into a Path2D in FONT UNITS (no scaling here — the
 * backend scales at draw). M→moveTo, L→lineTo, Q→quadraticCurveTo,
 * C→bezierCurveTo, Z→closePath.
 */
export function buildGlyphPath(cmds: OutlineCommand[], createPath: () => Path2D): Path2D {
  const path = createPath();
  for (const cmd of cmds) {
    switch (cmd.t) {
      case 'M':
        path.moveTo(cmd.x, cmd.y);
        break;
      case 'L':
        path.lineTo(cmd.x, cmd.y);
        break;
      case 'Q':
        path.quadraticCurveTo(cmd.cx, cmd.cy, cmd.x, cmd.y);
        break;
      case 'C':
        path.bezierCurveTo(cmd.c1x, cmd.c1y, cmd.c2x, cmd.c2y, cmd.x, cmd.y);
        break;
      case 'Z':
        path.closePath();
        break;
    }
  }
  return path;
}
