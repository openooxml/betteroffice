// @betteroffice/docx declaration layout/render/displayListQueries.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Synchronous query facade over one built DisplayList, backed by the Rust
 * hit-testing module (`crates/docx-layout/src/hit.rs`).
 *
 * This is the canvas renderer's only source of pointer and selection geometry:
 * every rect comes from the immutable display list, never from DOM rects. A
 * facade is created per display-list build, since the list never mutates in
 * place, and it holds no state a caller has to keep in sync.
 *
 * **Two sources.** A resident editing engine queries its own display list
 * directly and takes no display-list JSON at all. Otherwise the shared layout
 * wasm answers, and the facade prefers a SESSION HANDLE: the list is parsed
 * into the Rust store once (`openDisplayList`) and every later query goes by
 * handle, with no per-query re-serialization. When the session exports are
 * absent it stringifies the list once and reuses that string for the JSON-arg
 * exports. The two paths run the same hit and range logic, so results are
 * byte-identical and only the cost differs.
 *
 * **Handle acquisition is lazy** — on the first query that needs one, or when
 * a host calls `prime()` from idle time — so replacing the facade on every
 * keystroke costs no serialization on the input path. A new facade also tries
 * to ADOPT its predecessor's parsed list, shipping a page delta for the pages
 * that changed instead of reopening the whole list; superseded generations
 * chain their donor, so a deferred pickup still costs one delta. The donor's
 * own later queries degrade to its JSON-arg path.
 *
 * **Handle release.** `dispose()` frees the handle immediately, and ending a
 * facade's line frees every handle the line's facades hold. Only a facade
 * holding a handle is registered with a `FinalizationRegistry`, which covers
 * one dropped without `dispose()` (a `useMemo` replacement, say); the Rust
 * store additionally caps live handles and evicts the oldest, so a missed
 * finalize cannot grow memory without bound.
 *
 * **Failure.** A wasm TRAP poisons the whole instance: a guard held across the
 * query leaks and every later call into it fails, so the engine is marked dead,
 * all queries stop, and `onDisplayListQuerySourceFailure` subscribers are told
 * to rebuild the session. An ordinary error just falls back — a bad handle is
 * dropped and the query retried over JSON.
 *
 * Queries are synchronous and return `null`/`[]` until the lazily imported
 * wasm module resolves. In practice it is already loaded by the time a facade
 * exists, because building the display list went through the same module.
 */
import type { DisplayList } from './displayList';
import { type DisplayListImageRegion, type LocatedImagePrimitive } from './displayListImages';
import { type RustDisplayListQueryEngine } from './rustDisplayList';
/**
 * Query surface of an editing engine that already holds the display list.
 * None of these take display-list JSON, so this source never opens a handle
 * and never serializes anything.
 */
export interface ResidentDisplayListQueryEngine {
    displayHitTestRegionsJson(pageIndex: number, x: number, y: number): string;
    displayVerticalMoveJson(position: number, direction: 'up' | 'down', goalX: number): string;
    displayRangeRectsJson(from: number, to: number): string;
    /** @internal */
    displayRangeRectsOnPagesJson?(from: number, to: number, firstPage: number, lastPage: number): string;
    displayRangeRectsRegionJson(region: DisplayListHitRegion, partId: string, from: number, to: number): string;
}
/** which part of a page owns a hit — mirrors `HitRegion` in hit.rs */
export type DisplayListHitRegion = 'body' | 'header' | 'footer' | 'footnote' | 'endnote';
/**
 * What a click at a hit point would act on — mirrors `HoverTarget` in hit.rs.
 * `'text'` is the typeable area (a run's box, or the content box around it),
 * which is what a pointer cursor keys off; `pos` cannot say, since it resolves
 * everywhere on a page that carries text.
 */
export type DisplayListHoverTarget = 'text' | 'image' | 'none';
/**
 * Region-aware hit result. For `header`/`footer` the position refers to the
 * header/footer document identified by `rId`, and for `footnote`/`endnote` to
 * the note story named by `noteId` — NOT the body document, so the caller must
 * route the selection to that editor. `pos` is null when the point is inside
 * the region but resolves to no position.
 */
export interface DisplayListRegionHit {
    region: DisplayListHitRegion;
    rId?: string;
    /** note whose story a `footnote`/`endnote` position addresses */
    noteId?: number;
    pos: number | null;
    /** Absent from a wasm build predating it; read that as "not text". */
    target?: DisplayListHoverTarget;
}
/**
 * Result of an up/down caret move. `goalX` is the page-local x to feed back
 * into the next move so a run of them holds one column.
 */
export interface DisplayListVerticalMove {
    position: number;
    goalX: number;
}
/** one highlight rectangle of a document range, page-local px */
export interface DisplayListRect {
    pageIndex: number;
    x: number;
    y: number;
    width: number;
    height: number;
}
/** Explicit lifecycle of the mandatory Rust query source. */
export type DisplayListQuerySourceState = {
    status: 'loading';
} | {
    status: 'ready';
} | {
    status: 'error';
    error: Error;
};
/** One paragraph fragment box on one page, in page-local px. */
export interface DisplayListParagraphGeometry extends DisplayListRect {
    from: number;
    to: number;
    blockId?: number | string;
    paraId?: string;
}
/** One ordered visual line reconstructed from authoritative primitives. */
export interface VisualLineExtent {
    top: number;
    bottom: number;
}
export interface DisplayListVisualLine extends DisplayListRect {
    baseline: number;
    from: number;
    to: number;
    blockId?: number | string;
    paraId?: string;
}
/** Image primitive plus its explicit page/region geometry. */
export interface DisplayListImageGeometry extends LocatedImagePrimitive {
    rect: DisplayListRect;
    pos: number;
}
/**
 * Sync query surface over one immutable DisplayList. A new instance is
 * created per display-list build (the list never mutates in place).
 */
export interface DisplayListQueries {
    /** the list this instance queries (page sizes, region bands, …) */
    readonly displayList: DisplayList;
    /** false until the wasm module is loaded — queries return null/[] before */
    isReady(): boolean;
    /** Loading/ready/error state for hosts that must not silently fall back. */
    sourceState(): DisplayListQuerySourceState;
    /** Resolves when the Rust query engine is ready; rejects on load failure. */
    whenReady(): Promise<void>;
    pageCount(): number;
    pageSize(pageIndex: number): {
        width: number;
        height: number;
    } | null;
    pageBounds(pageIndex: number): DisplayListRect | null;
    contentBounds(pageIndex: number): DisplayListRect | null;
    columnBounds(pageIndex: number): DisplayListRect[];
    /** Body paragraph fragment boxes containing `pos`, including page splits. */
    paragraphRects(pos: number): DisplayListParagraphGeometry[];
    /** Ordered body visual lines across all pages. */
    visualLines(): readonly DisplayListVisualLine[];
    /** The part of {@link visualLines} on one page, computed for that page alone. */
    visualLinesOnPage(pageIndex: number): readonly DisplayListVisualLine[];
    /** Vertical span of the page's visual lines in page coordinates, or null when it has none. */
    visualLineExtent(pageIndex: number): VisualLineExtent | null;
    /** Visual line containing `pos`, or null. */
    visualLineAtPosition(pos: number): DisplayListVisualLine | null;
    /** Topmost image under a page-local point. Body by default. */
    imageAtPoint(pageIndex: number, x: number, y: number, region?: DisplayListImageRegion, rId?: string): DisplayListImageGeometry | null;
    /** Image whose atom starts at `pos`. Body by default. */
    imageByPos(pos: number, region?: DisplayListImageRegion, rId?: string): DisplayListImageGeometry | null;
    /** region-aware point → doc position (page-local coordinates) */
    hitTestRegions(pageIndex: number, x: number, y: number): DisplayListRegionHit | null;
    verticalMove(position: number, direction: 'up' | 'down', goalX?: number): DisplayListVerticalMove | null;
    /** body document range → highlight rects */
    rangeRects(from: number, to: number): DisplayListRect[];
    /** @internal */
    rangeRectsOnPages?(from: number, to: number, firstPage: number, lastPage: number): DisplayListRect[];
    /**
     * Header/footer document range → highlight rects for the region's band. `region` is
     * `'header' | 'footer'`; `rId` identifies the HF doc, and
     * `from`/`to` are positions in THAT doc. The same HF doc paints on every page
     * carrying the part, so this returns one rect-set per such page (each tagged
     * with its `pageIndex`) — the caller picks the page it is editing. Returns
     * `[]` when the region-aware exports are absent, which is feature-detected.
     */
    hfRangeRects(region: 'header' | 'footer', rId: string, from: number, to: number): DisplayListRect[];
    /**
     * Note document range → highlight rects for that note's story. `from`/`to`
     * are positions in the `fn:{noteId}` / `en:{noteId}` document, never the
     * body's. A note paints on one page only, so every rect shares its
     * `pageIndex`.
     */
    noteRangeRects(region: 'footnote' | 'endnote', noteId: number, from: number, to: number): DisplayListRect[];
    /**
     * Caret geometry for a collapsed HF selection — the HF twin of `caretRect`.
     * Resolves `[pos, pos+1)` in the HF doc (left edge is the caret), falling back
     * to `[pos-1, pos)` (right edge) at end-of-line / end-of-doc. Returns one
     * caret rect per page carrying the part; the caller picks the edited page.
     */
    hfCaretRects(region: 'header' | 'footer', rId: string, pos: number): DisplayListRect[];
    /**
     * Caret geometry for a collapsed selection in a note story — the note twin of
     * `hfCaretRects`. A note paints on one page, so this returns at most one rect.
     */
    noteCaretRects(region: 'footnote' | 'endnote', noteId: number, pos: number): DisplayListRect[];
    /** Header/footer sidebar anchors, one per page carrying the part. */
    hfAnchorRects(region: 'header' | 'footer', rId: string, pos: number): DisplayListRect[];
    /**
     * Caret geometry for a collapsed body selection: the collapsed-range rect.
     * Resolves `[pos, pos+1)` first (rect's left edge is the caret), then falls
     * back to `[pos-1, pos)` using the right edge (end-of-doc / end-of-line).
     */
    caretRect(pos: number): DisplayListRect | null;
    /**
     * Anchor geometry for sidebar markers: like `caretRect` but scans
     * `[pos, pos+2)` forward first so *node* positions (paragraph/table
     * markers carrying structural tracked-change attrs) resolve to their first
     * content line instead of the previous block's tail.
     */
    anchorRect(pos: number): DisplayListRect | null;
    /** Explicit body-sidebar alias retained alongside `anchorRect`. */
    sidebarAnchorRect(pos: number): DisplayListRect | null;
    /**
     * Acquire the Rust session handle now (adopting the donor facade's parsed
     * list when possible). Optional: the first query acquires it on demand;
     * hosts call this from idle time to keep serialization off interaction
     * paths. No-op once attempted, superseded, or disposed.
     */
    prime(): void;
    /**
     * Release the wasm session handle backing this facade. Idempotent; safe to
     * call even when no handle was opened (JSON-arg fallback path). Callers that
     * forget are covered by a `FinalizationRegistry`, but disposing eagerly frees
     * the parsed display list in the Rust store immediately.
     */
    dispose(): void;
}
/**
 * Ends `line`: closes the store handles its facades hold, which then query
 * over JSON, and its superseded facades answer nothing until a facade of the
 * same line is built. Call it when the session behind the line goes away.
 */
export declare function endDisplayListQueriesLine(line: object): void;
type DisplayListQuerySource = RustDisplayListQueryEngine | ResidentDisplayListQueryEngine;
/** True once a query trapped in this engine's wasm instance. */
export declare function isDisplayListQuerySourceDead(engine: DisplayListQuerySource | null | undefined): boolean;
/** Subscribe to wasm traps so the host can rebuild the dead session. */
export declare function onDisplayListQuerySourceFailure(listener: (error: Error) => void): () => void;
/**
 * Build a query facade for one display list. The optional `engine` makes the
 * facade synchronous and deterministic in tests; without it the shared wasm
 * module is loaded lazily and queries no-op (`null`/`[]`) until it resolves.
 *
 * `previous` (the facade this build replaces) enables handle adoption: instead
 * of re-serializing and re-parsing the WHOLE list into the Rust store, the new
 * facade takes over the previous parsed list and patches only the pages that
 * changed. The donor facade's remaining queries degrade to its own JSON-arg
 * path (same stale-list semantics it always had after replacement).
 *
 * `line` names the document the list lays out, such as its load. A facade of another
 * document starts a new line and ends the previous one, so a facade of the
 * replaced document never answers from the new one. Without `line` a superseded
 * facade answers from its own list.
 */
export declare function createDisplayListQueries(list: DisplayList, engine?: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine, previous?: DisplayListQueries | null, line?: object | null): DisplayListQueries;
export {};
