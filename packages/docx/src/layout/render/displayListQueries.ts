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
 * **Handle release.** `dispose()` frees the handle immediately. A facade
 * dropped without it — a `useMemo` replacement, say — is covered by a
 * `FinalizationRegistry`, and the Rust store additionally caps live handles
 * and evicts the oldest, so a missed finalize cannot grow memory without
 * bound.
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

import type { DisplayList, DisplayPage, DisplayPrimitive } from './displayList';
import { displayPageRevision, displayPageShiftsSince, type DisplayPageShift } from './frameDelta';
import { displayPrimitiveRect, type GeoRect } from './displayListGeometry';
import {
  findImagePrimitiveAtPoint,
  findImagePrimitiveByDocPos,
  type DisplayListImageRegion,
  type LocatedImagePrimitive,
} from './displayListImages';
import { loadRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from './rustDisplayList';

/**
 * Query surface of an editing engine that already holds the display list.
 * None of these take display-list JSON, so this source never opens a handle
 * and never serializes anything.
 */
export interface ResidentDisplayListQueryEngine {
  displayHitTestRegionsJson(pageIndex: number, x: number, y: number): string;
  displayVerticalMoveJson(
    position: number,
    direction: 'up' | 'down',
    goalX: number
  ): string;
  displayRangeRectsJson(from: number, to: number): string;
  displayRangeRectsRegionJson(
    region: DisplayListHitRegion,
    partId: string,
    from: number,
    to: number
  ): string;
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
export type DisplayListQuerySourceState =
  | { status: 'loading' }
  | { status: 'ready' }
  | { status: 'error'; error: Error };

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
  pageSize(pageIndex: number): { width: number; height: number } | null;
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
  imageAtPoint(
    pageIndex: number,
    x: number,
    y: number,
    region?: DisplayListImageRegion,
    rId?: string
  ): DisplayListImageGeometry | null;
  /** Image whose atom starts at `pos`. Body by default. */
  imageByPos(
    pos: number,
    region?: DisplayListImageRegion,
    rId?: string
  ): DisplayListImageGeometry | null;
  /** region-aware point → doc position (page-local coordinates) */
  hitTestRegions(pageIndex: number, x: number, y: number): DisplayListRegionHit | null;
  verticalMove(
    position: number,
    direction: 'up' | 'down',
    goalX?: number
  ): DisplayListVerticalMove | null;
  /** body document range → highlight rects */
  rangeRects(from: number, to: number): DisplayListRect[];
  /**
   * Header/footer document range → highlight rects for the region's band. `region` is
   * `'header' | 'footer'`; `rId` identifies the HF doc, and
   * `from`/`to` are positions in THAT doc. The same HF doc paints on every page
   * carrying the part, so this returns one rect-set per such page (each tagged
   * with its `pageIndex`) — the caller picks the page it is editing. Returns
   * `[]` when the region-aware exports are absent, which is feature-detected.
   */
  hfRangeRects(
    region: 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): DisplayListRect[];
  /**
   * Note document range → highlight rects for that note's story. `from`/`to`
   * are positions in the `fn:{noteId}` / `en:{noteId}` document, never the
   * body's. A note paints on one page only, so every rect shares its
   * `pageIndex`.
   */
  noteRangeRects(
    region: 'footnote' | 'endnote',
    noteId: number,
    from: number,
    to: number
  ): DisplayListRect[];
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

// FinalizationRegistry is ES2021; the core tsconfig `lib` may predate it, so it
// is referenced through a minimal local shape via `globalThis` rather than
// widening the lib. Available at runtime in every target (modern browsers, Node,
// Bun).
interface HandleFinalizationRegistry {
  register(target: object, heldValue: () => void, unregisterToken?: object): void;
  unregister(unregisterToken: object): void;
}
type HandleFinalizationRegistryCtor = new (
  cleanup: (heldValue: () => void) => void
) => HandleFinalizationRegistry;

/**
 * Closes session handles for facades dropped without an explicit `dispose()`
 * (the held value is a bound close-thunk that never references the facade
 * object, so registering it can't keep the facade alive). Null in environments
 * without `FinalizationRegistry` — the Rust store's handle cap is the hard
 * backstop there.
 */
const handleFinalizers: HandleFinalizationRegistry | null = (() => {
  const Ctor = (globalThis as unknown as { FinalizationRegistry?: HandleFinalizationRegistryCtor })
    .FinalizationRegistry;
  return Ctor ? new Ctor((close) => close()) : null;
})();

/** A facade's Rust store handle and the engine that owns it. */
interface HandleCell {
  handle: number | null;
  eng: RustDisplayListQueryEngine | null;
}

/**
 * The finalizer's held value. Built outside `createDisplayListQueries`: a
 * closure made there shares the factory's context, whose donor facade can lead
 * back to the registered facade, which then is never finalized. Only an engine
 * holds a handle (a resident source never opens one), so `cell.eng` is the
 * source the dead check needs.
 */
function handleCloser(cell: HandleCell): () => void {
  return () => {
    const handle = cell.handle;
    cell.handle = null;
    if (handle === null || !cell.eng || deadSources.has(cell.eng)) return;
    try {
      cell.eng.closeDisplayList?.(handle);
    } catch {
      // a close failure must never surface; the store caps handles anyway
    }
  };
}

/**
 * Internal handoff state for handle adoption between consecutive facades.
 * Keyed weakly so a dropped facade can never leak its list.
 */
interface FacadeDeltaSeed {
  list: DisplayList;
  /** Per-page mutation revisions as parsed into the Rust store (null before
   * a handle opened or adopted). Owned frame deltas patch through the same
   * page objects, so identity alone cannot prove a page matches the store. */
  storeRevisions(): readonly number[] | null;
  engine(): RustDisplayListQueryEngine | null;
  /** Relinquish the live handle (the donor's queries fall back to JSON-arg). */
  takeHandle(): number | null;
  hasHandle(): boolean;
  /** Nearest ancestor facade that held the handle when this one was created. */
  donor(): DisplayListQueries | null;
  /** A successor now owns this generation: never open/adopt a handle here. */
  supersede(): void;
  /** The newest facade of this facade's line of successive layouts. */
  lineage: FacadeLineage;
  /** The document the line lays out, as the caller named it. */
  line: object | null;
  disposed(): boolean;
}

/** Shared by a facade and every facade built from it; weak, so no layout outlives its holders. */
interface FacadeLineage {
  newest: WeakTo<DisplayListQueries> | null;
}

type WeakTo<T> = { deref(): T | undefined };

const WeakRefCtor = (globalThis as { WeakRef?: new <T extends object>(target: T) => WeakTo<T> })
  .WeakRef;

/** A native WeakRef: a wrapping closure would share a context that keeps the target alive. */
function weakly(queries: DisplayListQueries): WeakTo<DisplayListQueries> {
  return WeakRefCtor ? new WeakRefCtor(queries) : strongly(queries);
}

function strongly(queries: DisplayListQueries): WeakTo<DisplayListQueries> {
  return { deref: () => queries };
}

const facadeDeltaSeeds = new WeakMap<DisplayListQueries, FacadeDeltaSeed>();

/** One lineage per named document line, so a facade built after a gap in the chain rejoins it. */
const lineages = new WeakMap<object, FacadeLineage>();

/**
 * Ends `line`: its superseded facades answer nothing until a facade of the same
 * line is built. Call it when the session behind the line goes away.
 */
export function endDisplayListQueriesLine(line: object): void {
  const lineage = lineages.get(line);
  if (lineage) lineage.newest = null;
}

type DisplayListQuerySource = RustDisplayListQueryEngine | ResidentDisplayListQueryEngine;

/**
 * Engines whose wasm instance trapped. A trap aborts without running
 * destructors, so a `RefCell` guard held across the query leaks and every
 * later call into that instance fails — the instance must be rebuilt, and
 * until then nothing may query it.
 */
const deadSources = new WeakSet<DisplayListQuerySource>();
const sourceFailureListeners = new Set<(error: Error) => void>();

/** True once a query trapped in this engine's wasm instance. */
export function isDisplayListQuerySourceDead(
  engine: DisplayListQuerySource | null | undefined
): boolean {
  return engine !== null && engine !== undefined && deadSources.has(engine);
}

/** Subscribe to wasm traps so the host can rebuild the dead session. */
export function onDisplayListQuerySourceFailure(listener: (error: Error) => void): () => void {
  sourceFailureListeners.add(listener);
  return () => {
    sourceFailureListeners.delete(listener);
  };
}

/** A trap (panic) rather than a returned `Err`, which arrives as a string. */
function isWasmTrap(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  if (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError) return true;
  return error.name === 'RuntimeError';
}

type StoreShiftRun = [start: number, count: number, mask: number, delta: number];
type StoreNoteAnchor = [area: number, note: number, start: number | null, end: number | null];

/** A store page not parsed yet: its slot and size, but no primitives. */
const UNLOADED = -1;

function placeholderPage(page: DisplayPage): Pick<DisplayPage, 'pageIndex' | 'width' | 'height' | 'primitives'> {
  return { pageIndex: page.pageIndex, width: page.width, height: page.height, primitives: [] };
}

/** The baseline of a primitive that forms part of a visual line, or null. */
function visualLineBaseline(primitive: DisplayPrimitive): number | null {
  if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') return null;
  if (primitive.docStart === undefined || primitive.docEnd === undefined) return null;
  if (primitive.kind === 'glyphRun' && primitive.glyphs.length === 0) return null;
  const baseline =
    primitive.kind === 'text'
      ? primitive.baselineY
      : primitive.glyphs.reduce((max, glyph) => Math.max(max, glyph.y), -Infinity);
  return Number.isFinite(baseline) ? baseline : null;
}

// Geometry only: an owned position shift keeps the page object and moves no
// geometry, so an extent stays valid for as long as its page object lives.
const visualLineExtents = new WeakMap<DisplayPage, VisualLineExtent | null>();

/** Vertical span of a page's visual lines in page coordinates, or null when it has none. */
function visualLineExtent(page: DisplayPage): VisualLineExtent | null {
  if (visualLineExtents.has(page)) return visualLineExtents.get(page) ?? null;
  let top = Infinity;
  let bottom = -Infinity;
  for (const primitive of page.primitives) {
    if (visualLineBaseline(primitive) === null) continue;
    const rect = displayPrimitiveRect(primitive);
    top = Math.min(top, rect.y);
    bottom = Math.max(bottom, rect.y + rect.h);
  }
  const extent = top <= bottom ? { top, bottom } : null;
  visualLineExtents.set(page, extent);
  return extent;
}

/**
 * Page-delta between two display lists, exploiting the retained-frame
 * invariant that unchanged pages keep object identity across builds. A page
 * whose identity and in-place mutation revision are unchanged since the store
 * parsed it is reused outright; a page that only accumulated recorded
 * position shifts ships those shifts as compact ops the store replays. Any
 * other page becomes an unloaded placeholder that is parsed when a query first
 * needs it. Returns null when nothing is reusable (a fresh open costs the
 * same).
 */
function buildDisplayListUpdateJson(
  seed: FacadeDeltaSeed,
  next: DisplayList
): { json: string; revisions: number[] } | null {
  const storeRevisions = seed.storeRevisions();
  if (!storeRevisions) return null;
  const previousIndex = new Map<unknown, number>();
  seed.list.pages.forEach((page, index) => previousIndex.set(page, index));
  const reuse: Array<[number, number]> = [];
  const replace: Array<[number, unknown]> = [];
  const shift: Array<
    [number, number, StoreShiftRun[][]] | [number, number, StoreShiftRun[][], StoreNoteAnchor[][]]
  > = [];
  const revisions: number[] = [];
  next.pages.forEach((page, index) => {
    const from = previousIndex.get(page);
    if (from !== undefined) previousIndex.delete(page);
    const storeRevision = from === undefined ? undefined : storeRevisions[from];
    if (from === undefined || storeRevision === undefined) {
      replace.push([index, placeholderPage(page)]);
      revisions.push(UNLOADED);
      return;
    }
    if (storeRevision === UNLOADED) {
      reuse.push([index, from]);
      revisions.push(UNLOADED);
      return;
    }
    const shifts =
      displayPageRevision(page) === storeRevision
        ? []
        : displayPageShiftsSince(page, storeRevision);
    if (shifts === null) {
      replace.push([index, placeholderPage(page)]);
      revisions.push(UNLOADED);
    } else if (shifts.length === 0) {
      reuse.push([index, from]);
      revisions.push(storeRevision);
    } else {
      const runLists = shifts.map((step: DisplayPageShift) =>
        step.runs.map((run): StoreShiftRun => [run.start, run.count, run.changedMask, run.delta])
      );
      if (shifts.some((step: DisplayPageShift) => step.anchors.length > 0)) {
        const anchorLists = shifts.map((step: DisplayPageShift) =>
          step.anchors.map(
            (anchor): StoreNoteAnchor => [anchor.area, anchor.note, anchor.start, anchor.end]
          )
        );
        shift.push([index, from, runLists, anchorLists]);
      } else {
        shift.push([index, from, runLists]);
      }
      revisions.push(displayPageRevision(page));
    }
  });
  if (reuse.length === 0 && shift.length === 0) return null;
  return {
    json: JSON.stringify({
      total: next.pages.length,
      ...(next.contractVersion !== undefined ? { contractVersion: next.contractVersion } : {}),
      reuse,
      replace,
      ...(shift.length > 0 ? { shift } : {}),
    }),
    revisions,
  };
}

/** Lowest and highest body document position a page paints. */
interface PagePositionSpan {
  revision: number;
  min: number;
  max: number;
}

const pagePositionSpans = new WeakMap<DisplayPage, PagePositionSpan>();

function pagePositionSpan(page: DisplayPage): PagePositionSpan {
  const revision = displayPageRevision(page);
  const cached = pagePositionSpans.get(page);
  if (cached && cached.revision === revision) return cached;
  let min = Infinity;
  let max = -Infinity;
  const include = (value: number | null | undefined): void => {
    if (typeof value !== 'number') return;
    if (value < min) min = value;
    if (value > max) max = value;
  };
  for (const primitive of page.primitives) {
    include(primitive.docStart);
    include(primitive.docEnd);
    include(primitive.fragmentDocStart);
    include(primitive.fragmentDocEnd);
    include(primitive.inlineSdtWidget?.pos);
  }
  const span = { revision, min, max };
  pagePositionSpans.set(page, span);
  return span;
}

/**
 * Pages whose body positions can answer a query over `[from, to]`, plus
 * `spread` pages on either side of each (a vertical move reads neighbours).
 */
function pagesTouchingPositions(
  list: DisplayList,
  from: number,
  to: number,
  spread = 0
): number[] {
  const lower = Math.min(from, to) - 1;
  const upper = Math.max(from, to) + 1;
  const pages = new Set<number>();
  list.pages.forEach((page, index) => {
    const span = pagePositionSpan(page);
    if (span.min > upper || span.max < lower) return;
    for (let offset = -spread; offset <= spread; offset += 1) {
      const neighbour = index + offset;
      if (neighbour >= 0 && neighbour < list.pages.length) pages.add(neighbour);
    }
  });
  return [...pages];
}

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
export function createDisplayListQueries(
  list: DisplayList,
  engine?: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine,
  previous?: DisplayListQueries | null,
  line?: object | null
): DisplayListQueries {
  let json: string | null = null;
  const getJson = (): string => {
    json ??= JSON.stringify(list);
    return json;
  };
  // Revisions of the pages as parsed into the Rust store, UNLOADED for a
  // placeholder; null until a handle is opened or adopted. Kept exact so shift
  // replay can never double-apply.
  let storeRevisions: number[] | null = null;

  const resident: ResidentDisplayListQueryEngine | null = isResidentQueryEngine(engine)
    ? engine
    : null;
  // The finalizer holds this cell, so it must never lead to a facade.
  const cell: HandleCell = {
    handle: null,
    eng: resident ? null : ((engine as RustDisplayListQueryEngine | undefined) ?? null),
  };
  let sourceError: Error | null = null;
  let resolveReady!: () => void;
  let rejectReady!: (error: Error) => void;
  const readyPromise = new Promise<void>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });
  // Consumers opt into the rejection through whenReady(); keep an unobserved
  // lazy source failure from becoming a global unhandled-rejection event.
  void readyPromise.catch(() => undefined);

  // session-handle state: the parsed display list lives in the Rust store behind
  // `cell.handle`; null means the JSON-arg fallback (unsupported wasm, open failed,
  // or a handle dropped after a stale-handle error).
  let handleAttempted = false;
  let superseded = false;
  let disposed = false;
  // lets dispose() cancel the finalizer below so a handle is never double-closed
  const finalizerToken = {};

  // The adoption donor: `previous` when it holds the handle, otherwise the
  // donor `previous` itself recorded — pre-collapsed to one hop so a typing
  // burst retains at most one superseded list. Marking the predecessor
  // superseded keeps stale-closure queries on it from stealing the handle out
  // of the chain.
  let donorFacade: DisplayListQueries | null = null;
  const previousSeed = previous ? facadeDeltaSeeds.get(previous) : undefined;
  if (previousSeed) {
    donorFacade = previousSeed.hasHandle() ? previous! : previousSeed.donor();
    previousSeed.supersede();
  }
  if (previousSeed && previousSeed.line !== (line ?? null)) previousSeed.lineage.newest = null;
  let lineage: FacadeLineage = { newest: null };
  if (line) {
    lineage = lineages.get(line) ?? lineage;
    lineages.set(line, lineage);
  }

  const source = (): DisplayListQuerySource | null => resident ?? cell.eng;

  const isDead = (): boolean => {
    const current = source();
    return current !== null && deadSources.has(current);
  };

  /**
   * A trap poisons the whole wasm instance, so record it, stop querying, and
   * tell the host to rebuild. The leaked handle is abandoned rather than
   * closed — closing would re-enter the dead instance.
   */
  const killSource = (label: string, error: unknown): void => {
    const failure = error instanceof Error ? error : new Error(String(error));
    sourceError = failure;
    cell.handle = null;
    handleFinalizers?.unregister(finalizerToken);
    const current = source();
    if (!current || deadSources.has(current)) return;
    deadSources.add(current);
    console.error(
      `[CanvasRenderer] ${label} trapped in wasm; the session is unusable and must be rebuilt`,
      failure
    );
    for (const listener of sourceFailureListeners) {
      try {
        listener(failure);
      } catch (listenerError) {
        console.error('[CanvasRenderer] display-list source failure listener threw', listenerError);
      }
    }
  };

  const closeHandle = handleCloser(cell);

  // adopt the donor facade's parsed list when only some pages changed:
  // ships a page-delta into the Rust store instead of the whole list
  const adoptHandle = (): boolean => {
    const donor = donorFacade;
    donorFacade = null;
    if (!donor || !cell.eng?.updateDisplayList || !cell.eng.hasDisplayListUpdate?.()) return false;
    const seed = facadeDeltaSeeds.get(donor);
    if (!seed || seed.engine() !== cell.eng) return false;
    const update = buildDisplayListUpdateJson(seed, list);
    if (!update) return false;
    const adopted = seed.takeHandle();
    if (adopted === null) return false;
    try {
      cell.eng.updateDisplayList(adopted, update.json);
      cell.handle = adopted;
      storeRevisions = update.revisions;
      return true;
    } catch (error) {
      // the Rust side closes the handle on a failed update; close defensively
      // anyway (idempotent) in case the failure happened before wasm ran, then
      // fall through to a fresh full open
      try {
        cell.eng.closeDisplayList?.(adopted);
      } catch {
        // the capped store reclaims it eventually
      }
      console.warn('[CanvasRenderer] display-list delta update failed; reopening', error);
      return false;
    }
  };

  // acquire at most one handle, on the first query that wants it (or via
  // prime()); a failure leaves `handle` null so queries take the JSON-arg path
  const openHandle = (): void => {
    if (
      disposed ||
      superseded ||
      handleAttempted ||
      cell.handle !== null ||
      !cell.eng ||
      isDead()
    ) {
      return;
    }
    if (!cell.eng.hasDisplayListSession?.() || !cell.eng.openDisplayList) return;
    handleAttempted = true;
    if (adoptHandle()) return;
    try {
      if (cell.eng.updateDisplayList && cell.eng.hasDisplayListUpdate?.()) {
        // Pages are parsed into the store when a query first needs them, so
        // opening costs their sizes only.
        cell.handle = cell.eng.openDisplayList(
          JSON.stringify({
            ...(list.contractVersion !== undefined
              ? { contractVersion: list.contractVersion }
              : {}),
            pages: list.pages.map(placeholderPage),
          })
        );
        storeRevisions = list.pages.map(() => UNLOADED);
      } else {
        cell.handle = cell.eng.openDisplayList(getJson());
        storeRevisions = list.pages.map(displayPageRevision);
      }
    } catch (error) {
      cell.handle = null;
      if (isWasmTrap(error)) {
        killSource('display-list session open', error);
        return;
      }
      console.warn(
        '[CanvasRenderer] display-list session open failed; using JSON-arg queries',
        error
      );
    }
  };

  // Parse the pages a query reads into the store. A failed update closes the
  // handle on the Rust side; the query then takes the JSON-arg path.
  const ensurePages = (pageIndices: readonly number[]): void => {
    if (cell.handle === null || !storeRevisions || !cell.eng?.updateDisplayList) return;
    const revisions = storeRevisions;
    const replace: Array<[number, DisplayPage]> = [];
    for (const index of pageIndices) {
      const page = list.pages[index];
      if (page && revisions[index] !== displayPageRevision(page)) replace.push([index, page]);
    }
    if (replace.length === 0) return;
    try {
      cell.eng.updateDisplayList(
        cell.handle,
        JSON.stringify({
          total: list.pages.length,
          ...(list.contractVersion !== undefined
            ? { contractVersion: list.contractVersion }
            : {}),
          keep: true,
          replace,
        })
      );
      for (const [index, page] of replace) revisions[index] = displayPageRevision(page);
    } catch (error) {
      if (isWasmTrap(error)) {
        killSource('display-list page load', error);
        storeRevisions = null;
        return;
      }
      // Rust closes the handle on a failed update; close defensively in case
      // the failure happened before wasm ran.
      closeHandle();
      handleFinalizers?.unregister(finalizerToken);
      storeRevisions = null;
      console.warn('[CanvasRenderer] display-list page load failed; using JSON-arg queries', error);
    }
  };

  const allPages = (): number[] => list.pages.map((_, index) => index);

  if (resident || cell.eng) {
    resolveReady();
  } else {
    loadRustDisplayListQueryEngine().then(
      (loaded) => {
        cell.eng = loaded;
        resolveReady();
      },
      (error) => {
        sourceError = error instanceof Error ? error : new Error(String(error));
        rejectReady(sourceError);
        console.warn('[CanvasRenderer] display-list query engine failed to load', sourceError);
      }
    );
  }

  // run a query, preferring the session handle and falling back to the JSON-arg
  // path on any by-handle failure (a stale/evicted handle, or the by-handle
  // export missing). A bad handle is dropped so later calls skip straight to
  // JSON-arg. Returns the raw JSON string, or null when no engine is ready.
  const runQuery = (
    byHandle: ((h: number) => string) | undefined,
    byJson: () => string,
    label: string,
    pages: () => readonly number[]
  ): string | null => {
    if (!cell.eng || isDead()) return null;
    if (cell.handle === null) openHandle();
    if (cell.handle !== null && byHandle) ensurePages(pages());
    if (isDead()) return null;
    if (cell.handle !== null && byHandle) {
      try {
        return byHandle(cell.handle);
      } catch (error) {
        if (isWasmTrap(error)) {
          killSource(label, error);
          return null;
        }
        console.warn(`[CanvasRenderer] ${label} session query failed; falling back`, error);
        closeHandle();
      }
    }
    try {
      return byJson();
    } catch (error) {
      if (isWasmTrap(error)) {
        killSource(label, error);
        return null;
      }
      sourceError = error instanceof Error ? error : new Error(String(error));
      console.warn(`[CanvasRenderer] ${label} query failed`, error);
      return null;
    }
  };

  const parseQuery = <T>(raw: string | null, fallback: T, label: string): T => {
    if (raw === null) return fallback;
    try {
      return JSON.parse(raw) as T;
    } catch (error) {
      sourceError = error instanceof Error ? error : new Error(String(error));
      console.warn(`[CanvasRenderer] ${label} returned invalid JSON`, error);
      return fallback;
    }
  };

  /**
   * Where a superseded facade's queries go once its successor holds the handle:
   * the newest live facade, so a stale caller gets the current layout instead of
   * this whole list serialised for every call. Undefined while this facade
   * answers itself; null when no live facade is left, which answers nothing.
   */
  const handedOff = (): DisplayListQueries | null | undefined => {
    if (resident || !superseded || cell.handle !== null || !line) return undefined;
    // never names this facade: the finalizer's closure context must not retain it
    const newest = lineage.newest?.deref();
    return newest && !facadeDeltaSeeds.get(newest)?.disposed() ? newest : null;
  };

  /** A list-only read, answered from the live facade once this one handed off its handle. */
  const viaLive =
    <A extends unknown[], R>(
      local: (...args: A) => R,
      remote: (live: DisplayListQueries, ...args: A) => R
    ) =>
    (...args: A): R => {
      const live = handedOff();
      return live ? remote(live, ...args) : local(...args);
    };

  const residentQuery = (query: () => string, label: string): string | null => {
    if (isDead()) return null;
    try {
      return query();
    } catch (error) {
      if (isWasmTrap(error)) {
        killSource(`resident ${label}`, error);
        return null;
      }
      sourceError = error instanceof Error ? error : new Error(String(error));
      console.warn(`[CanvasRenderer] resident ${label} query failed`, error);
      return null;
    }
  };

  const hitTestRegions = (pageIndex: number, x: number, y: number): DisplayListRegionHit | null => {
    const live = handedOff();
    if (live !== undefined) return live?.hitTestRegions(pageIndex, x, y) ?? null;
    if (resident) {
      return parseQuery(
        residentQuery(
          () => resident.displayHitTestRegionsJson(pageIndex, x, y),
          'hit_test_regions'
        ),
        null,
        'hit_test_regions'
      );
    }
    const raw = runQuery(
      cell.eng?.hitTestRegionsByHandle &&
        ((h: number) => cell.eng!.hitTestRegionsByHandle!(h, pageIndex, x, y)),
      () => cell.eng!.hitTestRegionsJson(getJson(), pageIndex, x, y),
      'hit_test_regions',
      () => [pageIndex]
    );
    return parseQuery(raw, null, 'hit_test_regions');
  };

  const rangeRects = (from: number, to: number): DisplayListRect[] => {
    const live = handedOff();
    if (live !== undefined) return live?.rangeRects(from, to) ?? [];
    if (resident) {
      return parseQuery(
        residentQuery(() => resident.displayRangeRectsJson(from, to), 'range_rects'),
        [],
        'range_rects'
      );
    }
    const raw = runQuery(
      cell.eng?.rangeRectsByHandle && ((h: number) => cell.eng!.rangeRectsByHandle!(h, from, to)),
      () => cell.eng!.rangeRectsJson(getJson(), from, to),
      'range_rects',
      () => pagesTouchingPositions(list, from, to)
    );
    return parseQuery(raw, [], 'range_rects');
  };

  const verticalMove = (
    position: number,
    direction: 'up' | 'down',
    goalX?: number
  ): DisplayListVerticalMove | null => {
    const live = handedOff();
    if (live !== undefined) return live?.verticalMove(position, direction, goalX) ?? null;
    const resolvedGoalX = goalX ?? Number.NaN;
    if (resident) {
      return parseQuery(
        residentQuery(
          () => resident.displayVerticalMoveJson(position, direction, resolvedGoalX),
          'vertical_move'
        ),
        null,
        'vertical_move'
      );
    }
    if (!cell.eng?.verticalMoveJson) return null;
    const raw = runQuery(
      cell.eng.verticalMoveByHandle &&
        ((h: number) => cell.eng!.verticalMoveByHandle!(h, position, direction, resolvedGoalX)),
      () => cell.eng!.verticalMoveJson!(getJson(), position, direction, resolvedGoalX),
      'vertical_move',
      () => pagesTouchingPositions(list, position, position, 1)
    );
    return parseQuery(raw, null, 'vertical_move');
  };

  // The one scoped range-rect path. `partId` names the instance the positions
  // belong to: an HF part's rId, or a note's id.
  const regionRangeRects = (
    region: DisplayListHitRegion,
    partId: string,
    from: number,
    to: number
  ): DisplayListRect[] => {
    const live = handedOff();
    if (live !== undefined) {
      if (!live) return [];
      return region === 'header' || region === 'footer'
        ? live.hfRangeRects(region, partId, from, to)
        : region === 'footnote' || region === 'endnote'
          ? live.noteRangeRects(region, Number(partId), from, to)
          : [];
    }
    if (resident) {
      return parseQuery(
        residentQuery(
          () => resident.displayRangeRectsRegionJson(region, partId, from, to),
          'range_rects_region'
        ),
        [],
        'range_rects_region'
      );
    }
    // Probe capability first: invoking an absent by-handle export would trip
    // `runQuery`'s close-on-failure and drop the shared session handle,
    // degrading body queries too. Feature-detect and no-op instead.
    if (!cell.eng || !cell.eng.hasRangeRectsRegion?.()) return [];
    const raw = runQuery(
      cell.eng.rangeRectsRegionByHandle &&
        ((h: number) => cell.eng!.rangeRectsRegionByHandle!(h, region, partId, from, to)),
      () => cell.eng!.rangeRectsRegionJson!(getJson(), region, partId, from, to),
      'range_rects_region',
      allPages
    );
    return parseQuery(raw, [], 'range_rects_region');
  };

  const hfRangeRects = (
    region: 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): DisplayListRect[] => regionRangeRects(region, rId, from, to);

  const noteRangeRects = (
    region: 'footnote' | 'endnote',
    noteId: number,
    from: number,
    to: number
  ): DisplayListRect[] => regionRangeRects(region, String(noteId), from, to);

  /**
   * One caret per page from a scoped range query. An HF part paints on every
   * page carrying it, so the caller picks the edited page; a note paints on
   * exactly one, so the single answer is already the right one.
   */
  const scopedCaretRects = (
    scopedRangeRects: (from: number, to: number) => DisplayListRect[],
    pos: number
  ): DisplayListRect[] => {
    // The leading edge is the first slice on a page, the trailing edge the last.
    const caretsByPage = (rects: DisplayListRect[], leading: boolean) => {
      const byPage = new Map<number, DisplayListRect>();
      for (const rect of rects) {
        if (leading && byPage.has(rect.pageIndex)) continue;
        byPage.set(rect.pageIndex, {
          pageIndex: rect.pageIndex,
          x: leading ? rect.x : rect.x + rect.width,
          y: rect.y,
          width: 0,
          height: rect.height,
        });
      }
      return [...byPage.values()];
    };
    const forward = scopedRangeRects(pos, pos + 1);
    if (forward.length > 0) return caretsByPage(forward, true);
    if (pos > 0) {
      // end of line / end of doc: trailing edge of the previous position
      const backward = scopedRangeRects(pos - 1, pos);
      if (backward.length > 0) return caretsByPage(backward, false);
    }
    return [];
  };

  const hfCaretRects = (
    region: 'header' | 'footer',
    rId: string,
    pos: number
  ): DisplayListRect[] =>
    scopedCaretRects((from, to) => hfRangeRects(region, rId, from, to), pos);

  const noteCaretRects = (
    region: 'footnote' | 'endnote',
    noteId: number,
    pos: number
  ): DisplayListRect[] =>
    scopedCaretRects((from, to) => noteRangeRects(region, noteId, from, to), pos);

  // A position on a page whose content is not built yet resolves to the top of
  // that page's content box, which is enough to scroll it into view (and so
  // have it built). Unbuilt pages whose spans overlap, as a table row split
  // across them does, are picked in proportion to where the position falls in
  // their shared range; spans that only touch give it to the later page.
  const unbuiltPageRect = (pos: number): DisplayListRect | null => {
    const candidates = list.pages.filter((page) => {
      const span = page.unbuilt ? page.positionSpan : undefined;
      return span !== undefined && pos >= span[0] && pos <= span[1];
    });
    if (candidates.length === 0) return null;
    const low = Math.max(...candidates.map((page) => page.positionSpan![0]));
    const high = Math.min(...candidates.map((page) => page.positionSpan![1]));
    const share = high > low ? (pos - low) / (high - low + 1) : 1;
    const pick = Math.min(candidates.length - 1, Math.floor(share * candidates.length));
    const found = candidates[pick]!;
    return {
      pageIndex: found.pageIndex,
      x: found.contentBounds?.x ?? 0,
      y: found.contentBounds?.y ?? 0,
      width: 0,
      height: 0,
    };
  };

  const caretRect = (pos: number): DisplayListRect | null => {
    const live = handedOff();
    if (live !== undefined) return live?.caretRect(pos) ?? null;
    const forward = rangeRects(pos, pos + 1);
    if (forward.length > 0) {
      // left edge of the first covered slice is the caret
      const r = forward[0];
      return { pageIndex: r.pageIndex, x: r.x, y: r.y, width: 0, height: r.height };
    }
    // Before the trailing edge: at the start of an unbuilt page, the previous
    // position is still painted on the page before it.
    const unbuilt = unbuiltPageRect(pos);
    if (unbuilt) return unbuilt;
    if (pos > 0) {
      // end of doc / trailing edge: right edge of the previous position
      const backward = rangeRects(pos - 1, pos);
      if (backward.length > 0) {
        const r = backward[backward.length - 1];
        return { pageIndex: r.pageIndex, x: r.x + r.width, y: r.y, width: 0, height: r.height };
      }
    }
    return null;
  };

  const anchorRect = (pos: number): DisplayListRect | null => {
    const live = handedOff();
    if (live !== undefined) return live?.anchorRect(pos) ?? null;
    // [pos, pos+2) covers both "node position + first char at pos+1" and a
    // blank paragraph's zero-length marker at pos+1
    const forward = rangeRects(pos, pos + 2);
    if (forward.length > 0) return forward[0];
    return caretRect(pos);
  };

  const hfAnchorRects = (
    region: 'header' | 'footer',
    rId: string,
    pos: number
  ): DisplayListRect[] => {
    const forward = hfRangeRects(region, rId, pos, pos + 2);
    if (forward.length === 0) return hfCaretRects(region, rId, pos);
    const byPage = new Map<number, DisplayListRect>();
    for (const rect of forward) {
      if (!byPage.has(rect.pageIndex)) byPage.set(rect.pageIndex, rect);
    }
    return [...byPage.values()];
  };

  const pageRect = (pageIndex: number, rect: GeoRect): DisplayListRect => ({
    pageIndex,
    x: rect.x,
    y: rect.y,
    width: rect.w,
    height: rect.h,
  });

  const pageBounds = (pageIndex: number): DisplayListRect | null => {
    const page = list.pages[pageIndex];
    return page
      ? { pageIndex: page.pageIndex, x: 0, y: 0, width: page.width, height: page.height }
      : null;
  };

  const contentBounds = (pageIndex: number): DisplayListRect | null => {
    const page = list.pages[pageIndex];
    const bounds = page?.contentBounds;
    return page && bounds
      ? {
          pageIndex: page.pageIndex,
          x: bounds.x,
          y: bounds.y,
          width: bounds.width,
          height: bounds.height,
        }
      : null;
  };

  const columnBounds = (pageIndex: number): DisplayListRect[] => {
    const page = list.pages[pageIndex];
    if (!page) return [];
    return (page.columnBounds ?? []).map((bounds) => ({
      pageIndex: page.pageIndex,
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
    }));
  };

  const primitiveIdentity = (primitive: DisplayPrimitive): string | null => {
    if (primitive.paraId) return `para:${primitive.paraId}`;
    if (primitive.blockKey !== undefined) return `block-key:${primitive.blockKey}`;
    if (primitive.blockId !== undefined) return `block-id:${primitive.blockId}`;
    return null;
  };

  const publicBlockId = (primitive: DisplayPrimitive): number | string | undefined =>
    primitive.blockKey ?? primitive.blockId;

  let paragraphGroups: Map<string, DisplayListParagraphGeometry[]> | null = null;
  const getParagraphGroups = (): Map<string, DisplayListParagraphGeometry[]> => {
    if (paragraphGroups) return paragraphGroups;
    const accumulators = new Map<string, Map<number, DisplayListParagraphGeometry>>();
    for (const page of list.pages) {
      for (const primitive of page.primitives) {
        const identity = primitiveIdentity(primitive);
        if (!identity || primitive.docStart === undefined || primitive.docEnd === undefined) {
          continue;
        }
        const rect = displayPrimitiveRect(primitive);
        let byPage = accumulators.get(identity);
        if (!byPage) {
          byPage = new Map();
          accumulators.set(identity, byPage);
        }
        const current = byPage.get(page.pageIndex);
        if (!current) {
          byPage.set(page.pageIndex, {
            ...pageRect(page.pageIndex, rect),
            from: primitive.docStart,
            to: primitive.docEnd,
            blockId: publicBlockId(primitive),
            paraId: primitive.paraId,
          });
          continue;
        }
        const left = Math.min(current.x, rect.x);
        const top = Math.min(current.y, rect.y);
        const right = Math.max(current.x + current.width, rect.x + rect.w);
        const bottom = Math.max(current.y + current.height, rect.y + rect.h);
        current.x = left;
        current.y = top;
        current.width = right - left;
        current.height = bottom - top;
        current.from = Math.min(current.from, primitive.docStart);
        current.to = Math.max(current.to, primitive.docEnd);
      }
    }
    paragraphGroups = new Map(
      [...accumulators].map(([identity, byPage]) => [identity, [...byPage.values()]])
    );
    return paragraphGroups;
  };

  const paragraphRects = (pos: number): DisplayListParagraphGeometry[] => {
    let best: { identity: string; span: number; startsAtPos: boolean } | null = null;
    for (const page of list.pages) {
      for (const primitive of page.primitives) {
        const identity = primitiveIdentity(primitive);
        const from = primitive.docStart;
        const to = primitive.docEnd;
        if (!identity || from === undefined || to === undefined || pos < from || pos > to) continue;
        const candidate = { identity, span: Math.max(0, to - from), startsAtPos: from === pos };
        if (
          !best ||
          (candidate.startsAtPos && !best.startsAtPos) ||
          (candidate.startsAtPos === best.startsAtPos && candidate.span < best.span)
        ) {
          best = candidate;
        }
      }
    }
    return best ? (getParagraphGroups().get(best.identity) ?? []) : [];
  };

  const VISUAL_BASELINE_EPSILON = 1.5;
  const pageVisualLines: Array<readonly DisplayListVisualLine[] | undefined> = [];
  const visualLinesOnPage = (pageIndex: number): readonly DisplayListVisualLine[] => {
    const cached = pageVisualLines[pageIndex];
    if (cached) return cached;
    const page = list.pages[pageIndex];
    if (!page) return [];
    const pageLines: DisplayListVisualLine[] = [];
    // A line is found among the lines of its own identity, in the order they began.
    const linesByIdentity = new Map<string, DisplayListVisualLine[]>();
    let anonymous = 0;
    for (const primitive of page.primitives) {
      const baseline = visualLineBaseline(primitive);
      if (baseline === null || primitive.docStart === undefined || primitive.docEnd === undefined) {
        continue;
      }
      const identity = primitiveIdentity(primitive) ?? `anonymous:${anonymous++}`;
      const rect = displayPrimitiveRect(primitive);
      const sameIdentity = linesByIdentity.get(identity);
      const current = sameIdentity?.find(
        (line) => Math.abs(line.baseline - baseline) <= VISUAL_BASELINE_EPSILON
      );
      if (!current) {
        const line: DisplayListVisualLine = {
          ...pageRect(page.pageIndex, rect),
          baseline,
          from: primitive.docStart,
          to: primitive.docEnd,
          blockId: publicBlockId(primitive),
          paraId: primitive.paraId,
        };
        pageLines.push(line);
        if (sameIdentity) sameIdentity.push(line);
        else linesByIdentity.set(identity, [line]);
        continue;
      }
      const left = Math.min(current.x, rect.x);
      const top = Math.min(current.y, rect.y);
      const right = Math.max(current.x + current.width, rect.x + rect.w);
      const bottom = Math.max(current.y + current.height, rect.y + rect.h);
      current.x = left;
      current.y = top;
      current.width = right - left;
      current.height = bottom - top;
      current.from = Math.min(current.from, primitive.docStart);
      current.to = Math.max(current.to, primitive.docEnd);
    }
    pageVisualLines[pageIndex] = pageLines;
    return pageLines;
  };

  let visualLineCache: DisplayListVisualLine[] | null = null;
  const visualLines = (): readonly DisplayListVisualLine[] => {
    visualLineCache ??= list.pages.flatMap((_, pageIndex) => visualLinesOnPage(pageIndex));
    return visualLineCache;
  };

  const visualLineAtPosition = (pos: number): DisplayListVisualLine | null => {
    let best: DisplayListVisualLine | null = null;
    for (const line of visualLines()) {
      if (pos < line.from || pos > line.to) continue;
      if (!best || line.to - line.from < best.to - best.from) best = line;
    }
    return best;
  };

  const imageGeometry = (
    located: LocatedImagePrimitive | null
  ): DisplayListImageGeometry | null => {
    if (!located) return null;
    const pos = located.primitive.docStart;
    if (pos === undefined) return null;
    const { primitive } = located;
    return {
      ...located,
      pos,
      rect: {
        pageIndex: located.pageIndex,
        x: primitive.x,
        y: primitive.y,
        width: primitive.w,
        height: primitive.h,
      },
    };
  };

  const imageAtPoint = (
    pageIndex: number,
    x: number,
    y: number,
    region: DisplayListImageRegion = 'body',
    rId?: string
  ): DisplayListImageGeometry | null =>
    imageGeometry(findImagePrimitiveAtPoint(list, pageIndex, x, y, region, rId));

  const imageByPos = (
    pos: number,
    region: DisplayListImageRegion = 'body',
    rId?: string
  ): DisplayListImageGeometry | null =>
    imageGeometry(findImagePrimitiveByDocPos(list, pos, region, rId));

  const dispose = (): void => {
    if (disposed) return;
    disposed = true;
    closeHandle();
    handleFinalizers?.unregister(finalizerToken);
  };

  const queries: DisplayListQueries = {
    get displayList() {
      return handedOff()?.displayList ?? list;
    },
    isReady: () => (resident !== null || cell.eng !== null) && sourceError === null,
    sourceState: () =>
      sourceError
        ? { status: 'error', error: sourceError }
        : resident || cell.eng
          ? { status: 'ready' }
          : { status: 'loading' },
    whenReady: () => readyPromise,
    pageCount: () => {
      const live = handedOff();
      return live ? live.pageCount() : list.pages.length;
    },
    pageSize: (pageIndex: number) => {
      const live = handedOff();
      if (live) return live.pageSize(pageIndex);
      const page = list.pages[pageIndex];
      return page ? { width: page.width, height: page.height } : null;
    },
    pageBounds: (pageIndex: number) => {
      const live = handedOff();
      return live ? live.pageBounds(pageIndex) : pageBounds(pageIndex);
    },
    contentBounds: viaLive(contentBounds, (live, pageIndex) => live.contentBounds(pageIndex)),
    columnBounds: viaLive(columnBounds, (live, pageIndex) => live.columnBounds(pageIndex)),
    paragraphRects: viaLive(paragraphRects, (live, pos) => live.paragraphRects(pos)),
    visualLines: viaLive(visualLines, (live) => live.visualLines()),
    visualLinesOnPage: (pageIndex: number) => {
      const live = handedOff();
      return live ? live.visualLinesOnPage(pageIndex) : visualLinesOnPage(pageIndex);
    },
    visualLineExtent: viaLive(
      (pageIndex: number) => {
        const page = list.pages[pageIndex];
        return page ? visualLineExtent(page) : null;
      },
      (live, pageIndex) => live.visualLineExtent(pageIndex)
    ),
    visualLineAtPosition: viaLive(visualLineAtPosition, (live, pos) =>
      live.visualLineAtPosition(pos)
    ),
    imageAtPoint: viaLive(imageAtPoint, (live, ...args) => live.imageAtPoint(...args)),
    imageByPos: viaLive(imageByPos, (live, ...args) => live.imageByPos(...args)),
    hitTestRegions,
    verticalMove,
    rangeRects,
    hfRangeRects,
    noteRangeRects,
    hfCaretRects,
    noteCaretRects,
    hfAnchorRects,
    caretRect,
    anchorRect,
    sidebarAnchorRect: anchorRect,
    prime: openHandle,
    dispose,
  };

  // Auto-release the handle if the facade is dropped without dispose(). The held
  // value closes over `cell` only, never over `queries` or its donor, so
  // registering cannot keep `queries` alive.
  handleFinalizers?.register(queries, closeHandle, finalizerToken);

  facadeDeltaSeeds.set(queries, {
    list,
    storeRevisions: () => storeRevisions,
    engine: () => cell.eng,
    hasHandle: () => cell.handle !== null,
    donor: () => donorFacade,
    supersede: () => {
      superseded = true;
    },
    lineage,
    line: line ?? null,
    disposed: () => disposed,
    takeHandle: () => {
      const transferred = cell.handle;
      if (transferred !== null) {
        // ownership moves to the adopting facade: neither dispose() nor the
        // finalizer may close it here anymore
        cell.handle = null;
        handleFinalizers?.unregister(finalizerToken);
      }
      return transferred;
    },
  });
  lineage.newest = weakly(queries);

  return queries;
}

function isResidentQueryEngine(
  engine: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine | undefined
): engine is ResidentDisplayListQueryEngine {
  return (
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayHitTestRegionsJson ===
      'function' &&
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayVerticalMoveJson ===
      'function' &&
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayRangeRectsJson ===
      'function' &&
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayRangeRectsRegionJson ===
      'function'
  );
}
