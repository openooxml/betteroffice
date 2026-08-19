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

import type { DisplayList, DisplayPrimitive } from './displayList';
import { displayPageRevision } from './frameDelta';
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

/**
 * Internal handoff state for handle adoption between consecutive facades.
 * Keyed weakly so a dropped facade can never leak its list.
 */
interface FacadeDeltaSeed {
  list: DisplayList;
  /** Per-page in-place mutation revisions captured at facade creation. The
   * owned frame-delta path patches positions through the SAME page objects,
   * so identity alone cannot prove a page still matches the parsed store. */
  pageRevisions: number[];
  engine(): RustDisplayListQueryEngine | null;
  /** Relinquish the live handle (the donor's queries fall back to JSON-arg). */
  takeHandle(): number | null;
  hasHandle(): boolean;
  /** Nearest ancestor facade that held the handle when this one was created. */
  donor(): DisplayListQueries | null;
  /** A successor now owns this generation: never open/adopt a handle here. */
  supersede(): void;
}

const facadeDeltaSeeds = new WeakMap<DisplayListQueries, FacadeDeltaSeed>();

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

/**
 * Page-delta between two display lists, exploiting the retained-frame
 * invariant that unchanged pages keep object identity across builds. A page
 * is reusable only when both its identity and its in-place mutation revision
 * are unchanged since the donor parsed it. Returns null when nothing is
 * reusable (a full open costs the same).
 */
function buildDisplayListUpdateJson(seed: FacadeDeltaSeed, next: DisplayList): string | null {
  const previousIndex = new Map<unknown, number>();
  seed.list.pages.forEach((page, index) => previousIndex.set(page, index));
  const reuse: Array<[number, number]> = [];
  const replace: Array<[number, unknown]> = [];
  next.pages.forEach((page, index) => {
    const from = previousIndex.get(page);
    if (from !== undefined && displayPageRevision(page) === seed.pageRevisions[from]) {
      reuse.push([index, from]);
      previousIndex.delete(page);
    } else {
      replace.push([index, page]);
    }
  });
  if (reuse.length === 0) return null;
  return JSON.stringify({
    total: next.pages.length,
    ...(next.contractVersion !== undefined ? { contractVersion: next.contractVersion } : {}),
    reuse,
    replace,
  });
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
 */
export function createDisplayListQueries(
  list: DisplayList,
  engine?: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine,
  previous?: DisplayListQueries | null
): DisplayListQueries {
  let json: string | null = null;
  const getJson = (): string => (json ??= JSON.stringify(list));

  const resident: ResidentDisplayListQueryEngine | null = isResidentQueryEngine(engine)
    ? engine
    : null;
  let eng: RustDisplayListQueryEngine | null = resident
    ? null
    : (engine as RustDisplayListQueryEngine | undefined) ?? null;
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
  // `handle`; null means the JSON-arg fallback (unsupported wasm, open failed, or
  // a handle dropped after a stale-handle error).
  let handle: number | null = null;
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
  if (previous) {
    const previousSeed = facadeDeltaSeeds.get(previous);
    if (previousSeed) {
      donorFacade = previousSeed.hasHandle() ? previous : previousSeed.donor();
      previousSeed.supersede();
    }
  }

  const source = (): DisplayListQuerySource | null => resident ?? eng;

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
    handle = null;
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

  const closeHandle = (): void => {
    if (handle !== null && isDead()) {
      handle = null;
      return;
    }
    if (handle !== null) {
      try {
        eng?.closeDisplayList?.(handle);
      } catch {
        // a close failure must never surface — the store caps handles anyway
      }
      handle = null;
    }
  };

  // adopt the donor facade's parsed list when only some pages changed:
  // ships a page-delta into the Rust store instead of the whole list
  const adoptHandle = (): boolean => {
    const donor = donorFacade;
    donorFacade = null;
    if (!donor || !eng?.updateDisplayList || !eng.hasDisplayListUpdate?.()) return false;
    const seed = facadeDeltaSeeds.get(donor);
    if (!seed || seed.engine() !== eng) return false;
    const update = buildDisplayListUpdateJson(seed, list);
    if (!update) return false;
    const adopted = seed.takeHandle();
    if (adopted === null) return false;
    try {
      eng.updateDisplayList(adopted, update);
      handle = adopted;
      return true;
    } catch (error) {
      // the Rust side closes the handle on a failed update; close defensively
      // anyway (idempotent) in case the failure happened before wasm ran, then
      // fall through to a fresh full open
      try {
        eng.closeDisplayList?.(adopted);
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
    if (disposed || superseded || handleAttempted || handle !== null || !eng || isDead()) return;
    if (!eng.hasDisplayListSession?.() || !eng.openDisplayList) return;
    handleAttempted = true;
    if (adoptHandle()) return;
    try {
      handle = eng.openDisplayList(getJson());
    } catch (error) {
      handle = null;
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

  if (resident || eng) {
    resolveReady();
  } else {
    loadRustDisplayListQueryEngine().then(
      (loaded) => {
        eng = loaded;
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
    label: string
  ): string | null => {
    if (!eng || isDead()) return null;
    if (handle === null) openHandle();
    if (handle !== null && byHandle) {
      try {
        return byHandle(handle);
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
      eng?.hitTestRegionsByHandle &&
        ((h: number) => eng!.hitTestRegionsByHandle!(h, pageIndex, x, y)),
      () => eng!.hitTestRegionsJson(getJson(), pageIndex, x, y),
      'hit_test_regions'
    );
    return parseQuery(raw, null, 'hit_test_regions');
  };

  const rangeRects = (from: number, to: number): DisplayListRect[] => {
    if (resident) {
      return parseQuery(
        residentQuery(() => resident.displayRangeRectsJson(from, to), 'range_rects'),
        [],
        'range_rects'
      );
    }
    const raw = runQuery(
      eng?.rangeRectsByHandle && ((h: number) => eng!.rangeRectsByHandle!(h, from, to)),
      () => eng!.rangeRectsJson(getJson(), from, to),
      'range_rects'
    );
    return parseQuery(raw, [], 'range_rects');
  };

  const verticalMove = (
    position: number,
    direction: 'up' | 'down',
    goalX?: number
  ): DisplayListVerticalMove | null => {
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
    if (!eng?.verticalMoveJson) return null;
    const raw = runQuery(
      eng.verticalMoveByHandle &&
        ((h: number) => eng!.verticalMoveByHandle!(h, position, direction, resolvedGoalX)),
      () => eng!.verticalMoveJson!(getJson(), position, direction, resolvedGoalX),
      'vertical_move'
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
    if (!eng || !eng.hasRangeRectsRegion?.()) return [];
    const raw = runQuery(
      eng.rangeRectsRegionByHandle &&
        ((h: number) => eng!.rangeRectsRegionByHandle!(h, region, partId, from, to)),
      () => eng!.rangeRectsRegionJson!(getJson(), region, partId, from, to),
      'range_rects_region'
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

  const caretRect = (pos: number): DisplayListRect | null => {
    const forward = rangeRects(pos, pos + 1);
    if (forward.length > 0) {
      // left edge of the first covered slice is the caret
      const r = forward[0];
      return { pageIndex: r.pageIndex, x: r.x, y: r.y, width: 0, height: r.height };
    }
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
  let visualLineCache: DisplayListVisualLine[] | null = null;
  const visualLines = (): readonly DisplayListVisualLine[] => {
    if (visualLineCache) return visualLineCache;
    const lines: DisplayListVisualLine[] = [];
    for (const page of list.pages) {
      const pageLines: Array<DisplayListVisualLine & { identity: string }> = [];
      let anonymous = 0;
      for (const primitive of page.primitives) {
        if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
        if (primitive.docStart === undefined || primitive.docEnd === undefined) continue;
        if (primitive.kind === 'glyphRun' && primitive.glyphs.length === 0) continue;
        const baseline =
          primitive.kind === 'text'
            ? primitive.baselineY
            : primitive.glyphs.reduce((max, glyph) => Math.max(max, glyph.y), -Infinity);
        if (!Number.isFinite(baseline)) continue;
        const identity = primitiveIdentity(primitive) ?? `anonymous:${anonymous++}`;
        const rect = displayPrimitiveRect(primitive);
        const current = pageLines.find(
          (line) =>
            line.identity === identity &&
            Math.abs(line.baseline - baseline) <= VISUAL_BASELINE_EPSILON
        );
        if (!current) {
          pageLines.push({
            identity,
            ...pageRect(page.pageIndex, rect),
            baseline,
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
      lines.push(...pageLines.map(({ identity: _identity, ...line }) => line));
    }
    visualLineCache = lines;
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
    displayList: list,
    isReady: () => (resident !== null || eng !== null) && sourceError === null,
    sourceState: () =>
      sourceError
        ? { status: 'error', error: sourceError }
        : resident || eng
          ? { status: 'ready' }
          : { status: 'loading' },
    whenReady: () => readyPromise,
    pageCount: () => list.pages.length,
    pageSize: (pageIndex: number) => {
      const page = list.pages[pageIndex];
      return page ? { width: page.width, height: page.height } : null;
    },
    pageBounds,
    contentBounds,
    columnBounds,
    paragraphRects,
    visualLines,
    visualLineAtPosition,
    imageAtPoint,
    imageByPos,
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
  // value is `closeHandle` (a thunk over `handle`/`eng`, never over `queries`),
  // so registering cannot keep `queries` alive.
  handleFinalizers?.register(queries, closeHandle, finalizerToken);

  facadeDeltaSeeds.set(queries, {
    list,
    pageRevisions: list.pages.map(displayPageRevision),
    engine: () => eng,
    hasHandle: () => handle !== null,
    donor: () => donorFacade,
    supersede: () => {
      superseded = true;
    },
    takeHandle: () => {
      const transferred = handle;
      if (transferred !== null) {
        // ownership moves to the adopting facade: neither dispose() nor the
        // finalizer may close it here anymore
        handle = null;
        handleFinalizers?.unregister(finalizerToken);
      }
      return transferred;
    },
  });

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
