/**
 * Synchronous query facade over a built DisplayList, backed by the Rust
 * hit-testing module (`crates/docx-layout/src/hit.rs`).
 *
 * This is the sole pointer/selection geometry source for the canvas
 * renderer (it replaced the deleted painted-DOM resolvers): geometry comes
 * from the immutable display list, never from DOM rects.
 *
 * Perf model: when the embedded wasm carries the session-handle exports, this
 * facade parses the display list into the Rust store exactly ONCE
 * (`openDisplayList` → handle) and routes every hit-test / range-rect query
 * through the by-handle exports — no per-query re-serialization or re-parse.
 * When those exports are absent (before the integrator re-embeds the wasm) it
 * falls back to the JSON-arg exports, stringifying the list once and reusing
 * the string for every query. Either way the query RESULTS are byte-identical —
 * the handle path is pure perf.
 *
 * Handle lifecycle: the handle is opened once per facade (per display-list
 * build) and released by `dispose()`. Callers that drop the facade without
 * disposing (e.g. a React `useMemo` replacement) are covered by a
 * `FinalizationRegistry` that closes the handle when the facade is collected;
 * the Rust store additionally caps live handles and evicts the oldest, so a
 * missed finalize can never grow memory unbounded.
 *
 * Queries are synchronous; until the lazily-imported wasm module resolves they
 * return `null`/`[]`. In practice the module is already loaded by the time a
 * queries instance exists, because building the display list itself went
 * through the same module (`loadRustDisplayListQueryEngine` shares the promise
 * with `buildRustDisplayList`).
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { DisplayList, DisplayPrimitive } from './displayList';
import { displayPrimitiveRect, type GeoRect } from './displayListGeometry';
import {
  findImagePrimitiveAtPoint,
  findImagePrimitiveByDocPos,
  type LocatedImagePrimitive,
} from './displayListImages';
import { loadRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from './rustDisplayList';

/** Query surface implemented by the editing wasm over its resident list. */
export interface ResidentDisplayListQueryEngine {
  displayHitTestRegionsJson(pageIndex: number, x: number, y: number): string;
  displayRangeRectsJson(from: number, to: number): string;
  displayRangeRectsRegionJson(
    region: 'body' | 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): string;
}

/** which part of a page owns a hit — mirrors `HitRegion` in hit.rs */
export type DisplayListHitRegion = 'body' | 'header' | 'footer';

/**
 * Region-aware hit result. For `header`/`footer` the position refers to the
 * HF ProseMirror doc identified by `rId`, NOT the body doc — the caller must
 * route the selection to that HF editor, exactly like the DOM path scopes
 * clicks to `.layout-page-header|footer`.
 */
export interface DisplayListRegionHit {
  region: DisplayListHitRegion;
  rId?: string;
  pos: number | null;
}

/** one highlight rectangle of a PM range, page-local px */
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
    region?: DisplayListHitRegion,
    rId?: string
  ): DisplayListImageGeometry | null;
  /** Image whose atom starts at `pos`. Body by default. */
  imageByPos(
    pos: number,
    region?: DisplayListHitRegion,
    rId?: string
  ): DisplayListImageGeometry | null;
  /** region-aware point → doc position (page-local coordinates) */
  hitTestRegions(pageIndex: number, x: number, y: number): DisplayListRegionHit | null;
  /** body PM range → highlight rects */
  rangeRects(from: number, to: number): DisplayListRect[];
  /**
   * Header/footer PM range → highlight rects for the region's band. `region` is
   * `'header' | 'footer'`; `rId` identifies the HF ProseMirror doc, and
   * `from`/`to` are positions in THAT doc. The same HF doc paints on every page
   * carrying the part, so this returns one rect-set per such page (each tagged
   * with its `pageIndex`) — the caller picks the page it is editing. Returns
   * `[]` until the region-aware exports are embedded (feature-detected).
   */
  hfRangeRects(
    region: 'header' | 'footer',
    rId: string,
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
 * Build a query facade for one display list. The optional `engine` makes the
 * facade synchronous and deterministic in tests; without it the shared wasm
 * module is loaded lazily and queries no-op (`null`/`[]`) until it resolves.
 */
export function createDisplayListQueries(
  list: DisplayList,
  engine?: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine
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
  let disposed = false;

  const closeHandle = (): void => {
    if (handle !== null) {
      try {
        eng?.closeDisplayList?.(handle);
      } catch {
        // a close failure must never surface — the store caps handles anyway
      }
      handle = null;
    }
  };

  // open exactly one handle once the engine is known to support sessions; a
  // failure leaves `handle` null so queries take the JSON-arg path
  const openHandle = (): void => {
    if (disposed || handle !== null || !eng) return;
    if (!eng.hasDisplayListSession?.() || !eng.openDisplayList) return;
    try {
      handle = eng.openDisplayList(getJson());
    } catch (error) {
      handle = null;
      console.warn(
        '[CanvasRenderer] display-list session open failed; using JSON-arg queries',
        error
      );
    }
  };

  if (resident) {
    resolveReady();
  } else if (eng) {
    openHandle();
    resolveReady();
  } else {
    loadRustDisplayListQueryEngine().then(
      (loaded) => {
        eng = loaded;
        openHandle();
        resolveReady();
        // if disposed while the engine was loading, release the just-opened handle
        if (disposed) closeHandle();
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
    if (!eng) return null;
    if (handle !== null && byHandle) {
      try {
        return byHandle(handle);
      } catch (error) {
        console.warn(`[CanvasRenderer] ${label} session query failed; falling back`, error);
        closeHandle();
      }
    }
    try {
      return byJson();
    } catch (error) {
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
    try {
      return query();
    } catch (error) {
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

  const hfRangeRects = (
    region: 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): DisplayListRect[] => {
    if (resident) {
      return parseQuery(
        residentQuery(
          () => resident.displayRangeRectsRegionJson(region, rId, from, to),
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
        ((h: number) => eng!.rangeRectsRegionByHandle!(h, region, rId, from, to)),
      () => eng!.rangeRectsRegionJson!(getJson(), region, rId, from, to),
      'range_rects_region'
    );
    return parseQuery(raw, [], 'range_rects_region');
  };

  const hfCaretRects = (
    region: 'header' | 'footer',
    rId: string,
    pos: number
  ): DisplayListRect[] => {
    // The same HF doc paints on every page carrying the part, so a caret query
    // returns one candidate per page; the caller renders the edited one.
    const forward = hfRangeRects(region, rId, pos, pos + 1);
    if (forward.length > 0) {
      // left edge of the first slice on each page is the caret there
      const byPage = new Map<number, DisplayListRect>();
      for (const r of forward) {
        if (!byPage.has(r.pageIndex)) {
          byPage.set(r.pageIndex, {
            pageIndex: r.pageIndex,
            x: r.x,
            y: r.y,
            width: 0,
            height: r.height,
          });
        }
      }
      return [...byPage.values()];
    }
    if (pos > 0) {
      // end of line / end of doc: trailing edge of the previous position
      const backward = hfRangeRects(region, rId, pos - 1, pos);
      if (backward.length > 0) {
        const byPage = new Map<number, DisplayListRect>();
        for (const r of backward) {
          byPage.set(r.pageIndex, {
            pageIndex: r.pageIndex,
            x: r.x + r.width,
            y: r.y,
            width: 0,
            height: r.height,
          });
        }
        return [...byPage.values()];
      }
    }
    return [];
  };

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
    region: DisplayListHitRegion = 'body',
    rId?: string
  ): DisplayListImageGeometry | null =>
    imageGeometry(findImagePrimitiveAtPoint(list, pageIndex, x, y, region, rId));

  const imageByPos = (
    pos: number,
    region: DisplayListHitRegion = 'body',
    rId?: string
  ): DisplayListImageGeometry | null =>
    imageGeometry(findImagePrimitiveByDocPos(list, pos, region, rId));

  // lets dispose() cancel the finalizer below so a handle is never double-closed
  const finalizerToken = {};
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
    rangeRects,
    hfRangeRects,
    hfCaretRects,
    hfAnchorRects,
    caretRect,
    anchorRect,
    sidebarAnchorRect: anchorRect,
    dispose,
  };

  // Auto-release the handle if the facade is dropped without dispose(). The held
  // value is `closeHandle` (a thunk over `handle`/`eng`, never over `queries`),
  // so registering cannot keep `queries` alive.
  handleFinalizers?.register(queries, closeHandle, finalizerToken);

  return queries;
}

function isResidentQueryEngine(
  engine: RustDisplayListQueryEngine | ResidentDisplayListQueryEngine | undefined
): engine is ResidentDisplayListQueryEngine {
  return (
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayHitTestRegionsJson ===
      'function' &&
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayRangeRectsJson ===
      'function' &&
    typeof (engine as ResidentDisplayListQueryEngine | undefined)?.displayRangeRectsRegionJson ===
      'function'
  );
}
