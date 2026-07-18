/**
 * JS seam to the Rust display-list builder (crates/docx-layout).
 *
 * Marshals the same `{ measured, options, layout }` triple the golden
 * fixtures pin (see scripts/export-displaylist-fixtures.ts) across the wasm
 * boundary and parses the resulting DisplayList. All paint/geometry decisions
 * happen in Rust — this module is serialization glue only.
 *
 * The wasm module is loaded lazily via dynamic import on first build, so the
 * ~800KB inlined binary costs nothing until the experimental canvas renderer
 * is actually enabled.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { Layout } from '../pagination/types';
import type { MeasuredBlock } from '../pagination/measuredBlock';
import type { DisplayList } from './displayList';
import { applyFrameDelta, decodeFrameDelta, type RetainedFrame } from './frameDelta';

/**
 * One header/footer part in the `headersFooters` envelope field: the measured
 * blocks of the HF ProseMirror doc identified by `rId`, plus the band metrics
 * `convertHeaderFooterPmDocToContent` computed (passed through verbatim — the
 * Rust builder re-derives nothing).
 */
export interface DisplayListHfVariant {
  rId: string;
  kind: 'header' | 'footer';
  type: 'default' | 'first' | 'even';
  /** `toMeasuredBlocks(content.blocks, content.measures)` — same schema as body `measured` */
  measured: MeasuredBlock[];
  /** HeaderFooterContent.height (in-flow stack incl. floating blocks) */
  height: number;
  /** HeaderFooterContent.flowHeight (in-flow band height; fallback: height) */
  flowHeight?: number;
  visualTop?: number;
  visualBottom?: number;
  /**
   * Per-page resolved widths of PAGE/NUMPAGES field runs in this HF part, so a
   * centered/right-aligned line carrying such a field re-centers per page. The
   * line width is measured ONCE at the field's fallback text ("1"), so without
   * this the builder holds the same centered position on every page even though
   * "Page 2 of 3"/"Page 10 of 12" have different widths. Omitted ⇒ the fields
   * ride the char-distributed fallback width (byte-identical to before).
   */
  fieldWidths?: DisplayListFieldWidths[];
}

/**
 * Per-page widths of one PAGE/NUMPAGES field run in an HF part. `pmStart` is the
 * field run's position in the HF ProseMirror doc (the key the Rust builder
 * matches). `fallbackWidth` is the width the measure baked into `line.width`
 * (the field's fallback text); `perPage[i]` is the width of the field's resolved
 * text on layout page index `i` (PAGE → that page's number, NUMPAGES → total).
 */
export interface DisplayListFieldWidths {
  pmStart: number;
  fallbackWidth: number;
  perPage: number[];
}

/**
 * Optional header/footer payload of the builder envelope. Absent ⇒ the Rust
 * output is byte-identical to the body-only build. Flags come off the section
 * properties / settings the DOM painter uses; distances are the same px the
 * painter resolves (`headerDistance ?? page.margins.header ?? 48`).
 */
export interface DisplayListHeadersFooters {
  titlePg?: boolean;
  evenAndOddHeaders?: boolean;
  headerDistance?: number;
  footerDistance?: number;
  /** Page-level Word watermark, painted behind body/header/footer content. */
  watermark?: DisplayListWatermark;
  variants: DisplayListHfVariant[];
  /** Stable section identity. Undefined = document-global legacy envelope. */
  sectionId?: string;
  /** Zero-based section index. */
  sectionIndex?: number;
}

export type DisplayListWatermark = DisplayListTextWatermark | DisplayListPictureWatermark;

export interface DisplayListTextWatermark {
  kind: 'text';
  text: string;
  font: string;
  color: string;
  semitransparent: boolean;
  layout: 'diagonal' | 'horizontal';
  fontSize?: number;
  /** Undefined defaults to true for watermarks. */
  decorative?: boolean;
}

export interface DisplayListPictureWatermark {
  kind: 'picture';
  relId?: string;
  dataUrl?: string;
  scale: number;
  washout: boolean;
  widthEmu?: number;
  heightEmu?: number;
  /** Undefined defaults to true for watermarks. */
  decorative?: boolean;
}

/** One measured note in a page/section note area. */
export interface DisplayListNoteItem {
  kind?: 'footnote' | 'endnote';
  id?: number;
  displayLabel?: string;
  measured?: MeasuredBlock[];
  height?: number;
  anchorDocStart?: number;
  anchorDocEnd?: number;
  customMarkFollows?: boolean;
}

/** Typed non-fragmented note-region input for the Rust display builder. */
export interface DisplayListNoteArea {
  pageIndex?: number;
  sectionId?: string;
  kind?: 'footnote' | 'endnote';
  placement?: 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';
  y?: number;
  height?: number;
  columns?: number;
  separator?: DisplayListNoteItem;
  notes?: DisplayListNoteItem[];
}

/** Stable reviewer palette entry for comments/revisions. */
export interface DisplayListCommentAuthor {
  id?: string;
  name?: string;
  paletteIndex?: number;
  color?: string;
}

/** One reply of a comment thread (a11y announcement input). */
export interface DisplayListCommentReply {
  id?: number;
  authorId?: string;
  authorName?: string;
  date?: string;
  /** Plain-text body (the builder caps it before it reaches primitives). */
  text?: string;
}

/**
 * Per-comment thread metadata keyed by comment id, joined by the builder onto
 * every primitive carrying that id in `commentIds` so the a11y mirror can
 * announce the complete thread (author name, date, body, replies). All string
 * values are file-derived/attacker-controlled; consumers must only ever
 * assign them via textContent/setAttribute.
 */
export interface DisplayListCommentThread {
  /** Comment id — matches the numeric ids in `resolvedCommentIds`/`commentIds`. */
  id?: number;
  authorId?: string;
  authorName?: string;
  date?: string;
  /** Plain-text body (builder-capped before emission). */
  text?: string;
  /** Replies in thread order (builder-capped count and text). */
  replies?: DisplayListCommentReply[];
}

/** The `{ measured, options, layout }` envelope the Rust builder consumes. */
export interface DisplayListBuildInputs {
  /** Serialization contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  measured: MeasuredBlock[];
  options: unknown;
  layout: Layout;
  /** header/footer parts + section flags; omitted ⇒ body-only display list */
  headersFooters?: DisplayListHeadersFooters;
  /**
   * Doc-wide font-id chains keyed `"<family lowercase>|<bold 0|1>|<italic 0|1>"`
   * → measurement `FontStore` ids — the SAME map the measurement input carries.
   * Present only under Rust measurement (the ids must belong to a populated
   * store); its presence against that store is what gates GlyphRun emission.
   * Omitted ⇒ every text run takes the browser-measured `TextRunPrimitive` path
   * (byte-identical to the pre-glyph build).
   */
  fontChains?: Record<string, number[]>;
  /** Page/section note regions. Undefined = no note emission. */
  noteAreas?: DisplayListNoteArea[];
  /** Resolved comments suppressed from active tint. Undefined = none known. */
  resolvedCommentIds?: number[];
  /** Stable reviewer palette. Undefined = legacy hard-coded colors. */
  commentAuthors?: DisplayListCommentAuthor[];
  /** Per-comment thread metadata for a11y announcements. Undefined = none. */
  commentThreads?: DisplayListCommentThread[];
}

/** Minimal engine surface, injectable so tests can fake the wasm module. */
export interface RustDisplayListEngine {
  buildDisplayListJson(input: string): string;
  /** Resident editing engines expose binary deltas; the stateless layout wasm does not. */
  buildDisplayListFrame?(input: string, expectedFrameEpoch: number): Uint8Array;
  /** One-owner ordinary input path; available on the resident editing engine. */
  applyInput?(text: string, expectedFrameEpoch: number): Uint8Array;
  displayHitTestRegionsJson?(pageIndex: number, x: number, y: number): string;
  displayRangeRectsJson?(from: number, to: number): string;
  displayRangeRectsRegionJson?(
    region: 'body' | 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): string;
}

export type RustDisplayListSourceErrorStage = 'load' | 'build' | 'parse' | 'decode' | 'apply';

export interface RustDisplayFrameResult {
  displayList: DisplayList;
  /** Null only on the compatibility JSON engine path. */
  frame: RetainedFrame | null;
  transport: 'frame-delta-v1' | 'json';
}

/**
 * Typed terminal error from the mandatory Rust display-list source. Hosts may
 * surface this through their editor error boundary; it is never permission to
 * silently substitute painter geometry.
 */
export class RustDisplayListSourceError extends Error {
  readonly stage: RustDisplayListSourceErrorStage;

  constructor(stage: RustDisplayListSourceErrorStage, cause: unknown) {
    const detail = cause instanceof Error ? cause.message : String(cause);
    super(`Rust display-list ${stage} failed: ${detail}`);
    this.name = 'RustDisplayListSourceError';
    this.stage = stage;
  }
}

/**
 * Query surface over an already-built display list, injectable so tests can
 * fake the wasm module.
 *
 * The `*Json` calls take the display-list JSON string as an argument (the Rust
 * side re-parses it per query — callers should cache the string). The
 * session-handle members parse the list ONCE (`openDisplayList` → handle) and
 * answer many queries by handle with no re-serialization, reusing the same
 * hit/range logic so results are byte-identical. The handle members are
 * OPTIONAL: the embedded wasm carries them only after the integrator re-embeds,
 * so `createDisplayListQueries` feature-detects `hasDisplayListSession` and
 * falls back to the `*Json` path when they are absent.
 */
export interface RustDisplayListQueryEngine {
  /** region-aware hit test → `{"region","rId"?,"pos"}` or `"null"` JSON */
  hitTestRegionsJson(displayList: string, pageIndex: number, x: number, y: number): string;
  /** body PM range → JSON array of `{pageIndex,x,y,width,height}` rects */
  rangeRectsJson(displayList: string, from: number, to: number): string;
  /**
   * Region-aware PM range → JSON array of rects. `region` is
   * `'body' | 'header' | 'footer'`; `rId` scopes header/footer to one HF part
   * (empty for body / match-any). OPTIONAL: the embedded wasm carries it only
   * after the integrator re-embeds, so `createDisplayListQueries` falls back to
   * `[]` when absent (HF selection geometry stays a documented gap until then).
   */
  rangeRectsRegionJson?(
    displayList: string,
    region: string,
    rId: string,
    from: number,
    to: number
  ): string;
  /** true when the embedded wasm carries the region-aware range-rect exports */
  hasRangeRectsRegion?(): boolean;
  /** true when the embedded wasm carries the session-handle exports below */
  hasDisplayListSession?(): boolean;
  /** parse a display list once, return a handle for the by-handle queries */
  openDisplayList?(displayList: string): number;
  /** free a handle's parsed display list */
  closeDisplayList?(handle: number): void;
  /** region-aware hit test against a stored display list (by handle) */
  hitTestRegionsByHandle?(handle: number, pageIndex: number, x: number, y: number): string;
  /** body PM range against a stored display list (by handle) */
  rangeRectsByHandle?(handle: number, from: number, to: number): string;
  /** region-aware PM range against a stored display list (by handle) */
  rangeRectsRegionByHandle?(
    handle: number,
    region: string,
    rId: string,
    from: number,
    to: number
  ): string;
}

let enginePromise: Promise<RustDisplayListEngine & RustDisplayListQueryEngine> | null = null;

function loadEngine(): Promise<RustDisplayListEngine & RustDisplayListQueryEngine> {
  enginePromise ??= import('../wasm/index').then(async (m) => {
    await m.preloadLayoutWasm();
    return {
    buildDisplayListJson: m.buildDisplayListJson,
    hitTestRegionsJson: m.hitTestRegionsJson,
    rangeRectsJson: m.rangeRectsJson,
    rangeRectsRegionJson: m.rangeRectsRegionJson,
    hasRangeRectsRegion: m.hasRangeRectsRegion,
    hasDisplayListSession: m.hasDisplayListSession,
    openDisplayList: m.openDisplayList,
    closeDisplayList: m.closeDisplayList,
    hitTestRegionsByHandle: m.hitTestRegionsByHandle,
    rangeRectsByHandle: m.rangeRectsByHandle,
    rangeRectsRegionByHandle: m.rangeRectsRegionByHandle,
  };
  });
  return enginePromise;
}

/**
 * Lazily load the wasm query surface (hit-test + range-rects). Shares the
 * module promise with `buildRustDisplayList`, so on the canvas-renderer path
 * the engine is already resolved by the time the first display list lands.
 */
export function loadRustDisplayListQueryEngine(): Promise<RustDisplayListQueryEngine> {
  return loadEngine();
}

/**
 * Serialize the `{ measured, options, layout, headersFooters }` envelope to the
 * JSON string the wasm builder consumes, caching the serialized `measure`
 * sub-object of each MeasuredBlock across passes.
 *
 * Why this is a per-keystroke win: on every layout pass the paginator rebuilds
 * the block objects fresh from the resident engine, so the `block`
 * side of each MeasuredBlock has no reference stability and re-serializes every
 * pass. The `measure` side does NOT — the paragraph measure cache is keyed by a
 * content hash (see `layout/measure/cache.ts`), so an *unchanged* paragraph
 * gets back the *same* `BlockExtent` object it had last pass. A `measure` object
 * is also the larger half of a MeasuredBlock (per-line/per-segment geometry).
 * Keying the serialized measure by object identity therefore lets a keystroke
 * re-encode only the one edited block's measure (a cache miss) plus the
 * always-changing `layout`, skipping the re-serialization of every other
 * block's geometry.
 *
 * The produced string is assembled to deserialize to the identical structure as
 * `JSON.stringify(inputs)` (same key order; and serde_json is key-order /
 * whitespace agnostic regardless), so the DisplayList the builder returns is
 * byte-identical. `cache` is a WeakMap so entries die with their measure objects.
 *
 * Exported for the perf bench (`scripts/bench-displaylist.ts`); callers should
 * use `buildRustDisplayList`, which owns the process-wide cache.
 */
export function encodeDisplayListInputs(
  inputs: DisplayListBuildInputs,
  cache?: WeakMap<object, string>
): string {
  const m = inputs.measured;
  let measuredStr = '[';
  for (let i = 0; i < m.length; i++) {
    if (i > 0) measuredStr += ',';
    const mb = m[i];
    const measure = mb.measure as unknown;
    let measureStr: string | undefined;
    if (cache && measure !== null && typeof measure === 'object') {
      const key = measure as object;
      measureStr = cache.get(key);
      if (measureStr === undefined) {
        measureStr = JSON.stringify(mb.measure);
        cache.set(key, measureStr);
      }
    } else {
      measureStr = JSON.stringify(mb.measure);
    }
    measuredStr += '{"block":' + JSON.stringify(mb.block) + ',"measure":' + measureStr + '}';
  }
  measuredStr += ']';

  let out =
    '{"measured":' +
    measuredStr +
    ',"options":' +
    JSON.stringify(inputs.options) +
    ',"layout":' +
    JSON.stringify(inputs.layout);
  if (inputs.contractVersion !== undefined) {
    out += ',"contractVersion":' + JSON.stringify(inputs.contractVersion);
  }
  if (inputs.headersFooters !== undefined) {
    out += ',"headersFooters":' + JSON.stringify(inputs.headersFooters);
  }
  if (inputs.fontChains !== undefined) {
    out += ',"fontChains":' + JSON.stringify(inputs.fontChains);
  }
  if (inputs.noteAreas !== undefined) {
    out += ',"noteAreas":' + JSON.stringify(inputs.noteAreas);
  }
  if (inputs.resolvedCommentIds !== undefined) {
    out += ',"resolvedCommentIds":' + JSON.stringify(inputs.resolvedCommentIds);
  }
  if (inputs.commentAuthors !== undefined) {
    out += ',"commentAuthors":' + JSON.stringify(inputs.commentAuthors);
  }
  if (inputs.commentThreads !== undefined) {
    out += ',"commentThreads":' + JSON.stringify(inputs.commentThreads);
  }
  out += '}';
  return out;
}

/** Display-only fields sent to an engine that already retains pagination. */
export function encodeDisplayListFrameExtras(inputs: DisplayListBuildInputs): string {
  const extras: Omit<DisplayListBuildInputs, 'measured' | 'options' | 'layout'> = {};
  if (inputs.contractVersion !== undefined) extras.contractVersion = inputs.contractVersion;
  if (inputs.headersFooters !== undefined) extras.headersFooters = inputs.headersFooters;
  if (inputs.fontChains !== undefined) extras.fontChains = inputs.fontChains;
  if (inputs.noteAreas !== undefined) extras.noteAreas = inputs.noteAreas;
  if (inputs.resolvedCommentIds !== undefined) {
    extras.resolvedCommentIds = inputs.resolvedCommentIds;
  }
  if (inputs.commentAuthors !== undefined) extras.commentAuthors = inputs.commentAuthors;
  if (inputs.commentThreads !== undefined) extras.commentThreads = inputs.commentThreads;
  return JSON.stringify(extras);
}

// Process-wide cache of serialized MeasuredBlock `measure` fragments, keyed by
// the measure object's identity. Reference-stable across layout passes for
// unchanged paragraphs (the measure cache hands back the same object), so the
// heaviest part of the per-keystroke encode is reused. WeakMap ⇒ entries are
// collected with their measure objects; no lifecycle management needed.
const measureFragmentCache = new WeakMap<object, string>();

/**
 * Build a DisplayList for a computed layout through the Rust engine. Throws
 * when the wasm module fails to load or the builder rejects the input — the
 * caller is expected to fall back to the DOM painter.
 */
export async function buildRustDisplayList(
  inputs: DisplayListBuildInputs,
  engine?: RustDisplayListEngine
): Promise<DisplayList> {
  let eng: RustDisplayListEngine;
  try {
    eng = engine ?? (await loadEngine());
  } catch (error) {
    throw new RustDisplayListSourceError('load', error);
  }
  let inputJson: string;
  let json: string;
  try {
    inputJson = encodeDisplayListInputs(inputs, measureFragmentCache);
    json = eng.buildDisplayListJson(inputJson);
  } catch (error) {
    throw new RustDisplayListSourceError('build', error);
  }
  try {
    // Rust serializes the display contract in camelCase. Parsing without a
    // field-picking reviver is intentional: additive primitive metadata remains
    // present for downstream canvas and accessibility consumers.
    return JSON.parse(json) as DisplayList;
  } catch (error) {
    throw new RustDisplayListSourceError('parse', error);
  }
}

/**
 * Production display build. Resident editing engines return FrameDelta v1;
 * older/stateless engines retain the full-JSON parity path.
 */
export async function buildRustDisplayFrame(
  inputs: DisplayListBuildInputs,
  engine?: RustDisplayListEngine,
  previous: RetainedFrame | null = null
): Promise<RustDisplayFrameResult> {
  let eng: RustDisplayListEngine;
  try {
    eng = engine ?? (await loadEngine());
  } catch (error) {
    throw new RustDisplayListSourceError('load', error);
  }
  if (!eng.buildDisplayListFrame) {
    return {
      displayList: await buildRustDisplayList(inputs, eng),
      frame: null,
      transport: 'json',
    };
  }

  const inputJson = encodeDisplayListFrameExtras(inputs);
  let encoded: Uint8Array;
  try {
    encoded = eng.buildDisplayListFrame(inputJson, previous?.frameEpoch ?? 0);
  } catch (error) {
    throw new RustDisplayListSourceError('build', error);
  }
  let delta;
  try {
    delta = decodeFrameDelta(encoded);
  } catch (error) {
    throw new RustDisplayListSourceError('decode', error);
  }
  let frame: RetainedFrame;
  try {
    frame = applyFrameDelta(previous, delta);
  } catch (error) {
    throw new RustDisplayListSourceError('apply', error);
  }
  return { displayList: frame.displayList, frame, transport: 'frame-delta-v1' };
}
