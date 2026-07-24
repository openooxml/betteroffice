/**
 * Typed facade over the Rust yrs editing core (`crates/docx-edit`).
 *
 * This module is the ONLY JS entry to that crate — the `docx/zipContainer.ts`
 * facade precedent. It is types + JSON marshaling only; every editing
 * semantic (pilcrow-as-character stories, explicit-attribute inserts,
 * suggesting mode, sticky comment anchors) lives in Rust.
 *
 * `createYrsSession()` lazily loads the embedded wasm (`./wasm`) on first call.
 *
 * Addressing is the op-contract public vocabulary: a {@link YrsLoc} is
 * `(story, paraId, offset)` with UTF-16 offsets scoped to one paragraph
 * (`offset ∈ [0, paragraphLength]`; the paragraph's own pilcrow is excluded
 * from its offset space). Story-global indices never cross this boundary.
 */

import type { EditSession } from './wasm/index';
import type { Document } from '../types/document';
import { decodeS9Envelope, decodeS9EnvelopeValue } from '../docx/rustParseFacade';
import type {
  CollaborationCursor,
  CollaborationReplica,
  CollaborationTextInsertion,
  CollaborationUpdateOrigin,
} from '../collaboration/types';

export * from './inputPositionMap';
export {
  ResidentEngineWorkerClient,
  canUseResidentEngineWorker,
  type ResidentEngineWorkerApplyResult,
  type ResidentEngineWorkerFrame,
  type ResidentEngineOffscreenPage,
} from './residentEngineWorkerClient';
export {
  residentCaretSnapshotForFrame,
  residentCaretDeviceRect,
  type ResidentCaretPaintStyle,
} from './residentCaret';
export { documentToYrs } from './documentToYrs';
export { yrsToDocument } from './yrsToDocument';

export interface YrsDocxHost {
  document: Document;
  referencedFonts: string[];
  embeddedFonts: Map<string, ArrayBuffer>;
  fontTableRelationshipsXml?: string;
}

/** Durable authorship metadata for one suggested (tracked-change) operation. */
export interface YrsAuthor {
  name: string;
  /** ISO timestamp supplied by the host so all peers share one clock policy. */
  date: string;
}

/** One position: `(story, paraId, offset)` — offsets are UTF-16 units within the paragraph. */
export interface YrsLoc {
  story: string;
  paraId: string;
  offset: number;
}

/** A paragraph-addressed position without the story (used inside {@link YrsStoryRange}). */
export interface YrsParaOffset {
  paraId: string;
  offset: number;
}

/**
 * A half-open range `[start, end)` inside one story. Ends may sit in
 * different paragraphs; the boundary pilcrows are then part of the range
 * (deleting such a range merges paragraphs).
 */
export interface YrsStoryRange {
  story: string;
  start: YrsParaOffset;
  end: YrsParaOffset;
}

/** One run-level mark for {@link YrsSession.toggleMark}. */
export type YrsRunMark =
  | { type: 'bold' }
  | { type: 'italic' }
  | { type: 'underline' }
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'color'; value: string };

/** A direct text color written by {@link YrsSession.formatRange}. */
export type YrsTextColor =
  | { rgb: string; themeColor?: never }
  | { rgb?: never; themeColor: string };

/**
 * Set-valued inline formatting for {@link YrsSession.formatRange}. Omitted
 * fields are left unchanged; `null` clears a field. Boolean `false` also
 * clears the corresponding mark.
 */
export interface YrsInlineFormatDelta {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | { style?: string; color?: string } | null;
  strike?: boolean | { double?: boolean } | null;
  color?: YrsTextColor | null;
  highlight?: string | null;
  /** Font size in points; Rust writes both OOXML half-point size fields. */
  fontSize?: number | null;
  fontFamily?: { ascii: string; hAnsi?: string } | null;
  /** Passive run properties not covered by the typed fields. `null` clears. */
  other?: Readonly<Record<string, unknown | null>>;
}

export interface YrsHyperlinkAttrs {
  href: string;
  tooltip?: string | null;
  rId?: string | null;
}

/** Seed shape for {@link YrsSession.loadStories}. */
export interface YrsParagraphSeed {
  /** Paragraph text; must not contain paragraph breaks. */
  text: string;
  /** Defaults to `"Normal"`. */
  pStyle?: string;
  /** Defaults to `"left"`. */
  alignment?: string;
}

/** One story to seed via {@link YrsSession.loadStories}. */
export interface YrsStorySeed {
  storyId: string;
  /** At least one paragraph. */
  paragraphs: readonly YrsParagraphSeed[];
}

/** Snapshot of one paragraph from {@link YrsSession.paragraphs}. */
export interface YrsParagraph {
  paraId: string;
  text: string;
  /** pStyle / alignment plus any op-set extras. */
  properties: Record<string, unknown>;
}

/** One direct numbering reference (`w:numPr`) on a paragraph. */
export interface YrsNumberingProperties {
  numId?: number;
  ilvl?: number;
}

/** One paragraph tab stop, measured in twips from the left margin. */
export interface YrsParagraphTabStop {
  position: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

/**
 * Tri-state properties for {@link YrsSession.setParagraphAttrs}. Omitted
 * fields are kept and `null` clears. Numeric spacing and indents are OOXML
 * twips/line-spacing units, never CSS pixels.
 */
export interface YrsParagraphAttrs {
  alignment?:
    | 'left'
    | 'center'
    | 'right'
    | 'both'
    | 'distribute'
    | 'mediumKashida'
    | 'highKashida'
    | 'lowKashida'
    | 'thaiDistribute'
    | null;
  lineSpacing?: number | null;
  lineSpacingRule?: 'auto' | 'exact' | 'atLeast' | null;
  spaceBefore?: number | null;
  spaceAfter?: number | null;
  indentLeft?: number | null;
  indentRight?: number | null;
  indentFirstLine?: number | null;
  hangingIndent?: boolean | null;
  bidi?: boolean | null;
  tabs?: readonly YrsParagraphTabStop[] | null;
  defaultTextFormatting?: Readonly<Record<string, unknown>> | null;
  numPr?: YrsNumberingProperties | null;
  numPrFromStyle?: YrsNumberingProperties | null;
  listNumFmt?: string | null;
  listIsBullet?: boolean | null;
  listMarker?: string | null;
  listMarkerHidden?: boolean | null;
  listMarkerFontFamily?: string | null;
  listMarkerFontSize?: number | null;
  listMarkerSuffix?: 'tab' | 'space' | 'nothing' | null;
  listLevelNumFmts?: readonly string[] | null;
  listAbstractNumId?: number | null;
  listStartOverride?: number | null;
  /** Passive pPr/model properties not covered by the typed fields. */
  other?: Readonly<Record<string, unknown | null>>;
}

/** Typed value stored on a structured document tag (content control). */
export type YrsContentControlValue =
  | { kind: 'dropdown'; value: string }
  | { kind: 'checkbox'; checked: boolean }
  | { kind: 'date'; date: string }
  | string;

/** Image geometry authored in OOXML EMUs plus wrapping/position metadata. */
export interface YrsImageGeometry {
  widthEmu: number;
  heightEmu: number;
  wrap?: 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';
  hOffsetEmu?: number | null;
  vOffsetEmu?: number | null;
  distTopEmu?: number | null;
  distBottomEmu?: number | null;
  distLeftEmu?: number | null;
  distRightEmu?: number | null;
  relativeFromHorizontal?: string | null;
  relativeFromVertical?: string | null;
  /** Additional image payload fields; `null` clears an existing field. */
  other?: Readonly<Record<string, unknown | null>>;
}

/** A text or picture watermark payload for {@link YrsSession.insertWatermark}. */
export type YrsWatermark =
  | {
      kind: 'text';
      text: string;
      font: string;
      color: string;
      semitransparent: boolean;
      layout: 'diagonal' | 'horizontal';
      fontSize?: number;
      decorative?: boolean;
    }
  | {
      kind: 'picture';
      relId?: string;
      mediaPath?: string;
      dataUrl?: string;
      contentType?: string;
      scale: number;
      washout: boolean;
      widthEmu?: number;
      heightEmu?: number;
      decorative?: boolean;
    };

/**
 * One formatted segment from {@link YrsSession.storySegments} — the render
 * bridge's input. `attributes` carries run marks plus `ins` / `del` revision
 * values (suggested deletions are retained text with a `del` attribute).
 */
export type YrsStorySegment =
  | { kind: 'text'; text: string; attributes: Record<string, unknown> }
  | {
      kind: 'pilcrow';
      paraId: string;
      properties: Record<string, unknown>;
      attributes: Record<string, unknown>;
    }
  | {
      kind: 'embed';
      /** Map-backed embed discriminator (`table`, `noteRef`, `image`, …). */
      embedKind: string;
      /** Authored map entries excluding the discriminator. */
      payload: Record<string, unknown>;
      attributes: Record<string, unknown>;
    };

/** Current offsets of one sticky comment anchor. */
export interface YrsResolvedCommentAnchor {
  story: string;
  start: number;
  end: number;
}

/** A paragraph's story span; `end` is its pilcrow, so `end - start` is the paragraph length. */
export interface YrsParagraphSpan {
  start: number;
  end: number;
}

/** Compact paragraph entry used to build the live input-position projection. */
export interface YrsParagraphLength {
  paraId: string;
  /** UTF-16 text and inline embed units before the paragraph's pilcrow. */
  length: number;
}

/** Receipt of an op that minted a paragraph id. */
export interface YrsParagraphReceipt {
  paraId: string;
}

/** Receipt of an op that may have minted a tracked-change revision (suggesting mode). */
export interface YrsRevisionReceipt {
  revisionId: string | null;
}

/**
 * Receipt of {@link YrsSession.splitParagraph}. Under the S1 split the FIRST
 * half keeps the original paraId and the SECOND half is re-minted; suggesting
 * mode stamps a `pPrIns` revision on the first half.
 */
export interface YrsSplitReceipt {
  firstParaId: string;
  secondParaId: string;
  revisionId: string | null;
}

/**
 * One low-level story mirror operation for {@link YrsSession.applyRawOps}. The
 * coexistence bridge lowers a ProseMirror transaction to a list of these and
 * applies them in one yrs transaction. Indices are UTF-16 story units, each read
 * against the story state after all prior ops in the batch.
 */
export type YrsRawOp =
  | { op: 'insert'; index: number; text: string; attrs?: Record<string, unknown> }
  | { op: 'delete'; index: number; len: number }
  | { op: 'format'; index: number; len: number; attrs?: Record<string, unknown> }
  | {
      op: 'insertEmbed';
      index: number;
      kind: string;
      payload?: Record<string, unknown>;
      attrs?: Record<string, unknown>;
    }
  | { op: 'setEmbedAttr'; index: number; key: string; value: unknown }
  | {
      /**
       * Upserts the side-map comment keyed by `id`, (re-)anchoring it to
       * `ranges` — non-empty `[start, end)` story-unit spans in the batch's
       * story (sticky `Assoc::After` starts / `Assoc::Before` ends). The
       * coexistence bridge keys by the PM comment id so identity survives the
       * mirror.
       */
      op: 'setComment';
      id: string;
      ranges: ReadonlyArray<readonly [number, number]>;
      author?: string;
      date?: string;
      body?: unknown;
    }
  | {
      /** Removes the side-map comment keyed by `id`; errors when missing. */
      op: 'removeComment';
      id: string;
    };

/** Host context for {@link YrsSession.yrsBlocksForStory} (theme + list numbering). */
export interface YrsRenderEnv {
  /** Theme color name → hex (`accent1` → `4472C4`), for theme-color resolution. */
  themeColors?: Record<string, string>;
  /** The document default tab stop in twips. */
  defaultTabStopTwips?: number | null;
  /** Current page content height in CSS px, for proportional drawing constraints. */
  pageContentHeight?: number | null;
  /** yrs revision/paragraph id → dense numeric layout id (list markers, revisions). */
  numericIds?: Record<string, number>;
}

/** Receipt of {@link YrsSession.addComment}. */
export interface YrsCommentReceipt {
  commentId: string;
}

/**
 * Target of {@link YrsSession.acceptChange} / {@link YrsSession.rejectChange}:
 * one coalesced revision id (resolved at every site, in any story — the
 * `acceptChangeById` twin) or an explicit story range (the range-command twin;
 * every tracked change overlapping the range resolves, no id filtering).
 */
export type YrsChangeTarget = { revisionId: string } | YrsStoryRange;

/** Receipt of {@link YrsSession.acceptChange} / {@link YrsSession.rejectChange}. */
export interface YrsResolveReceipt {
  /** The revision ids resolved by the op (resolution order, deduplicated). */
  revisionIds: string[];
}

/** This peer's awareness selection, resolved from two yrs sticky positions. */
export interface YrsSelection {
  anchor: YrsLoc;
  head: YrsLoc;
}

export interface YrsResidentCaretRect {
  pageIndex: number;
  pageId: string;
  x: number;
  y: number;
  height: number;
}

export interface YrsResidentCaretSnapshot {
  frameEpoch: number;
  caretRect: YrsResidentCaretRect | null;
  /** Selection the rect was computed for. Absent on raw engine snapshots. */
  selection?: YrsSelection | null;
}

export function sameYrsSelection(left: YrsSelection | null, right: YrsSelection | null): boolean {
  if (!left || !right) return left === right;
  return (
    left.anchor.story === right.anchor.story &&
    left.anchor.paraId === right.anchor.paraId &&
    left.anchor.offset === right.anchor.offset &&
    left.head.story === right.head.story &&
    left.head.paraId === right.head.paraId &&
    left.head.offset === right.head.offset
  );
}

/** Opt-in internal stage timings for one resident engine input. @internal */
export interface YrsEngineApplyProfile {
  selectionMs: number;
  editMs: number;
  lowerMs: number;
  measureMs: number;
  paginateMs: number;
  displayInputMs: number;
  displayBuildMs: number;
  displayFinalizeMs: number;
  displayMs: number;
  encodeMs: number;
}

/**
 * Serializable bootstrap state for the dedicated resident-layout worker.
 * The main-thread yrs replica records the inputs that established resident
 * render, measurement, and pagination state; the worker replays them once and
 * then owns every ordinary input-to-frame transaction.
 *
 * @internal
 */
export interface YrsResidentWorkerSnapshot {
  clientId: number;
  /** Full document state, or a state-vector diff when the caller supplied
   * the worker's known vector — both apply through the same merge path. */
  state: Uint8Array;
  selection: YrsSelection | null;
  /** Empty when the caller declared the worker's fonts current
   * (`knownFontsRevision` matches); the worker then keeps its registrations. */
  fonts: Uint8Array[];
  /** Monotonic revision of the resident font set (bumped by register/clear). */
  fontsRevision: number;
  renderInputs: Array<{ story: string; env: YrsRenderEnv }>;
  measureInputs: string[];
  layoutInput: string;
  layoutWithRegions: boolean;
  layoutRevision: number;
}

/**
 * What the sync target already holds, so a snapshot can ship deltas instead
 * of the whole world.
 *
 * @internal
 */
export interface YrsResidentWorkerSyncOptions {
  /** The worker replica's last reported yrs state vector. */
  knownStateVector?: Uint8Array | null;
  /** The `fontsRevision` the worker last applied. */
  knownFontsRevision?: number | null;
}

/** A range-aggregated toggle mark. @public */
export type YrsTriState = boolean | 'mixed';

/**
 * Read-only toolbar/a11y state aggregated from yrs over one story range.
 * Toggle marks are tri-state; value marks are `null` when absent or mixed.
 *
 * @public
 */
export interface YrsSelectionContext {
  bold: YrsTriState;
  italic: YrsTriState;
  underline: YrsTriState;
  strike: YrsTriState;
  /** Uniform ASCII font family, or `null` when absent/mixed. */
  fontFamily: string | null;
  /** Uniform font size in half-points (the PM/OOXML `w:sz` unit). */
  fontSize: number | null;
  /** Uniform RGB hex or theme-color name, or `null` when absent/mixed. */
  color: string | null;
  /** Paragraph containing the range start. */
  paraId: string;
  styleId: string | null;
  alignment: string | null;
  /**
   * Full authored pilcrow property bag. Known toolbar fields retain their PM
   * names (`indentLeft`, `spaceBefore`, `lineSpacing`, `numPr`, and so on);
   * paragraph style is stored as `pStyle`.
   */
  paragraphProperties: {
    [key: string]: unknown;
    pStyle?: string;
    alignment?: string;
    indentLeft?: number;
    indentRight?: number;
    indentFirstLine?: number;
    hangingIndent?: boolean;
    spaceBefore?: number;
    spaceAfter?: number;
    lineSpacing?: number;
    lineSpacingRule?: string;
    numPr?: { numId?: number; ilvl?: number };
  };
  hasSelection: boolean;
  isMultiParagraph: boolean;
  /** The range belongs to a table-cell story. */
  inTable: boolean;
  /** The range covers exactly one non-pilcrow embed unit. */
  isSingleEmbed: boolean;
  /** Embed discriminator (`image`, `drawing`, …), else `null`. */
  embedKind: string | null;
  /** Convenience flag for a single `image` embed selection. */
  isImage: boolean;
  inInsertion: boolean;
  inDeletion: boolean;
}

/**
 * One tracked insertion, deletion, or paragraph-mark revision read from yrs.
 *
 * @public
 */
export interface YrsRevisionInfo {
  revisionId: string;
  author: string;
  date: string;
  kind:
    | 'insertion'
    | 'deletion'
    | 'pPrIns'
    | 'pPrDel'
    | 'pPrChange'
    | 'trIns'
    | 'trDel'
    | 'tableIns'
    | 'tableDel';
  story: string;
  /** Raw affected text, capped at 80 Unicode code points. */
  preview: string;
  range: YrsStoryRange;
}

/**
 * Story-local locator for one native table embed.
 *
 * @public
 */
export interface YrsTableLoc {
  story: string;
  /** Zero-based table ordinal in the parent story. */
  tableIndex: number;
}

/**
 * One table cell addressed in the resolved rectangular grid.
 *
 * @public
 */
export interface YrsCellLoc extends YrsTableLoc {
  row: number;
  column: number;
}

/**
 * Anchor-cell to head-cell rectangular selection, analogous to PM's
 * `CellSelection` but independent of ProseMirror positions.
 *
 * @public
 */
export interface YrsTableRange {
  anchor: YrsCellLoc;
  head: YrsCellLoc;
}

/**
 * Receipt shared by native table-structure and cell-format operations.
 *
 * @public
 */
export interface YrsTableReceipt {
  table: YrsTableLoc;
  rows: number;
  columns: number;
  createdStoryIds: string[];
  deletedStoryIds: string[];
  newParaIds: string[];
  deletedTable: boolean;
  revisionIds: string[];
}

/**
 * OOXML/PM-shaped cell border value passed through to `tcPr.borders`.
 *
 * @public
 */
export interface YrsCellBorder {
  style: string;
  size?: number;
  color?: { rgb: string };
}

/**
 * Complete per-side border object for {@link YrsSession.setCellBorders}.
 *
 * @public
 */
export type YrsCellBorders = Partial<
  Record<'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV', YrsCellBorder>
>;

/**
 * One live replica of the yrs editing model. Thin typed wrapper over the
 * wasm `EditSession` — no editing logic on this side of the boundary.
 */
export interface YrsSession extends CollaborationReplica {
  /** The yrs client id this replica writes with. */
  readonly clientId: number;

  // -- resident layout engine (same wasm instance as EditingDoc) --

  /** Register raw sfnt bytes in the session's measurement/display font store. */
  registerFont(bytes: Uint8Array): number;
  /** Clear the session's registered measurement/display fonts. */
  clearFonts(): void;
  /** Measure one paragraph through the session's resident text engine. */
  measureParagraphJson(input: string): string;
  /** Paginate and retain the measured input and Layout in the session. */
  layoutDocumentJson(input: string): string;
  /** Return the compact font requirements for resident region layout. */
  layoutFontRequirementsJson(input: string): string;
  /** Paginate and compose section/page regions in the resident engine. */
  layoutDocumentWithRegionsJson(input: string): string;
  /** Build display primitives against the session's resident font store. */
  buildDisplayListJson(input: string): string;
  /** Build a binary FrameDelta v1 against the last host-applied frame. */
  buildDisplayListFrame(input: string, expectedFrameEpoch: number): Uint8Array;
  /** Caret geometry from the current resident display frame. */
  residentCaretSnapshot(): YrsResidentCaretSnapshot;
  /** Apply a collapsed plain-text insertion and return its resident FrameDelta. */
  applyInput(text: string, expectedFrameEpoch: number): Uint8Array;
  /** Apply a collapsed character deletion/paragraph merge and return its resident FrameDelta. */
  applyDelete(direction: 'backward' | 'forward', expectedFrameEpoch: number): Uint8Array;
  /** Instrumented apply used only by opt-in browser performance traces. */
  applyInputProfiled(
    text: string,
    expectedFrameEpoch: number
  ): { frame: Uint8Array; profile: YrsEngineApplyProfile };
  /** Instrumented deletion used only by opt-in browser performance traces. */
  applyDeleteProfiled(
    direction: 'backward' | 'forward',
    expectedFrameEpoch: number
  ): { frame: Uint8Array; profile: YrsEngineApplyProfile };
  /** Snapshot the inputs needed to move resident layout ownership to a worker. */
  residentWorkerSnapshot(options?: YrsResidentWorkerSyncOptions): YrsResidentWorkerSnapshot | null;
  /**
   * Cheap worker-sync probe: the resident layout revision when a worker
   * snapshot would be available, without encoding document state or copying
   * font bytes. Steady-state frame builds consult this instead of building a
   * full snapshot.
   */
  residentWorkerProbe(): { layoutRevision: number } | null;
  /** Resident display-list hit/range queries; results are small JSON records. */
  displayHitTestRegionsJson(pageIndex: number, x: number, y: number): string;
  displayRangeRectsJson(from: number, to: number): string;
  displayRangeRectsRegionJson(
    region: 'body' | 'header' | 'footer',
    rId: string,
    from: number,
    to: number
  ): string;
  /** Read a glyph outline from the session's resident font store. */
  outlineGlyphJson(fontId: number, glyphId: number): string;

  // -- lifecycle --

  /** Hydrates from an encoded yrs v1 update (typically a peer's {@link encodeState} output). */
  loadState(update: Uint8Array): void;
  /** Parses a DOCX, seeds its stories, and returns thin host metadata. */
  seedFromDocx(bytes: Uint8Array): YrsDocxHost;
  /** Parses a DOCX and optionally seeds its stories. */
  openDocx(bytes: Uint8Array, seedStories: boolean): YrsDocxHost;
  /** Materializes the retained canonical package for compatibility APIs. */
  materializeDocx(): Document | null;
  /**
   * Seeds stories from parsed content (S1 scaffold; the real
   * `load(ParsedDocument)` lands with the ops track). Returns paraIds per
   * story in document order.
   */
  loadStories(stories: readonly YrsStorySeed[]): Record<string, string[]>;
  /** Full document state as one yrs v1 update (Yjs wire format). */
  encodeState(): Uint8Array;
  /** Current Yrs state vector in the Yjs v1 wire format. */
  encodeStateVector(): Uint8Array;
  /** Full state, or only the state missing from a peer vector. */
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  /** Applies a remote/incremental yrs v1 update. */
  applyUpdate(update: Uint8Array): CollaborationTextInsertion | null;
  /** Apply a same-user worker update under the local undo origin. @internal */
  applyLocalUpdate(update: Uint8Array, story: string): void;
  /**
   * Subscribes to every committed transaction's v1 update (local AND
   * applied-remote). Returns an unsubscribe function.
   */
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;

  // -- local input state --

  /** Store this peer's awareness selection as sticky positions. */
  setSelection(anchor: YrsLoc, head?: YrsLoc): void;
  /** Resolve this peer's current sticky selection, or null before initialization. */
  selection(): YrsSelection | null;
  /** Encode the current selection as binary Yrs sticky indices. */
  encodeSelection(): CollaborationCursor | null;
  /** Resolve a peer's binary Yrs sticky indices against this replica. */
  resolveSelection(cursor: CollaborationCursor): YrsSelection | null;
  /** Store this peer's rectangular table selection outside the document. */
  setCellSelection(range: YrsTableRange): void;
  /** Resolve the current sticky cell selection, or null before initialization. */
  cellSelection(): YrsTableRange | null;
  /** Lazily begin local-origin undo capture after import/seeding has completed. */
  beginUndoCapture(story: string, includeTableStories?: boolean): void;
  /** Coalesce the stack entries added since `startDepth` into one host undo intent. */
  markUndoGroup(startDepth: number): void;
  /** Undo/redo only local-origin direct operations (never remote/system transactions). */
  undo(): boolean;
  redo(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  /** Current local undo/redo stack sizes (zero before tracking starts). */
  undoDepth(): number;
  redoDepth(): number;

  // -- S1 ops --

  /** Adds a story with one paragraph; the receipt carries its paraId. */
  createStory(
    storyId: string,
    initialText: string,
    pStyle?: string,
    alignment?: string
  ): YrsParagraphReceipt;
  /** Removes a complete story (including an unreachable table-cell story). */
  deleteStory(storyId: string): void;
  /** Inserts a row above or below the cell that `at` resolves into. */
  insertRow(at: YrsCellLoc, side: 'above' | 'below', suggesting?: YrsAuthor): YrsTableReceipt;
  /** Inserts a rectangular structural table at a paragraph-keyed location. */
  insertTable(at: YrsLoc, rows: number, columns: number, suggesting?: YrsAuthor): YrsTableReceipt;
  /** Inserts a column left or right of the cell that `at` resolves into. */
  insertColumn(at: YrsCellLoc, side: 'left' | 'right'): YrsTableReceipt;
  /** Deletes every row covered by an explicit rectangular cell range. */
  deleteRow(range: YrsTableRange, suggesting?: YrsAuthor): YrsTableReceipt;
  /** Deletes every column covered by an explicit rectangular cell range. */
  deleteColumn(range: YrsTableRange): YrsTableReceipt;
  /** Removes a complete table plus its reachable cell stories. */
  deleteTable(table: YrsTableLoc): YrsTableReceipt;
  /** Merges a rectangular range into its top-left cell. */
  mergeCells(range: YrsTableRange): YrsTableReceipt;
  /** Splits the merged cell covering `at` into one cell per grid slot. */
  splitCell(at: YrsCellLoc, rows?: number, columns?: number): YrsTableReceipt;
  /** Sets or clears the selected cells' background color. */
  setCellShading(range: YrsTableRange, color: string | null): YrsTableReceipt;
  /**
   * Merges an OOXML/PM-shaped patch into selected cells' `tcPr`. JSON `null`
   * clears a property; merge/split-owned span keys are rejected.
   */
  setCellTextFormat(
    range: YrsTableRange,
    patch: Readonly<Record<string, unknown>>
  ): YrsTableReceipt;
  /** Replaces the complete selected-cell border property object. */
  setCellBorders(range: YrsTableRange, borders: YrsCellBorders): YrsTableReceipt;
  /** Sets one authored grid-column width in twips. */
  setColumnWidth(at: YrsCellLoc, widthTwips: number): YrsTableReceipt;
  /** Sets the table-wide preferred width in twips. */
  setTableWidth(table: YrsTableLoc, widthTwips: number): YrsTableReceipt;
  /** Inserts paragraph-break-free text. Suggesting mode mints a revision. */
  insertText(at: YrsLoc, text: string, suggesting?: YrsAuthor): YrsRevisionReceipt;
  /**
   * Deletes a range (plain) or marks it as a suggested deletion (suggesting).
   * A range spanning paragraphs also merges them (pilcrow-as-character).
   */
  deleteRange(range: YrsStoryRange, suggesting?: YrsAuthor): YrsRevisionReceipt;
  /** Replaces a range with text in one transaction (one shared revision when suggesting). */
  replaceRange(range: YrsStoryRange, text: string, suggesting?: YrsAuthor): YrsRevisionReceipt;
  /**
   * Splits a paragraph by inserting one pilcrow. The FIRST half keeps the
   * original paraId; the SECOND half is re-minted (`secondParaId`).
   */
  splitParagraph(at: YrsLoc, suggesting?: YrsAuthor): YrsSplitReceipt;
  /** Merges `paraId` with the FOLLOWING paragraph. Errors on the final paragraph. */
  mergeParagraphs(story: string, paraId: string, suggesting?: YrsAuthor): YrsRevisionReceipt;
  /** Toggles one run mark across a range (PM toggleMark range semantics). */
  toggleMark(range: YrsStoryRange, mark: YrsRunMark): void;
  /** Applies set-valued direct formatting; omitted fields are kept and `null` fields clear. */
  formatRange(range: YrsStoryRange, delta: YrsInlineFormatDelta): void;
  /** Sets or clears the protected hyperlink attribute over a non-empty range. */
  setHyperlink(range: YrsStoryRange, hyperlink: YrsHyperlinkAttrs | null): void;
  /** Clears direct formatting while retaining hyperlinks and tracked-change stamps. */
  clearFormatting(range: YrsStoryRange): void;
  /** Applies a paragraph style id to every paragraph intersecting the range. */
  applyParagraphStyle(range: YrsStoryRange, styleId: string, suggesting?: YrsAuthor): void;
  /** Applies tri-state paragraph properties to every paragraph intersecting the range. */
  setParagraphAttrs(range: YrsStoryRange, attrs: YrsParagraphAttrs, suggesting?: YrsAuthor): void;
  /** Inserts one inline image embed, optionally as a tracked insertion. */
  insertImage(
    at: YrsLoc,
    image: Readonly<Record<string, unknown>>,
    suggesting?: YrsAuthor
  ): YrsRevisionReceipt;
  /** Sets the authored value on a content-control embed addressed by stable payload id. */
  setContentControlValue(embedId: string, value: YrsContentControlValue): void;
  /** Sets a content-control value at a paragraph-keyed embed position. */
  setContentControlValueAt(at: YrsLoc, value: YrsContentControlValue): void;
  /** Removes the authored value from a content-control embed. */
  clearContentControlValue(embedId: string): void;
  /** Commits image size/wrapping/position fields in one transaction. */
  setImageGeometry(embedId: string, geometry: YrsImageGeometry): void;
  /** Inserts a native page-break embed at a paragraph-keyed location. */
  insertPageBreak(at: YrsLoc): void;
  /** Inserts a native section-break embed at a paragraph-keyed location. */
  insertSectionBreak(at: YrsLoc, type: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage'): void;
  /** Inserts a typed watermark embed at a paragraph-keyed location. */
  insertWatermark(at: YrsLoc, watermark: YrsWatermark): void;
  /**
   * Applies a batch of raw story mirror ops in one transaction — the coexistence
   * bridge's mirror-into-yrs path (a faithful mirror of lowered PM state, not a
   * user-intent op). Not for direct app use.
   */
  applyRawOps(story: string, ops: readonly YrsRawOp[]): void;
  /** Sets one paragraph property (any JSON value). `paraId` is reserved. */
  setParagraphAttr(paraId: string, key: string, value: unknown): void;
  /** Adds a sticky-anchored comment over one or more ranges. */
  addComment(
    ranges: readonly YrsStoryRange[],
    author: string,
    date: string,
    body: unknown
  ): YrsCommentReceipt;
  /**
   * Accepts tracked changes (S4b): pending insertions become plain content,
   * pending deletions are carried out; a `pPrIns` paragraph mark clears (the
   * split stays), a `pPrDel` mark joins with the following paragraph (whose
   * pPr survives — Word's surviving-`w:p` rule). Resolving never stamps a new
   * revision. Throws on an unknown revision id.
   */
  acceptChange(target: YrsChangeTarget): YrsResolveReceipt;
  /**
   * Rejects tracked changes — the inverse of {@link acceptChange}: pending
   * insertions roll back, pending deletions restore their text; a `pPrIns`
   * mark joins back with the following paragraph, a `pPrDel` mark clears.
   */
  rejectChange(target: YrsChangeTarget): YrsResolveReceipt;

  // -- read queries --

  /** Aggregate toolbar/a11y state from yrs over one paragraph-addressed range. */
  selectionContext(range: YrsStoryRange): YrsSelectionContext;
  /** Enumerate tracked changes across every story in deterministic order. */
  listRevisions(): YrsRevisionInfo[];
  /** Current offsets of a comment's sticky anchors. Throws when an anchor no longer resolves. */
  resolveComment(commentId: string): YrsResolvedCommentAnchor[];
  /** Story ids in the document, sorted. */
  storyIds(): string[];
  /** Story length in UTF-16 units (every embed, pilcrows included, counts 1). */
  storyLength(story: string): number;
  /**
   * The story's `canonical-stream-v1` FNV-1a checksum. The coexistence watchdog
   * compares it against the PM projector's checksum after every mirrored edit.
   */
  storyChecksum(story: string): bigint;
  /**
   * Lowers a story through the resident Rust bridge. Throws with an
   * unsupported-embed message on any non-native content (opaque blobs) until
   * that class is promoted to native.
   * The return is the layout pipeline's `LayoutBlock[]` (kept as `unknown[]` to
   * keep this facade decoupled from the layout types).
   */
  yrsBlocksForStory(story: string, env?: YrsRenderEnv): unknown[];
  /** Paragraph snapshots in document order. */
  paragraphs(story: string): YrsParagraph[];
  /** Paragraph ids and inline-unit lengths, resolved in one Rust story traversal. */
  paragraphSpans(story: string): YrsParagraphLength[];
  /** The raw formatted-segment view (the render bridge's input). */
  storySegments(story: string): YrsStorySegment[];
  /** A paragraph's story span (start unit, pilcrow index). */
  locateParagraph(story: string, paraId: string): YrsParagraphSpan;

  /** Drops the observer and frees the wasm-side replica. Idempotent. */
  destroy(): void;
}

/** Options for {@link createYrsSession}. */
export interface CreateYrsSessionOptions {
  /**
   * The yrs client id (non-negative safe integer). Omit to allocate a random
   * 32-bit id, yjs-style.
   */
  clientId?: number;
}

function randomClientId(): number {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const buffer = new Uint32Array(1);
    crypto.getRandomValues(buffer);
    return buffer[0];
  }
  return Math.floor(Math.random() * 0xffffffff);
}

function wireChangeTarget(target: YrsChangeTarget): string {
  return JSON.stringify(
    'revisionId' in target
      ? { revisionId: target.revisionId }
      : {
          story: target.story,
          startPara: target.start.paraId,
          startOffset: target.start.offset,
          endPara: target.end.paraId,
          endOffset: target.end.offset,
        }
  );
}

function wireRanges(ranges: readonly YrsStoryRange[]): string {
  return JSON.stringify(
    ranges.map((range) => ({
      story: range.story,
      startPara: range.start.paraId,
      startOffset: range.start.offset,
      endPara: range.end.paraId,
      endOffset: range.end.offset,
    }))
  );
}

function docxSourceBuffer(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === bytes.buffer.byteLength
  ) {
    return bytes.buffer;
  }
  return bytes.slice().buffer as ArrayBuffer;
}

function decodeDocxHost(json: string, source: Uint8Array): YrsDocxHost {
  const value: unknown = JSON.parse(json);
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('DOCX host metadata must be an object');
  }
  const wire = value as Record<string, unknown>;
  if (
    !Array.isArray(wire.referencedFonts) ||
    !wire.referencedFonts.every((name) => typeof name === 'string')
  ) {
    throw new TypeError('DOCX host referencedFonts must be a string array');
  }
  const result = decodeS9EnvelopeValue(wire.envelope, docxSourceBuffer(source));
  return {
    document: result.document,
    referencedFonts: wire.referencedFonts,
    embeddedFonts: result.embeddedFonts,
    ...(result.fontTableRelationshipsXml === undefined
      ? {}
      : { fontTableRelationshipsXml: result.fontTableRelationshipsXml }),
  };
}

function wrapSession(session: EditSession, clientId: number): YrsSession {
  const listeners = new Map<
    number,
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  const pendingUpdates: Array<{
    update: Uint8Array;
    origin: CollaborationUpdateOrigin;
  }> = [];
  let observing = false;
  let destroyed = false;
  let nextListenerId = 0;
  let wasmCallDepth = 0;
  let flushingUpdates = false;
  let undoStory: string | null = null;
  let cachedSelection: YrsSelection | null | undefined;
  let cachedSelectionContext: { key: string; json: string } | null = null;
  const undoGroups = new Map<number, number>();
  const redoGroups = new Map<number, number>();
  const residentFonts: Uint8Array[] = [];
  const residentRenderInputs = new Map<string, YrsRenderEnv>();
  const residentMeasureInputs = new Map<string, string>();
  let residentLayoutInput: string | null = null;
  let residentLayoutWithRegions = false;
  let residentLayoutRevision = 0;
  let residentFontsRevision = 0;
  let ownsResidentFontStore = false;
  let docxSource: Uint8Array | null = null;

  const invalidateReadCaches = (): void => {
    cachedSelection = undefined;
    cachedSelectionContext = null;
  };

  const flushUpdates = (): void => {
    if (destroyed || flushingUpdates || wasmCallDepth !== 0) return;
    flushingUpdates = true;
    try {
      while (!destroyed && pendingUpdates.length > 0) {
        const event = pendingUpdates.shift();
        if (!event) break;
        for (const [id, listener] of [...listeners]) {
          if (destroyed) return;
          if (listeners.get(id) !== listener) continue;
          try {
            listener(event.update.slice(), event.origin);
          } catch {}
        }
      }
    } finally {
      flushingUpdates = false;
      if (destroyed) pendingUpdates.length = 0;
    }
  };

  const mutate = <T>(operation: () => T): T => {
    invalidateReadCaches();
    wasmCallDepth += 1;
    try {
      return operation();
    } finally {
      wasmCallDepth -= 1;
      if (wasmCallDepth === 0) flushUpdates();
    }
  };

  const cloneSelection = (value: YrsSelection | null): YrsSelection | null =>
    value
      ? {
          anchor: { ...value.anchor },
          head: { ...value.head },
        }
      : null;

  const ensureUndo = (story: string): void => {
    if (undoStory === story) return;
    session.track_undo(story);
    undoStory = story;
  };

  const ensureTableUndo = (story: string): void => {
    const scope = `table:${story}`;
    if (undoStory === scope) return;
    session.track_table_undo(story);
    undoStory = scope;
  };

  const storyForEmbedId = (embedId: string): string | null => {
    const matches = (value: unknown): boolean =>
      (typeof value === 'string' && value === embedId) ||
      (typeof value === 'number' && Number.isFinite(value) && String(value) === embedId);
    for (const story of session.story_ids()) {
      const segments = JSON.parse(session.story_segments(story)) as YrsStorySegment[];
      if (
        segments.some(
          (segment) =>
            segment.kind === 'embed' &&
            (matches(segment.payload.embedId) ||
              matches(segment.payload.id) ||
              matches(segment.payload.rId))
        )
      ) {
        return story;
      }
    }
    return null;
  };

  const ensureEmbedUndo = (embedId: string): void => {
    const story = storyForEmbedId(embedId);
    if (story) ensureUndo(story);
  };

  const ensureObserver = () => {
    if (observing) return;
    session.set_update_observer((update: Uint8Array, origin: number) => {
      if (origin !== 0 && origin !== 1) return;
      pendingUpdates.push({
        update: update.slice(),
        origin: origin === 0 ? 'local' : 'remote',
      });
      flushUpdates();
    });
    observing = true;
  };

  const clearUnusedObserver = (): void => {
    if (!observing || listeners.size > 0 || destroyed) return;
    pendingUpdates.length = 0;
    session.clear_update_observer();
    observing = false;
  };

  const openDocx = (bytes: Uint8Array, seedStories: boolean): YrsDocxHost => {
    const source = bytes.slice();
    const json = mutate(() => session.open_docx(source, seedStories));
    const host = decodeDocxHost(json, source);
    docxSource = source;
    return host;
  };

  return {
    clientId,

    registerFont: (bytes) => {
      // The Rust font store is module-global, while document sessions are
      // replaceable. Claim a fresh id space on the first registration for a
      // new session so a worker replay sees the same dense ids (0..N) after a
      // document load; otherwise ids would retain gaps from the old session.
      if (!ownsResidentFontStore) {
        session.clear_measure_fonts();
        ownsResidentFontStore = true;
      }
      const id = session.register_measure_font(bytes);
      residentFonts.push(bytes.slice());
      residentFontsRevision += 1;
      return id;
    },
    clearFonts: () => {
      session.clear_measure_fonts();
      residentFonts.length = 0;
      residentMeasureInputs.clear();
      residentFontsRevision += 1;
      ownsResidentFontStore = true;
    },
    measureParagraphJson: (input) => {
      const output = session.measure_paragraph_json(input);
      residentMeasureInputs.set(input, input);
      return output;
    },
    layoutDocumentJson: (input) => {
      const output = session.layout_document_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = false;
      residentLayoutRevision += 1;
      return output;
    },
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsJson: (input) => {
      const output = session.layout_document_with_regions_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      return output;
    },
    buildDisplayListJson: (input) => session.build_display_list_json(input),
    buildDisplayListFrame: (input, expectedFrameEpoch) =>
      session.build_display_list_frame(input, expectedFrameEpoch),
    residentCaretSnapshot: () =>
      JSON.parse(session.resident_caret_snapshot_json()) as YrsResidentCaretSnapshot,
    applyInput: (text, expectedFrameEpoch) => {
      const story = cachedSelection?.head.story ?? 'body';
      ensureUndo(story);
      return mutate(() => session.apply_input(text, expectedFrameEpoch));
    },
    applyDelete: (direction, expectedFrameEpoch) => {
      const story = cachedSelection?.head.story ?? 'body';
      ensureUndo(story);
      return mutate(() => session.apply_delete(direction, expectedFrameEpoch));
    },
    applyInputProfiled: (text, expectedFrameEpoch) => {
      const story = cachedSelection?.head.story ?? 'body';
      ensureUndo(story);
      const frame = mutate(() => session.apply_input_profiled(text, expectedFrameEpoch));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch) => {
      const story = cachedSelection?.head.story ?? 'body';
      ensureUndo(story);
      const frame = mutate(() => session.apply_delete_profiled(direction, expectedFrameEpoch));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    residentWorkerSnapshot: (options) => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      const selectionJson = session.selection();
      const fontsCurrent = options?.knownFontsRevision === residentFontsRevision;
      let state: Uint8Array | null = null;
      if (options?.knownStateVector) {
        try {
          state = session.encode_diff(options.knownStateVector.slice());
        } catch {
          state = null;
        }
      }
      return {
        clientId,
        state: state ?? session.encode_state(),
        selection: JSON.parse(selectionJson) as YrsSelection | null,
        fonts: fontsCurrent ? [] : residentFonts.map((bytes) => bytes.slice()),
        fontsRevision: residentFontsRevision,
        renderInputs: [...residentRenderInputs].map(([story, env]) => ({
          story,
          env: structuredClone(env),
        })),
        measureInputs: [...residentMeasureInputs.values()],
        layoutInput: residentLayoutInput,
        layoutWithRegions: residentLayoutWithRegions,
        layoutRevision: residentLayoutRevision,
      };
    },
    residentWorkerProbe: () => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      return { layoutRevision: residentLayoutRevision };
    },
    displayHitTestRegionsJson: (pageIndex, x, y) =>
      session.display_hit_test_regions_json(pageIndex, x, y),
    displayRangeRectsJson: (from, to) => session.display_range_rects_json(from, to),
    displayRangeRectsRegionJson: (region, rId, from, to) =>
      session.display_range_rects_region_json(region, rId, from, to),
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),

    loadState: (update) => mutate(() => session.load(update)),
    seedFromDocx: (bytes) => openDocx(bytes, true),
    openDocx,
    materializeDocx: () => {
      const source = docxSource;
      const json = session.materialize_docx();
      if (!source || json === undefined) return null;
      return decodeS9Envelope(json, docxSourceBuffer(source)).document;
    },
    loadStories: (stories) =>
      mutate(
        () => JSON.parse(session.load_json(JSON.stringify(stories))) as Record<string, string[]>
      ),
    encodeState: () => session.encode_state(),
    encodeStateVector: () => session.encode_state_vector(),
    encodeStateAsUpdate: (remoteStateVector) =>
      remoteStateVector === undefined
        ? session.encode_state()
        : session.encode_diff(remoteStateVector.slice()),
    applyUpdate: (update) =>
      mutate(
        () =>
          JSON.parse(
            session.apply_update_with_inference(update)
          ) as CollaborationTextInsertion | null
      ),
    applyLocalUpdate: (update, story) => {
      ensureUndo(story);
      mutate(() => session.apply_local_update(update));
    },
    onUpdate: (listener) => {
      if (destroyed) throw new Error('yrs session is destroyed');
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      const id = nextListenerId++;
      listeners.set(id, listener);
      ensureObserver();
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        listeners.delete(id);
        clearUnusedObserver();
      };
    },

    setSelection: (anchor, head = anchor) => {
      if (anchor.story !== head.story) throw new Error('yrs selection must stay inside one story');
      session.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
      cachedSelection = {
        anchor: { ...anchor },
        head: { ...head },
      };
    },
    selection: () => {
      if (cachedSelection !== undefined) return cloneSelection(cachedSelection);
      cachedSelection = JSON.parse(session.selection()) as YrsSelection | null;
      return cloneSelection(cachedSelection);
    },
    encodeSelection: () => {
      const encoded = JSON.parse(session.encoded_selection()) as {
        story: string;
        anchor: number[];
        head: number[];
      } | null;
      return encoded
        ? {
            story: encoded.story,
            anchor: Uint8Array.from(encoded.anchor),
            head: Uint8Array.from(encoded.head),
          }
        : null;
    },
    resolveSelection: (cursor) => {
      try {
        return JSON.parse(
          session.resolve_encoded_selection(cursor.story, cursor.anchor, cursor.head)
        ) as YrsSelection;
      } catch {
        return null;
      }
    },
    setCellSelection: (range) => session.set_cell_selection(JSON.stringify(range)),
    cellSelection: () => JSON.parse(session.cell_selection()) as YrsTableRange | null,
    beginUndoCapture: (story, includeTableStories = false) =>
      includeTableStories ? ensureTableUndo(story) : ensureUndo(story),
    markUndoGroup: (startDepth) => {
      const endDepth = session.undo_depth();
      const size = Math.max(0, endDepth - startDepth);
      if (size > 1) undoGroups.set(endDepth, size);
      redoGroups.clear();
    },
    undo: () =>
      mutate(() => {
        const depth = session.undo_depth();
        const count = undoGroups.get(depth) ?? 1;
        let changed = false;
        for (let index = 0; index < count; index += 1) changed = session.undo() || changed;
        if (changed && count > 1) {
          undoGroups.delete(depth);
          redoGroups.set(session.redo_depth(), count);
        }
        return changed;
      }),
    redo: () =>
      mutate(() => {
        const depth = session.redo_depth();
        const count = redoGroups.get(depth) ?? 1;
        let changed = false;
        for (let index = 0; index < count; index += 1) changed = session.redo() || changed;
        if (changed && count > 1) {
          redoGroups.delete(depth);
          undoGroups.set(session.undo_depth(), count);
        }
        return changed;
      }),
    canUndo: () => session.can_undo(),
    canRedo: () => session.can_redo(),
    undoDepth: () => session.undo_depth(),
    redoDepth: () => session.redo_depth(),

    createStory: (storyId, initialText, pStyle = 'Normal', alignment = 'left') =>
      mutate(
        () =>
          JSON.parse(session.create_story(storyId, initialText, pStyle, alignment)) as {
            paraId: string;
          }
      ),
    deleteStory: (storyId) => mutate(() => session.delete_story(storyId)),
    insertTable: (at, rows, columns, suggesting) => {
      ensureTableUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_table(
              at.story,
              at.paraId,
              at.offset,
              rows,
              columns,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsTableReceipt
      );
    },
    insertRow: (at, side, suggesting) => {
      ensureTableUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_row(
              JSON.stringify(at),
              side === 'below',
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsTableReceipt
      );
    },
    insertColumn: (at, side) => {
      ensureTableUndo(at.story);
      return mutate(
        () =>
          JSON.parse(session.insert_column(JSON.stringify(at), side === 'right')) as YrsTableReceipt
      );
    },
    deleteRow: (range, suggesting) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () =>
          JSON.parse(
            session.delete_row(JSON.stringify(range), suggesting?.name, suggesting?.date)
          ) as YrsTableReceipt
      );
    },
    deleteColumn: (range) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () => JSON.parse(session.delete_column(JSON.stringify(range))) as YrsTableReceipt
      );
    },
    deleteTable: (table) => {
      ensureTableUndo(table.story);
      return mutate(
        () => JSON.parse(session.delete_table(JSON.stringify(table))) as YrsTableReceipt
      );
    },
    mergeCells: (range) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () => JSON.parse(session.merge_cells(JSON.stringify(range))) as YrsTableReceipt
      );
    },
    splitCell: (at, rows, columns) => {
      ensureTableUndo(at.story);
      return mutate(
        () => JSON.parse(session.split_cell(JSON.stringify(at), rows, columns)) as YrsTableReceipt
      );
    },
    setCellShading: (range, color) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () =>
          JSON.parse(
            session.set_cell_shading(JSON.stringify(range), color ?? undefined)
          ) as YrsTableReceipt
      );
    },
    setCellTextFormat: (range, patch) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () =>
          JSON.parse(
            session.set_cell_text_format(JSON.stringify(range), JSON.stringify(patch))
          ) as YrsTableReceipt
      );
    },
    setCellBorders: (range, borders) => {
      ensureTableUndo(range.anchor.story);
      return mutate(
        () =>
          JSON.parse(
            session.set_cell_borders(JSON.stringify(range), JSON.stringify(borders))
          ) as YrsTableReceipt
      );
    },
    setColumnWidth: (at, widthTwips) => {
      ensureTableUndo(at.story);
      return mutate(
        () =>
          JSON.parse(session.set_column_width(JSON.stringify(at), widthTwips)) as YrsTableReceipt
      );
    },
    setTableWidth: (table, widthTwips) => {
      ensureTableUndo(table.story);
      return mutate(
        () =>
          JSON.parse(session.set_table_width(JSON.stringify(table), widthTwips)) as YrsTableReceipt
      );
    },
    insertText: (at, text, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_text(
              at.story,
              at.paraId,
              at.offset,
              text,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsRevisionReceipt
      );
    },
    deleteRange: (range, suggesting) => {
      ensureUndo(range.story);
      return mutate(
        () =>
          JSON.parse(
            session.delete_range(
              range.story,
              range.start.paraId,
              range.start.offset,
              range.end.paraId,
              range.end.offset,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsRevisionReceipt
      );
    },
    replaceRange: (range, text, suggesting) => {
      ensureUndo(range.story);
      return mutate(
        () =>
          JSON.parse(
            session.replace_range(
              range.story,
              range.start.paraId,
              range.start.offset,
              range.end.paraId,
              range.end.offset,
              text,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsRevisionReceipt
      );
    },
    splitParagraph: (at, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.split_paragraph(
              at.story,
              at.paraId,
              at.offset,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsSplitReceipt
      );
    },
    mergeParagraphs: (story, paraId, suggesting) => {
      ensureUndo(story);
      return mutate(
        () =>
          JSON.parse(
            session.merge_paragraphs(story, paraId, suggesting?.name, suggesting?.date)
          ) as YrsRevisionReceipt
      );
    },
    toggleMark: (range, mark) => {
      ensureUndo(range.story);
      mutate(() =>
        session.toggle_mark(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(mark)
        )
      );
    },
    formatRange: (range, delta) => {
      ensureUndo(range.story);
      mutate(() =>
        session.format_range(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(delta)
        )
      );
    },
    setHyperlink: (range, hyperlink) => {
      ensureUndo(range.story);
      mutate(() =>
        session.set_hyperlink(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(hyperlink)
        )
      );
    },
    clearFormatting: (range) => {
      ensureUndo(range.story);
      mutate(() =>
        session.clear_formatting(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset
        )
      );
    },
    applyParagraphStyle: (range, styleId, suggesting) => {
      ensureUndo(range.story);
      mutate(() =>
        session.apply_paragraph_style(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          styleId,
          suggesting?.name,
          suggesting?.date
        )
      );
    },
    setParagraphAttrs: (range, attrs, suggesting) => {
      ensureUndo(range.story);
      mutate(() =>
        session.set_paragraph_attrs(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(attrs),
          suggesting?.name,
          suggesting?.date
        )
      );
    },
    insertImage: (at, image, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_image(
              at.story,
              at.paraId,
              at.offset,
              JSON.stringify(image),
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsRevisionReceipt
      );
    },
    setContentControlValue: (embedId, value) => {
      ensureEmbedUndo(embedId);
      mutate(() => session.set_content_control_value(embedId, JSON.stringify(value)));
    },
    setContentControlValueAt: (at, value) => {
      ensureUndo(at.story);
      mutate(() =>
        session.set_content_control_value_at(at.story, at.paraId, at.offset, JSON.stringify(value))
      );
    },
    clearContentControlValue: (embedId) => {
      ensureEmbedUndo(embedId);
      mutate(() => session.clear_content_control_value(embedId));
    },
    setImageGeometry: (embedId, geometry) => {
      ensureEmbedUndo(embedId);
      mutate(() => session.set_image_geometry(embedId, JSON.stringify(geometry)));
    },
    insertPageBreak: (at) => {
      ensureUndo(at.story);
      mutate(() => session.insert_page_break(at.story, at.paraId, at.offset));
    },
    insertSectionBreak: (at, type) => {
      ensureUndo(at.story);
      mutate(() => session.insert_section_break(at.story, at.paraId, at.offset, type));
    },
    insertWatermark: (at, watermark) => {
      ensureUndo(at.story);
      mutate(() =>
        session.insert_watermark(at.story, at.paraId, at.offset, JSON.stringify(watermark))
      );
    },
    applyRawOps: (story, ops) => mutate(() => session.apply_raw_ops(story, JSON.stringify(ops))),
    setParagraphAttr: (paraId, key, value) =>
      mutate(() => session.set_paragraph_attr(paraId, key, JSON.stringify(value ?? null))),
    addComment: (ranges, commentAuthor, date, body) =>
      mutate(
        () =>
          JSON.parse(
            session.add_comment(
              wireRanges(ranges),
              commentAuthor,
              date,
              JSON.stringify(body ?? null)
            )
          ) as YrsCommentReceipt
      ),
    acceptChange: (target) =>
      mutate(
        () => JSON.parse(session.accept_change(wireChangeTarget(target))) as YrsResolveReceipt
      ),
    rejectChange: (target) =>
      mutate(
        () => JSON.parse(session.reject_change(wireChangeTarget(target))) as YrsResolveReceipt
      ),

    selectionContext: (range) => {
      const key = JSON.stringify(range);
      if (cachedSelectionContext?.key === key) {
        return JSON.parse(cachedSelectionContext.json) as YrsSelectionContext;
      }
      const json = session.selection_context(
        range.story,
        range.start.paraId,
        range.start.offset,
        range.end.paraId,
        range.end.offset
      );
      const context = JSON.parse(json) as YrsSelectionContext;
      cachedSelectionContext = { key, json };
      return context;
    },
    listRevisions: () => JSON.parse(session.list_revisions()) as YrsRevisionInfo[],
    resolveComment: (commentId) =>
      JSON.parse(session.resolve_comment(commentId)) as YrsResolvedCommentAnchor[],
    storyIds: () => session.story_ids(),
    storyLength: (story) => session.story_len(story),
    storyChecksum: (story) => BigInt(session.story_checksum(story)),
    yrsBlocksForStory: (story, env = {}) => {
      const json = session.yrs_blocks_for_story(story, JSON.stringify(env));
      const blocks = JSON.parse(json) as unknown[];
      residentRenderInputs.set(story, structuredClone(env));
      return blocks;
    },
    paragraphs: (story) => JSON.parse(session.paragraphs(story)) as YrsParagraph[],
    paragraphSpans: (story) => JSON.parse(session.paragraph_spans(story)) as YrsParagraphLength[],
    storySegments: (story) => JSON.parse(session.story_segments(story)) as YrsStorySegment[],
    locateParagraph: (story, paraId) =>
      JSON.parse(session.locate_paragraph(story, paraId)) as YrsParagraphSpan,

    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      listeners.clear();
      pendingUpdates.length = 0;
      if (observing) session.clear_update_observer();
      session.free();
    },
  };
}

/**
 * Creates a yrs editing replica. The first call dynamically imports and
 * initializes the embedded docx-edit wasm (~440KB base64) — callers must
 * load it lazily so non-editor consumers avoid the wasm startup cost.
 */
export async function createYrsSession(options?: CreateYrsSessionOptions): Promise<YrsSession> {
  const clientId = options?.clientId ?? randomClientId();
  const wasm = await import('./wasm/index');
  await wasm.preloadEditWasm();
  return wrapSession(wasm.createEditSession(clientId), clientId);
}
