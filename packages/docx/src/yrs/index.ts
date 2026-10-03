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
import type { CompatibilityFlags } from '../docx/settingsParser';
import { resolveCommentMedia } from './hostMedia';
import { registerSessionInternals } from './sessionInternals';
import { editorSaveKeys } from './editorSaveKeys';
import { noteYrsStoriesDirty } from './yrsToDocument';
import type {
  DocxParagraphAnchor,
  DocxParagraphAnchorResult,
  DocxParagraphIdentityReceipt,
  DocxParagraphIdentitySnapshot,
  DocxParagraphSavePlan,
} from './paragraphIdentity';
import { decodeS9Envelope, decodeS9EnvelopeValue } from '../docx/rustParseFacade';
import { decodeEncodedSelection } from './encodedSelection';
import type {
  CollaborationCursor,
  CollaborationReplica,
  CollaborationTextInsertion,
  CollaborationUpdateOrigin,
} from '../collaboration/types';
import type {
  DocxEditRefusal,
  DocxEditRequest,
  DocxEditResult,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  DocxTextTarget,
  DocxValidationResult,
} from './edits';
import type { DocxParagraphHeading } from './readTypes';
import {
  createProposalRegistry,
  type DocxProposalRegistryState,
  type DocxProposalRequest,
  type DocxProposalResult,
  type DocxProposalSnapshot,
  type DocxProposalStateRequest,
  type DocxProposalWithdrawRequest,
} from './proposals';
import type {
  DocxContentControlQuery,
  DocxContentControlsOptions,
  DocxContentControlsResult,
} from './contentControls';
import type {
  DocxExportFailure,
  DocxExportOptions,
  DocxExportResult,
  DocxMarkdownContent,
  DocxStructuredContent,
} from './structuredExport';
import type {
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxPagedStructuredContent,
  DocxSnapshotLayoutMap,
} from './pagedExport';

export * from './edits';
export * from './contentControls';
export * from './readTypes';
export * from './structuredExport';
export * from './pagedExport';
export * from './inputPositionMap';
export {
  ResidentEngineWorkerClient,
  ResidentWorkerFailureError,
  ResidentWorkerOutOfMemoryError,
  canUseResidentEngineWorker,
  preloadResidentEngineWorker,
  retainPreloadedResidentEngineWorker,
  takePreloadedResidentEngineWorker,
  type ResidentEngineWorkerApplyResult,
  type ResidentEngineWorkerFrame,
  type ResidentEngineWorkerOpened,
  type ResidentProposalReply,
  type ResidentEngineOffscreenPage,
} from './residentEngineWorkerClient';
export { preloadDocxEngine } from './preloadDocxEngine';
export {
  residentCaretSnapshotForFrame,
  residentCaretDeviceRect,
  type ResidentCaretPaintStyle,
} from './residentCaret';
export {
  WASM32_MEMORY_LIMIT_BYTES,
  wasmModuleMemories,
  type WasmHeapStats,
  type WasmModuleMemory,
} from '../wasm/loadWasmAsset';
export { documentToYrs } from './documentToYrs';
export { ownProjectedParagraphs, yrsToDocument } from './yrsToDocument';
export * from './paragraphIdentity';
export {
  proposalRevisionPreview,
  type DocxOccurrence,
  type DocxProposalFailure,
  type DocxProposalInput,
  type DocxProposalRecord,
  type DocxProposalRegistryState,
  type DocxProposalRequest,
  type DocxProposalResult,
  type DocxProposalSnapshot,
  type DocxProposalState,
  type DocxProposalStateRequest,
  type DocxProposalWithdrawRequest,
} from './proposals';
export {
  captureSessionSave,
  saveYrsDocx,
  writeSessionSave,
  type DocxSavedDocument,
  type DocxSavedParagraph,
  type DocxSessionSave,
} from './saveYrsDocx';
export { sessionSourcePackage } from './sessionInternals';
export { editorSaveKeys } from './editorSaveKeys';
export * from './yrsPositionProjection';
export * from './proposalGeometry';

export interface YrsDocxHost {
  document: Document;
  referencedFonts: string[];
  /**
   * The `referencedFonts` a document seeded from its package names only for
   * East Asian or complex-script text it does not contain, so no text is
   * measured or drawn with them. Empty when its stories were not seeded; a
   * preview's cover only its own first pages.
   */
  unusedScriptFonts?: string[];
  /** A preview whose cut holds the whole body. */
  wholeBody?: true;
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

/** One binary Yrs position that transforms with collaborative edits. */
export interface YrsStickyPosition {
  story: string;
  encoded: Uint8Array;
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
  | { type: 'superscript' }
  | { type: 'subscript' }
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

/** How a seeding entry point starts its opening; see {@link YrsSession.beginOpening}. */
export interface YrsOpeningOptions {
  /**
   * A fixed opening generation, for a deterministic seed every replica loads
   * as one session. Each opening mints a fresh one by default.
   */
  generation?: string;
  /**
   * Seeds images as `media:{n}` tokens naming their part of the package,
   * instead of `data:` URLs, keeping them out of the document state. Only a
   * replica opened from the same package on a version that reads tokens
   * shows them, so every client of a shared room must be one. Off by default.
   */
  mediaTokens?: boolean;
}

/** SHA-256 digests of the byte copies {@link prepareDocxBytes} made, by copy. */
const preparedDigests = new WeakMap<Uint8Array, string>();

/**
 * A copy of `bytes` whose SHA-256 the platform takes off the calling thread,
 * where it has Web Crypto. Opening that copy unchanged skips hashing the
 * package on the calling thread; any other bytes open as before.
 * @internal
 */
export async function prepareDocxBytes(bytes: Uint8Array): Promise<Uint8Array> {
  const copy = new Uint8Array(bytes);
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return copy;
  try {
    const hash = new Uint8Array(await subtle.digest('SHA-256', copy));
    preparedDigests.set(
      copy,
      Array.from(hash, (byte) => byte.toString(16).padStart(2, '0')).join('')
    );
  } catch {
    // Opening hashes the copy itself.
  }
  return copy;
}

/** The SHA-256 {@link prepareDocxBytes} took of `bytes`, if it took one. @internal */
export function preparedDocxDigest(bytes: Uint8Array): string | undefined {
  return preparedDigests.get(bytes);
}

/** Snapshot of one paragraph from {@link YrsSession.paragraphs}. */
export interface YrsParagraph {
  /** Session key; not the paragraph's Word `w14:paraId`. */
  paraId: string;
  text: string;
  /** pStyle / alignment plus any op-set extras. */
  properties: Record<string, unknown>;
}

export interface YrsTextSearchOptions {
  /** Defaults to false. */
  caseSensitive?: boolean;
  /** Maximum matches; unlimited by default. */
  limit?: number;
}

/** Paragraph-local UTF-16 offsets. */
export interface YrsTextMatch {
  story: string;
  paraId: string;
  start: number;
  end: number;
  text: string;
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
  listMarkerBold?: boolean | null;
  listMarkerItalic?: boolean | null;
  listMarkerColor?: import('../types/colors').ColorValue | null;
  listMarkerSuffix?: 'tab' | 'space' | 'nothing' | null;
  listLevelNumFmts?: readonly string[] | null;
  listAbstractNumId?: number | null;
  listStartOverride?: number | null;
  /** Passive pPr/model properties not covered by the typed fields. */
  other?: Readonly<Record<string, unknown | null>>;
}

/**
 * Typed value for a content control. A string fills a plain- or rich-text control's content
 * (as a `setContentControlText` edit step would); the objects set checkbox, dropdown and date
 * controls.
 */
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
 * Where an edit landed, in whole characters: the inserted text (after the struck-out text when
 * suggesting), or what a delete left (collapsed when plain, the struck-out text when suggesting).
 */
export interface YrsReplaceReceipt extends YrsRevisionReceipt {
  range?: YrsStoryRange;
}

/**
 * Receipt of {@link YrsSession.splitParagraph}. The first half keeps the
 * original paraId and the second half is re-minted; suggesting
 * mode stamps a `pPrIns` revision on the first half.
 */
export interface YrsSplitReceipt {
  firstParaId: string;
  secondParaId: string;
  revisionId: string | null;
}

/** Low-level UTF-16 story operation for {@link YrsSession.applyRawOps}. */
export type YrsRawOp =
  | { op: 'insert'; index: number; text: string; attrs?: Record<string, unknown> }
  | { op: 'delete'; index: number; len: number }
  | { op: 'format'; index: number; len: number; attrs?: Record<string, unknown> }
  | {
      /**
       * A pilcrow's `ooxmlParaId` binds its Word paragraph ID and is dropped
       * when it is not one; `sourceParaId` and `paraOrigin` are seeding's to set.
       */
      op: 'insertEmbed';
      index: number;
      kind: string;
      payload?: Record<string, unknown>;
      attrs?: Record<string, unknown>;
    }
  | { op: 'setEmbedAttr'; index: number; key: string; value: unknown }
  | {
      /** Upserts a side-map comment with sticky story-unit ranges. */
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
  compatibilityFlags?: Partial<CompatibilityFlags>;
  tocStyleIds?: string[];
  paragraphSpacingLinePx?: number;
  /** Section document-grid snap pitch in px (w:docGrid). The engine derives this from sections; hosts may omit it. */
  docGridPitchPx?: number;
  defaultParagraphStyleId?: string;
  /** Theme color name → hex (`accent1` → `4472C4`), for theme-color resolution. */
  themeColors?: Record<string, string>;
  /** The document default tab stop in twips. */
  defaultTabStopTwips?: number | null;
  /** Current page content height in CSS px, for proportional drawing constraints. */
  pageContentHeight?: number | null;
  /** yrs revision/paragraph id → dense numeric layout id (list markers, revisions). */
  numericIds?: Record<string, number>;
  /** Include hidden text in visible layout without changing the document. */
  showHiddenText?: boolean;
  /** Revision id → decision shown in layout; unlisted revisions render as tracked changes. */
  revisionPreview?: Readonly<Record<string, 'accepted' | 'rejected'>>;
  /**
   * Lay the opened package's images out as `media:{n}` tokens, which a
   * resolver given the session's `mediaSource` reads. @internal
   */
  mediaTokens?: boolean;
}

/** Receipt of {@link YrsSession.addComment}. */
export interface YrsCommentReceipt {
  commentId: string;
}

/** A comment the session holds; see {@link YrsSession.listComments}. @internal */
export interface YrsCommentInfo {
  id: string;
  author: string;
  date: string;
  done: boolean;
  parentId: string | null;
  /** The body as it was given to {@link YrsSession.addComment} or `setComment`. */
  body: unknown;
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

/** How far a region layout begun with {@link YrsSession.beginRegionLayout} has come. */
export interface YrsRegionLayoutProgress {
  measuredBlocks: number;
  bodyBlocks: number;
  /** The retained region layout reply, once the pass is complete. */
  layoutJson?: string;
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
/**
 * One resident font-store registration, in the order the main thread made it:
 * raw bytes, or a measurement view of an earlier id. The worker replays the
 * list to reproduce the same dense ids, so the order and the mix both matter.
 *
 * @internal
 */
export type YrsResidentFontRegistration =
  | Uint8Array
  | { substituteOf: number; family: string };

export interface YrsResidentWorkerSnapshot {
  /** @internal */
  workerAuthoritative?: true;
  clientId: number;
  /** Full document state, or a state-vector diff when the caller supplied
   * the worker's known vector — both apply through the same merge path. */
  state: Uint8Array;
  selection: YrsSelection | null;
  /** Empty when the caller declared the worker's fonts current
   * (`knownFontsRevision` matches); the worker then keeps its registrations. */
  fonts: YrsResidentFontRegistration[];
  /** Monotonic revision of the resident font set (bumped by register/clear). */
  fontsRevision: number;
  renderInputs: Array<{ story: string; env: YrsRenderEnv }>;
  measureInputs: string[];
  layoutInput: string;
  layoutWithRegions: boolean;
  layoutRevision: number;
  /** The document is a preview's cut of a package: its layouts render NUMPAGES empty. */
  partialDocument?: boolean;
  /** Which seeded `data:` image sources lay out as `media:{n}` tokens. @internal */
  mediaSources?: string;
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
  superscript: YrsTriState;
  subscript: YrsTriState;
  /** Uniform ASCII font family, or `null` when absent/mixed. */
  fontFamily: string | null;
  /** Uniform font size in half-points (the OOXML `w:sz` unit). */
  fontSize: number | null;
  /** Uniform RGB hex or theme-color name, or `null` when absent/mixed. */
  color: string | null;
  /** Uniform highlight name (`yellow`, …) or unmapped hex, or `null` when absent/mixed. */
  highlight: string | null;
  /** Paragraph containing the range start. */
  paraId: string;
  styleId: string | null;
  alignment: string | null;
  /**
   * Full authored pilcrow property bag. Known toolbar fields retain their
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
 * Anchor-cell to head-cell rectangular selection, addressed by grid
 * coordinates rather than document positions.
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
  /** Existing stories whose content changed: the table's story, plus a merge's surviving cell. */
  changedStoryIds: string[];
  newParaIds: string[];
  deletedTable: boolean;
  revisionIds: string[];
}

/**
 * OOXML-shaped cell border value passed through to `tcPr.borders`.
 *
 * @public
 */
export interface YrsCellBorder {
  style: string;
  size?: number;
  color?: { rgb: string };
}

/**
 * Per-side border patch for {@link YrsSession.setCellBorders}. An omitted side
 * is left alone, `style: 'none'` authors an explicit no-border, and `null`
 * drops the authored side.
 *
 * @public
 */
export type YrsCellBorders = Partial<
  Record<'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV', YrsCellBorder | null>
>;

/** Undo grouping for tracked local transactions. */
export type YrsUndoCaptureMode = 'auto' | 'manual';

/** Outcome of a target-resolving helper edit. @internal */
export type YrsTargetEditResult = { ok: true; version: string } | DocxEditRefusal;

/** Accepted-view texts around a selection; U+FFFC stands for each inline atom. @internal */
export interface YrsSelectionText {
  paraId: string;
  selectedText: string;
  paragraphText: string;
  before: string;
  after: string;
}

/**
 * One live replica of the yrs editing model. Thin typed wrapper over the
 * wasm `EditSession` — no editing logic on this side of the boundary.
 */
/** An embedded image's bytes, as its `media:{n}` source displays them. */
export interface YrsMediaSource {
  bytes: Uint8Array;
  mimeType: string;
}

/** @internal */
export interface YrsWorkerDocumentMirror {
  version: string;
  proposals: DocxProposalRegistryState;
}

export interface YrsSession extends CollaborationReplica {
  /** The yrs client id this replica writes with. */
  readonly clientId: number;

  // -- resident layout engine (same wasm instance as EditingDoc) --

  /** Register raw sfnt bytes in the session's measurement/display font store. */
  registerFont(bytes: Uint8Array): number;
  /**
   * Register a measurement view of `base` carrying the vertical metrics Word
   * measures `family` with, for a face the host substituted; `base` when the
   * engine knows no metrics for the family.
   */
  registerSubstituteFont(base: number, family: string): number;
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
  /** Same pass, but the reply omits the measured arena (fetch it on demand
   * through {@link YrsSession.retainedKernelInputsJson}). */
  layoutDocumentWithRegionsRetainedJson(input: string): string;
  /**
   * {@link YrsSession.layoutDocumentWithRegionsRetainedJson} a step at a time:
   * this lowers the body, and each {@link YrsSession.resumeRegionLayout}
   * measures up to `blocks` more body blocks. `layoutJson` arrives with the
   * step that completes the pass, equal to the one-call reply. A document
   * change, a font registration or another layout in between abandons the
   * pass, and resuming it then throws.
   */
  beginRegionLayout(input: string): YrsRegionLayoutProgress;
  resumeRegionLayout(blocks: number): YrsRegionLayoutProgress;
  /** Retained `{ measured, options }` for the main-thread display fallback. */
  retainedKernelInputsJson(expectedLayoutRevision: number): string;
  /**
   * Record region layout `input` as the resident layout without running it
   * here: a resident worker replica runs it, and snapshots carry it there.
   * Returns the new layout revision. @internal
   */
  adoptResidentWorkerLayout?(input: string): number;
  /** The current resident layout ran only in a worker replica. @internal */
  residentLayoutInWorker?(): boolean;
  /** Build display primitives against the session's resident font store. */
  buildDisplayListJson(input: string): string;
  /** Build a binary FrameDelta v1 against the last host-applied frame. */
  buildDisplayListFrame(input: string, expectedFrameEpoch: number): Uint8Array;
  /**
   * Limit full display builds to pages `start..end` plus the pages already
   * built; the rest stay unbuilt placeholders carrying their geometry. @internal
   */
  setDisplayWindow(start: number, end: number): void;
  /** Keep every previously built page while windowed builds are on. @internal */
  setDisplayRetainBuiltPages(retain: boolean): void;
  /** @internal */
  setWindowedIncrementalBuilds(enabled: boolean): void;
  /** Build the listed unbuilt pages into a FrameDelta v1. @internal */
  buildDisplayPagesFrame(pages: readonly number[], expectedFrameEpoch: number): Uint8Array;
  /** Release built pages; null means the request was superseded. @internal */
  releaseDisplayPagesFrame(pages: number[], expectedFrameEpoch: number): Uint8Array | null;
  /** Make the next frame a full one, for a host taking over from another engine; no-op once destroyed. */
  resetFrameBase(): void;
  /** Caret geometry from the current resident display frame. */
  residentCaretSnapshot(): YrsResidentCaretSnapshot;
  /** Apply a collapsed plain-text insertion and return its resident FrameDelta. */
  applyInput(text: string, expectedFrameEpoch: number): Uint8Array;
  /**
   * Apply up to `count` (default 1) collapsed character deletions/paragraph
   * merges, lay out once, and return the resident FrameDelta.
   */
  applyDelete(direction: 'backward' | 'forward', expectedFrameEpoch: number, count?: number): Uint8Array;
  /** Characters the last resident deletion removed; fewer than asked at a document boundary. */
  residentDeletedUnits(): number;
  /** Instrumented apply used only by opt-in browser performance traces. */
  applyInputProfiled(
    text: string,
    expectedFrameEpoch: number
  ): { frame: Uint8Array; profile: YrsEngineApplyProfile };
  /** Instrumented deletion used only by opt-in browser performance traces. */
  applyDeleteProfiled(
    direction: 'backward' | 'forward',
    expectedFrameEpoch: number,
    count?: number
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
  displayVerticalMoveJson(
    position: number,
    direction: 'up' | 'down',
    goalX: number
  ): string;
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
  /** Parses a DOCX, seeds its stories, and returns thin host metadata; see {@link openDocx}. */
  seedFromDocx(bytes: Uint8Array, options?: YrsOpeningOptions): YrsDocxHost;
  /**
   * Parses a DOCX and optionally seeds its stories. Seeding starts a new
   * opening, so its session anchors never resolve in another opening, even
   * one seeded alike by the same client; see {@link beginOpening}.
   */
  openDocx(bytes: Uint8Array, seedStories: boolean, options?: YrsOpeningOptions): YrsDocxHost;
  /**
   * Opens a DOCX for display only, from the body's first `blocks` blocks:
   * enough to lay out its first pages with a prefix pass before the whole
   * document is opened. The session cannot save. `null`, opening nothing,
   * for a document with a float placed from outside the text or a section
   * with columns, which no cut of the body lays out like the whole: open it
   * with {@link openDocx}. @internal
   */
  openDocxPreview(bytes: Uint8Array, blocks: number): YrsDocxHost | null;
  /** Opened by {@link openDocxPreview}: its document refuses every change. @internal */
  isDisplayOnly(): boolean;
  /**
   * Marks this empty session as showing a preview another engine opened: it refuses every
   * change, and {@link openDocxPreview} can still load the preview here. @internal
   */
  markDisplayOnly(): void;
  /**
   * Marks whether the document is a preview's, as a replica of one is: its
   * layouts render NUMPAGES empty. @internal
   */
  setPartialDocument(partial: boolean): void;
  /**
   * The region layout of only as much of the body as fills `pages` pages;
   * a reply marked `provisional` covers a prefix. @internal
   */
  layoutDocumentWithRegionsPrefixRetainedJson(input: string, pages: number): string;
  /**
   * Starts a new opening of the document: its generation, replicated to
   * every replica, becomes part of every session anchor. Every seeding entry
   * point calls it; call it after building a document another way.
   */
  beginOpening(generation?: string): void;
  /** Unions seeded opaque sequence names into document state. @internal */
  seedOpaqueSequences(names: readonly string[]): void;
  /** Materializes the retained canonical package for compatibility APIs. */
  materializeDocx(): Document | null;
  /**
   * The displayed bytes and media type of the package part a `media:{n}`
   * image source names, or `null` for any other source or a destroyed session.
   */
  mediaSource(token: string): YrsMediaSource | null;
  /** Changes on package opening and session destruction. */
  mediaScope(): number;
  /** The `data:` URL a token stands for, or `null` when unavailable or destroyed. */
  mediaDataUrl(token: string): string | null;
  /**
   * Lays this replica's `data:` image sources out as the `media:{n}` tokens a
   * snapshot's `mediaSources` names. @internal
   */
  loadMediaSources(json: string): void;
  /**
   * Seeds stories and returns paragraph IDs in document order. Seeding a
   * document that has no opening yet starts one; see {@link beginOpening}.
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
  applyLocalUpdate(update: Uint8Array): void;
  /** Adopt another replica's host batch outside undo history. @internal */
  applyHostUpdate(update: Uint8Array, stories?: readonly string[]): void;
  /**
   * Subscribes to every committed transaction's v1 update (local AND
   * applied-remote). Returns an unsubscribe function.
   */
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;

  // -- local input state --

  /** Encode one location as a binary Yrs sticky position. */
  encodeStickyPosition(loc: YrsLoc): YrsStickyPosition;
  /** Resolve a binary Yrs sticky position against this replica. */
  resolveStickyPosition(position: YrsStickyPosition): YrsLoc | null;
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
  /** Begin local-origin undo capture once import/seeding has completed. */
  beginUndoCapture(): void;
  /** Separates subsequent local edits from the current undo step; safe before capture starts. */
  addUndoBoundary(): void;
  /** Changes grouping policy, closing the current group while retaining history. */
  setUndoCaptureMode(mode: YrsUndoCaptureMode): void;
  /** Current grouping policy; defaults to auto. */
  undoCaptureMode(): YrsUndoCaptureMode;
  /** Stories changed by the latest undo or redo, sorted. */
  historyStories(): string[];
  /** Undo/redo only local-origin direct operations (never remote/system transactions). */
  undo(): boolean;
  redo(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;

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
   * Merges an OOXML-shaped patch into selected cells' `tcPr`. JSON `null`
   * clears a property; merge/split-owned span keys are rejected.
   */
  setCellTextFormat(
    range: YrsTableRange,
    patch: Readonly<Record<string, unknown>>
  ): YrsTableReceipt;
  /**
   * Merges border sides into the selected cells. `insideH`/`insideV` resolve
   * per cell to the physical edges interior to the selection.
   */
  setCellBorders(range: YrsTableRange, borders: YrsCellBorders): YrsTableReceipt;
  /** Sets one authored grid-column width in twips. */
  setColumnWidth(at: YrsCellLoc, widthTwips: number): YrsTableReceipt;
  /** Sets the table-wide preferred width in twips. */
  setTableWidth(table: YrsTableLoc, widthTwips: number): YrsTableReceipt;
  /** Inserts paragraph-break-free text. Suggesting mode mints a revision. */
  insertText(at: YrsLoc, text: string, suggesting?: YrsAuthor): YrsReplaceReceipt;
  /**
   * Deletes a range (plain) or marks it as a suggested deletion (suggesting).
   * A range spanning paragraphs also merges them (pilcrow-as-character).
   */
  deleteRange(range: YrsStoryRange, suggesting?: YrsAuthor): YrsReplaceReceipt;
  /** Replaces a range with text in one transaction (one shared revision when suggesting). */
  replaceRange(range: YrsStoryRange, text: string, suggesting?: YrsAuthor): YrsReplaceReceipt;
  /**
   * Splits a paragraph by inserting one pilcrow. The FIRST half keeps the
   * original paraId; the SECOND half is re-minted (`secondParaId`).
   */
  splitParagraph(at: YrsLoc, suggesting?: YrsAuthor): YrsSplitReceipt;
  /** Merges `paraId` with the FOLLOWING paragraph. Errors on the final paragraph. */
  mergeParagraphs(story: string, paraId: string, suggesting?: YrsAuthor): YrsRevisionReceipt;
  /**
   * Toggles one run mark across a range: removes it when every unit already
   * carries it, otherwise adds it. Adding superscript clears subscript and
   * vice versa.
   */
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
  ): YrsReplaceReceipt;
  /**
   * Sets the value of a content-control embed addressed by stable payload id. A string fills a
   * text control's content as one version-checked step and throws when the fill is refused.
   */
  setContentControlValue(embedId: string, value: YrsContentControlValue): void;
  /** {@link setContentControlValue} for the embed at a paragraph-keyed position. */
  setContentControlValueAt(at: YrsLoc, value: YrsContentControlValue): void;
  /**
   * Removes an authored value from a content-control embed. It never erases a text control's
   * text: fill it with `''` for that.
   */
  clearContentControlValue(embedId: string): void;
  /** Commits image size/wrapping/position fields in one transaction. */
  setImageGeometry(embedId: string, geometry: YrsImageGeometry): void;
  /** Commits image geometry at a paragraph-keyed position; reaches images that share an id. */
  setImageGeometryAt(at: YrsLoc, geometry: YrsImageGeometry): void;
  /** Inserts a native page-break embed at a paragraph-keyed location. */
  insertPageBreak(at: YrsLoc): void;
  /** Inserts a native section-break embed at a paragraph-keyed location. */
  insertSectionBreak(at: YrsLoc, type: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage'): void;
  /** Inserts a typed watermark embed at a paragraph-keyed location. */
  insertWatermark(at: YrsLoc, watermark: YrsWatermark): void;
  /**
   * Applies raw story operations in one transaction, then repairs any
   * paragraph identity a pilcrow they insert or re-key duplicates and
   * promotes an editor-only paragraph they author into.
   */
  applyRawOps(story: string, ops: readonly YrsRawOp[]): void;
  /**
   * Applies seed raw operations with deterministic item ordering. A
   * `setComment` for a comment already seeded adds its ranges to it.
   */
  applySeedRawOps(story: string, ops: readonly YrsRawOp[]): void;
  /** Sets one paragraph property (any JSON value). `paraId` is reserved. */
  setParagraphAttr(paraId: string, key: string, value: unknown): void;
  /** Adds a sticky-anchored comment over one or more ranges. */
  addComment(
    ranges: readonly YrsStoryRange[],
    author: string,
    date: string,
    body: unknown
  ): YrsCommentReceipt;
  /** Reanchors an existing comment with non-empty ranges; preserves metadata and joins local undo capture. */
  setCommentRanges(commentId: string, ranges: readonly YrsStoryRange[]): void;
  /**
   * Accepts tracked changes: pending insertions become plain content,
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
  /** Every comment the session holds, sorted by id. @internal */
  listComments(): YrsCommentInfo[];
  /** Story ids in the document, sorted. */
  storyIds(): string[];
  /** Whether the document has a story with this id, without listing them all. */
  hasStory(story: string): boolean;
  /** Story length in UTF-16 units (every embed, pilcrows included, counts 1). */
  storyLength(story: string): number;
  /** Returns the story's canonical-stream checksum. */
  storyChecksum(story: string): bigint;
  /** Lowers a story to layout blocks and rejects unsupported embeds. */
  yrsBlocksForStory(story: string, env?: YrsRenderEnv): unknown[];
  /** Paragraph snapshots in document order. */
  paragraphs(story: string): YrsParagraph[];
  /** Literal search in document order. */
  searchText(query: string, options?: YrsTextSearchOptions): YrsTextMatch[];
  /** Paragraph ids and inline-unit lengths, resolved in one Rust story traversal. */
  paragraphSpans(story: string): YrsParagraphLength[];
  /** The raw formatted-segment view (the render bridge's input). */
  storySegments(story: string): YrsStorySegment[];
  /**
   * The current story revision and the sorted ids of the stories created,
   * edited, or deleted after revision `since` (0 lists every story).
   */
  storiesChangedSince(since: number): { revision: number; stories: string[] };
  /**
   * One digest per unit of {@link YrsSession.storySegments}, split after each
   * pilcrow. Equal digests mean equal segments.
   */
  storySegmentUnitDigests(story: string): string[];
  /** The segments of the listed units, each as {@link YrsSession.storySegments} gives them. */
  storySegmentUnits(story: string, units: readonly number[]): YrsStorySegment[][];
  /**
   * The payload of the story's `tableIndex`-th table embed, as
   * {@link YrsSession.storySegments} gives it, or null when there is no such
   * table. Reads the one table rather than the whole story.
   */
  tablePayload(story: string, tableIndex: number): Record<string, unknown> | null;
  /** A paragraph's story span (start unit, pilcrow index). */
  locateParagraph(story: string, paraId: string): YrsParagraphSpan;
  /** How many paragraphs of `story` carry `paraId`: 0, 1, or 2 for two or more. */
  paragraphIdCount(story: string, paraId: string): number;

  // -- paragraph identity --

  /**
   * Session, persisted and source anchors of every paragraph, including the
   * source package's paragraphs outside the stories, and the Word paragraph ID
   * each saves with. Reads only.
   */
  paragraphIdentities(): DocxParagraphIdentitySnapshot;
  /**
   * Gives every paragraph that saves without a Word paragraph ID a fresh one,
   * editor-only paragraphs excepted, and repairs duplicates, so persisted
   * anchors survive save and reopen. Source IDs and saved claims keep their
   * IDs over unsaved claims and copies. A replicated change outside undo
   * history that later saves keep; a refusal changes nothing.
   */
  persistParagraphIds(): DocxParagraphIdentityReceipt;
  /** Resolves an anchor to the paragraph that currently holds it. Reads only. */
  resolveParagraphAnchor(anchor: DocxParagraphAnchor): DocxParagraphAnchorResult;
  /** The Word paragraph ID each paragraph of a story saves with, in pilcrow order. @internal */
  storyParagraphIds(story: string): Array<string | null>;
  /** The paragraph IDs a save of the session applies. @internal */
  paragraphSavePlan(): DocxParagraphSavePlan;
  /**
   * Publishes the `[owner, paraId]` pairs a save captured and returns those
   * whose binding changed since. @internal
   */
  recordSavedParagraphIds(
    saved: ReadonlyArray<readonly [string, string]>
  ): Array<[string, string]>;
  /**
   * The Word paragraph IDs a DOCX package holds, by part URI, in document
   * order and upper case. Reads only. @internal
   */
  writtenParagraphIds(bytes: Uint8Array): Record<string, string[]>;

  // -- version-checked host edits --

  /**
   * The session-scoped version token. It changes on every committed change, local or remote,
   * and when the document is replaced; compare tokens only within this session.
   */
  version(): string;
  /** Paragraph texts in one view, with the version they were read at. */
  readParagraphs(request: DocxReadParagraphsRequest): DocxReadParagraphsResult;
  /** Exact, case-sensitive, paragraph-local search; overlapping matches count separately. */
  findText(request: DocxFindTextRequest): DocxFindTextResult;
  /** Resolves and checks an edit batch without changing anything or reserving ids. */
  validateEdits(request: DocxEditRequest): DocxValidationResult;
  /**
   * Applies every step or none, resolving all targets against `expectVersion`. Policy
   * failures are returned; malformed requests throw.
   */
  applyEdits(request: DocxEditRequest): DocxEditResult;
  /** Resolves a text target and formats it in one call. @internal */
  formatTextTarget(target: DocxTextTarget, delta: YrsInlineFormatDelta): YrsTargetEditResult;
  /** Resolves a text target and anchors side-map comment `id` over it in one call. @internal */
  commentTextTarget(
    target: DocxTextTarget,
    comment: { id: string; author: string; date: string; body: unknown }
  ): YrsTargetEditResult;
  /** Accepted-view texts around a paragraph-keyed selection. @internal */
  selectionText(range: YrsStoryRange): YrsSelectionText;

  // -- host proposals --

  /**
   * Proposes a round of tracked changes: every new proposal resolves against `expectVersion` and
   * the round applies as one batch outside undo history, or nothing changes. A retried id with
   * the same edit is a no-op; the same id with another edit refuses. Opening another document
   * forgets every proposal.
   */
  proposeChanges(request: DocxProposalRequest): DocxProposalResult;
  /**
   * Sets how proposals render. Decisions change neither the document, its version nor undo
   * history; each call that changes one increments `previewVersion`.
   */
  setProposalStates(request: DocxProposalStateRequest): DocxProposalResult;
  /**
   * Withdraws proposals, settling each as its decision previews it: accepted ones apply for good,
   * rejected and undecided ones are removed. The settlement is one change against `expectVersion`
   * outside undo history, so a later round resolves against the text the preview showed.
   */
  withdrawProposals(request: DocxProposalWithdrawRequest): DocxProposalResult;
  /** The proposals in the order they were made. */
  getProposals(): DocxProposalSnapshot;
  /** @internal */
  mirrorWorkerDocument(mirror: YrsWorkerDocumentMirror | null): void;
  /** @internal */
  workerDocumentMirrored(): boolean;
  /** Listens for new proposals, decisions and a forgotten registry. Returns the unsubscribe. */
  onProposalChange(listener: (snapshot: DocxProposalSnapshot) => void): () => void;

  // -- structured export --

  /**
   * Exports the committed document as read-only structured content with the version it was read
   * at; anchors resolve against that version. Nothing is committed, flushed or published.
   * Unusable options are refused; malformed ones throw.
   */
  exportStructured(options: DocxExportOptions): DocxExportResult<DocxStructuredContent>;
  /** {@link exportStructured} rendered as Markdown from the same read. */
  exportMarkdown(options: DocxExportOptions): DocxExportResult<DocxMarkdownContent>;
  /**
   * {@link exportStructured} with the page map of the region layout this session retains. The
   * layout must have lowered the current version from this session's stories with section,
   * settings and note metadata that describe it, and measured every font the document uses
   * with the fonts registered now; this call lays nothing out, flushes nothing and loads no
   * font, and refuses a stale or incomplete layout.
   */
  exportStructuredWithPages(
    options: DocxPageExportOptions
  ): DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;
  /**
   * {@link exportStructuredWithPages} for an editor: `currentRequest` is the region layout
   * request it would lay the document out with now, and the retained layout must have used the
   * same fonts, measurement defaults, render environment and pagination options. @internal
   */
  exportStructuredWithPagesFor(
    options: DocxPageExportOptions,
    currentRequest: string
  ): DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;
  /**
   * Lays this private session out with `fonts` alone, in a measurement font store of its own,
   * and exports it with snapshot anchors and a snapshot map. `fonts` holds the font files back
   * to back and `fontLengths` their lengths; `request`'s font chains name them by index. Throws
   * for a font the engine rejects. @internal
   */
  exportSnapshotWithPrivateFonts(
    fonts: Uint8Array,
    fontLengths: Uint32Array,
    request: string,
    options: DocxPageExportOptions
  ):
    | { ok: true; content: DocxPagedStructuredContent<DocxSnapshotLayoutMap> }
    | { ok: false; failure: DocxExportFailure };
  /**
   * The heading paragraphs of `story` in document order, classified as the structured export
   * classifies them. Throws for an unknown story.
   */
  headings(story: string): DocxParagraphHeading[];

  // -- content controls --

  /**
   * The content controls of the committed document with the version they were read at; control
   * ids and anchors resolve against that version. Nothing is committed, flushed or published.
   * Unusable options are refused; malformed ones throw.
   */
  listContentControls(options?: DocxContentControlsOptions): DocxContentControlsResult;
  /** The controls {@link listContentControls} lists that match `query` exactly; none or several. */
  findContentControls(
    query: DocxContentControlQuery,
    options?: DocxContentControlsOptions
  ): DocxContentControlsResult;

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

function targetStory(target: DocxTextTarget): string {
  return target.kind === 'search' ? target.within.story : target.story;
}

function docxSourceBuffer(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === bytes.buffer.byteLength
  ) {
    return bytes.buffer;
  }
  return new Uint8Array(bytes).buffer as ArrayBuffer;
}

/**
 * Decodes the host metadata a resident worker's `open` replied with, for the
 * package `source` it opened. @internal
 */
export function decodeDocxHostJson(json: string, source: Uint8Array): YrsDocxHost {
  return decodeDocxHost(json, source);
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
  const unusedScriptFonts = wire.unusedScriptFonts ?? [];
  if (
    !Array.isArray(unusedScriptFonts) ||
    !unusedScriptFonts.every((name) => typeof name === 'string')
  ) {
    throw new TypeError('DOCX host unusedScriptFonts must be a string array');
  }
  const result = decodeS9EnvelopeValue(wire.envelope, docxSourceBuffer(source));
  return {
    document: result.document,
    referencedFonts: wire.referencedFonts,
    unusedScriptFonts,
    ...(wire.wholeBody === true ? { wholeBody: true as const } : {}),
    embeddedFonts: result.embeddedFonts,
    ...(result.fontTableRelationshipsXml === undefined
      ? {}
      : { fontTableRelationshipsXml: result.fontTableRelationshipsXml }),
  };
}

/** A lone UTF-16 surrogate, which crossing into Wasm would turn into U+FFFD. */
const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

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
  let undoTracked = false;
  let cachedSelection: YrsSelection | null | undefined;
  let cachedSelectionContext: { key: string; json: string } | null = null;
  const residentFonts: YrsResidentFontRegistration[] = [];
  const residentRenderInputs = new Map<string, YrsRenderEnv>();
  const residentMeasureInputs = new Map<string, string>();
  let residentLayoutInput: string | null = null;
  let residentLayoutWithRegions = false;
  let residentLayoutRevision = 0;
  // The current resident layout ran only in a resident worker replica.
  let layoutRanInWorker = false;
  // The request of a region layout begun a step at a time.
  let steppedLayoutInput: string | null = null;
  const completedRegionLayout = (progress: YrsRegionLayoutProgress): YrsRegionLayoutProgress => {
    if (progress.layoutJson !== undefined && steppedLayoutInput !== null) {
      residentLayoutInput = steppedLayoutInput;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      steppedLayoutInput = null;
    }
    return progress;
  };
  let residentFontsRevision = 0;
  let docxSource: Uint8Array | null = null;
  let docxSourceKeys: ReturnType<typeof editorSaveKeys> | null = null;

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

  // A preview session refuses every change to its document.
  let displayOnly = false;
  // A preview's cut of a package, whose layouts count only its own pages.
  let partialDocument = false;
  const mutate = <T>(operation: () => T): T => {
    if (displayOnly) throw new Error('A document preview is display-only');
    return mutateAlways(operation);
  };
  const mutateAlways = <T>(operation: () => T): T => {
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

  const ensureUndo = (targetStory?: string): void => {
    if (!undoTracked) {
      session.track_undo();
      undoTracked = true;
    }
    if (targetStory !== undefined) {
      session.select_story(targetStory);
      markDirty(targetStory);
    }
  };

  const markDirty = (stories: 'all' | string | Iterable<string>): void => {
    noteYrsStoriesDirty(facade, stories);
  };

  const markReceiptStories = (receipt: YrsTableReceipt): YrsTableReceipt => {
    markDirty(receipt.createdStoryIds);
    markDirty(receipt.deletedStoryIds);
    markDirty(receipt.changedStoryIds);
    return receipt;
  };

  const selectionStory = (): 'all' | string =>
    (
      (cachedSelection !== undefined
        ? cachedSelection
        : (JSON.parse(session.selection()) as YrsSelection | null))?.head.story ?? 'all'
    );

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

  // `data:` URLs of the opened package's `media:{n}` sources.
  const mediaDataUrls = new Map<string, string | null>();
  let mediaScope = 0;
  const resetMedia = (): void => {
    mediaDataUrls.clear();
    mediaScope += 1;
  };

  const openDocx = (
    bytes: Uint8Array,
    seedStories: boolean,
    options: YrsOpeningOptions = {}
  ): YrsDocxHost => {
    const source = bytes.slice();
    markDirty('all');
    resetMedia();
    const json = mutate(() => {
      session.set_media_tokens(options.mediaTokens === true);
      const opened = session.open_docx(
        source,
        seedStories,
        options.generation,
        preparedDigests.get(bytes)
      );
      proposals.reset();
      return opened;
    });
    const host = withHostMedia(decodeDocxHost(json, source));
    docxSource = source;
    docxSourceKeys = editorSaveKeys(host.document);
    partialDocument = false;
    return host;
  };

  const mediaDataUrl = (token: string): string | null => {
    if (destroyed || !token.startsWith('media:')) return null;
    let url = mediaDataUrls.get(token);
    if (url === undefined) {
      url = session.media_data_url(token) ?? null;
      mediaDataUrls.set(token, url);
    }
    return url;
  };

  const withHostMedia = (host: YrsDocxHost): YrsDocxHost => {
    resolveCommentMedia(host.document.package.document.comments, mediaDataUrl);
    return host;
  };

  let workerDocumentVersion: string | null = null;
  const proposals = createProposalRegistry({
    version: () => session.version(),
    resolveParagraphAnchor: (anchor) => facade.resolveParagraphAnchor(anchor),
    findText: (request) => facade.findText(request),
    readParagraphs: (request) => facade.readParagraphs(request),
    applyEdits: (request) => facade.applyEdits(request),
    listRevisions: () => facade.listRevisions(),
    revisionStamps: (ids) => JSON.parse(session.revision_stamps_json(JSON.stringify(ids))),
    settleRevisions: (accept, reject) => {
      const since = facade.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      session.settle_revisions_json(JSON.stringify({ accept, reject }));
      markDirty(facade.storiesChangedSince(since).stories);
    },
    ...(typeof session.begin_shared_reads === 'function' &&
    typeof session.end_shared_reads === 'function'
      ? {
          sharedReads: <R>(read: () => R): R => {
            session.begin_shared_reads();
            try {
              return read();
            } finally {
              session.end_shared_reads();
            }
          },
        }
      : {}),
  });

  const facade: YrsSession = {
    clientId,
    openDocxPreview: (bytes, blocks) => {
      markDirty('all');
      resetMedia();
      const json = mutateAlways(() => session.open_docx_preview(bytes, blocks));
      if (json === undefined) return null;
      displayOnly = true;
      partialDocument = true;
      return withHostMedia(decodeDocxHost(json, bytes));
    },
    isDisplayOnly: () => displayOnly,
    markDisplayOnly: () => {
      displayOnly = true;
      partialDocument = true;
    },
    setPartialDocument: (partial) => {
      partialDocument = partial;
      session.set_partial_document(partial);
    },
    layoutDocumentWithRegionsPrefixRetainedJson: (input, pages) => {
      const output = session.layout_document_with_regions_prefix_retained_json(input, pages);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },

    registerFont: (bytes) => {
      const id = session.register_measure_font(bytes);
      residentFonts.push(bytes.slice());
      residentFontsRevision += 1;
      return id;
    },
    registerSubstituteFont: (base, family) => {
      const id = session.register_substitute_measure_font(base, family);
      if (id === base) return id;
      residentFonts.push({ substituteOf: base, family });
      residentFontsRevision += 1;
      return id;
    },
    clearFonts: () => {
      session.clear_measure_fonts();
      residentFonts.length = 0;
      residentMeasureInputs.clear();
      residentFontsRevision += 1;
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
      layoutRanInWorker = false;
      return output;
    },
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsJson: (input) => {
      const output = session.layout_document_with_regions_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    layoutDocumentWithRegionsRetainedJson: (input) => {
      const output = session.layout_document_with_regions_retained_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    beginRegionLayout: (input) => {
      steppedLayoutInput = input;
      return completedRegionLayout(
        JSON.parse(session.begin_region_layout(input)) as YrsRegionLayoutProgress
      );
    },
    resumeRegionLayout: (blocks) =>
      completedRegionLayout(
        JSON.parse(session.resume_region_layout(blocks)) as YrsRegionLayoutProgress
      ),
    adoptResidentWorkerLayout: (input) => {
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = true;
      return residentLayoutRevision;
    },
    residentLayoutInWorker: () => layoutRanInWorker,
    retainedKernelInputsJson: (expectedLayoutRevision) => {
      if (layoutRanInWorker) {
        throw new Error('the retained layout was computed in the resident worker');
      }
      if (expectedLayoutRevision !== residentLayoutRevision) {
        throw new Error(
          `retained layout revision mismatch: expected ${expectedLayoutRevision}, current ${residentLayoutRevision}`
        );
      }
      return session.retained_kernel_inputs_json();
    },
    buildDisplayListJson: (input) => session.build_display_list_json(input),
    resetFrameBase: () => {
      if (!destroyed) session.reset_frame_base();
    },
    buildDisplayListFrame: (input, expectedFrameEpoch) =>
      session.build_display_list_frame(input, expectedFrameEpoch),
    setDisplayWindow: (start, end) => session.set_display_window(start, end),
    setDisplayRetainBuiltPages: (retain) => session.set_display_retain_built_pages(retain),
    setWindowedIncrementalBuilds: (enabled) => session.set_windowed_incremental_builds(enabled),
    buildDisplayPagesFrame: (pages, expectedFrameEpoch) =>
      session.build_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch),
    releaseDisplayPagesFrame: (pages, expectedFrameEpoch) => {
      const frame = session.release_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch);
      return frame.length === 0 ? null : frame;
    },
    residentCaretSnapshot: () =>
      JSON.parse(session.resident_caret_snapshot_json()) as YrsResidentCaretSnapshot,
    applyInput: (text, expectedFrameEpoch) => {
      ensureUndo();
      markDirty(selectionStory());
      return mutate(() => session.apply_input(text, expectedFrameEpoch));
    },
    applyDelete: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      markDirty(selectionStory());
      return mutate(() => session.apply_delete(direction, expectedFrameEpoch, count));
    },
    residentDeletedUnits: () => session.resident_deleted_units(),
    applyInputProfiled: (text, expectedFrameEpoch) => {
      ensureUndo();
      markDirty(selectionStory());
      const frame = mutate(() => session.apply_input_profiled(text, expectedFrameEpoch));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      markDirty(selectionStory());
      const frame = mutate(() =>
        session.apply_delete_profiled(direction, expectedFrameEpoch, count)
      );
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    residentWorkerSnapshot: (options) => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      const mirrored = workerDocumentVersion !== null;
      const selectionJson = mirrored ? 'null' : session.selection();
      const mediaSources = mirrored ? undefined : session.media_sources_json();
      const fontsCurrent = options?.knownFontsRevision === residentFontsRevision;
      let state: Uint8Array | null = null;
      if (!mirrored && options?.knownStateVector) {
        try {
          state = session.encode_diff(options.knownStateVector.slice());
        } catch {
          state = null;
        }
      }
      return {
        clientId,
        ...(mirrored ? { workerAuthoritative: true as const } : {}),
        state: mirrored ? new Uint8Array(0) : (state ?? session.encode_state()),
        selection: JSON.parse(selectionJson) as YrsSelection | null,
        fonts: fontsCurrent
          ? []
          : residentFonts.map((font) =>
              font instanceof Uint8Array ? font.slice() : { ...font }
            ),
        fontsRevision: residentFontsRevision,
        renderInputs: [...residentRenderInputs].map(([story, env]) => ({
          story,
          env: structuredClone(env),
        })),
        measureInputs: [...residentMeasureInputs.values()],
        layoutInput: residentLayoutInput,
        layoutWithRegions: residentLayoutWithRegions,
        layoutRevision: residentLayoutRevision,
        ...(partialDocument ? { partialDocument: true } : {}),
        ...(mediaSources ? { mediaSources } : {}),
      };
    },
    residentWorkerProbe: () => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      return { layoutRevision: residentLayoutRevision };
    },
    displayHitTestRegionsJson: (pageIndex, x, y) =>
      session.display_hit_test_regions_json(pageIndex, x, y),
    displayVerticalMoveJson: (position, direction, goalX) =>
      session.display_vertical_move_json(position, direction, goalX),
    displayRangeRectsJson: (from, to) => session.display_range_rects_json(from, to),
    displayRangeRectsRegionJson: (region, rId, from, to) =>
      session.display_range_rects_region_json(region, rId, from, to),
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),

    loadState: (update) => {
      markDirty('all');
      mutate(() => {
        session.load(update);
        proposals.reset();
      });
    },
    seedFromDocx: (bytes, options) => openDocx(bytes, true, options),
    openDocx,
    beginOpening: (generation) => {
      mutate(() => {
        session.begin_opening(generation);
        proposals.reset();
      });
    },
    seedOpaqueSequences: (names) => {
      markDirty('all');
      mutate(() => session.seed_opaque_sequences(JSON.stringify(names)));
    },
    mediaSource: (token) => {
      if (destroyed || !token.startsWith('media:')) return null;
      const bytes = session.media_bytes(token);
      const mimeType = bytes && session.media_type(token);
      return bytes && mimeType ? { bytes, mimeType } : null;
    },
    loadMediaSources: (json) => session.load_media_sources(json),
    mediaDataUrl,
    mediaScope: () => mediaScope,
    materializeDocx: () => {
      const source = docxSource;
      const json = session.materialize_docx();
      if (!source || json === undefined) return null;
      return decodeS9Envelope(json, docxSourceBuffer(source)).document;
    },
    loadStories: (stories) => {
      markDirty(stories.map((seed) => seed.storyId));
      return mutate(
        () => JSON.parse(session.load_json(JSON.stringify(stories))) as Record<string, string[]>
      );
    },
    encodeState: () => session.encode_state(),
    encodeStateVector: () => session.encode_state_vector(),
    encodeStateAsUpdate: (remoteStateVector) =>
      remoteStateVector === undefined
        ? session.encode_state()
        : session.encode_diff(remoteStateVector.slice()),
    applyUpdate: (update) => {
      markDirty('all');
      return mutate(
        () =>
          JSON.parse(
            session.apply_update_with_inference(update)
          ) as CollaborationTextInsertion | null
      );
    },
    applyLocalUpdate: (update) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.apply_local_update(update));
    },
    applyHostUpdate: (update, stories) => {
      markDirty(stories ?? 'all');
      mutate(() => session.apply_host_update(update));
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

    encodeStickyPosition: (loc) => ({
      story: loc.story,
      encoded: session.encode_sticky_position(loc.story, loc.paraId, loc.offset),
    }),
    resolveStickyPosition: (position) => {
      try {
        return JSON.parse(
          session.resolve_sticky_position(position.story, position.encoded)
        ) as YrsLoc;
      } catch {
        return null;
      }
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
    encodeSelection: () => decodeEncodedSelection(session.encoded_selection()),
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
    beginUndoCapture: ensureUndo,
    addUndoBoundary: () => session.add_undo_boundary(),
    setUndoCaptureMode: (mode) => session.set_undo_capture_mode(mode),
    undoCaptureMode: () => session.undo_capture_mode() as YrsUndoCaptureMode,
    historyStories: () => session.history_stories(),
    undo: () =>
      mutate(() => {
        const applied = session.undo();
        if (applied) markDirty(session.history_stories());
        return applied;
      }),
    redo: () =>
      mutate(() => {
        const applied = session.redo();
        if (applied) markDirty(session.history_stories());
        return applied;
      }),
    canUndo: () => session.can_undo(),
    canRedo: () => session.can_redo(),

    createStory: (storyId, initialText, pStyle = 'Normal', alignment = 'left') => {
      markDirty(storyId);
      return mutate(
        () =>
          JSON.parse(session.create_story(storyId, initialText, pStyle, alignment)) as {
            paraId: string;
          }
      );
    },
    deleteStory: (storyId) => {
      markDirty(storyId);
      return mutate(() => session.delete_story(storyId));
    },
    insertTable: (at, rows, columns, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
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
          )
      );
    },
    insertRow: (at, side, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.insert_row(
                JSON.stringify(at),
                side === 'below',
                suggesting?.name,
                suggesting?.date
              )
            ) as YrsTableReceipt
          )
      );
    },
    insertColumn: (at, side) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.insert_column(JSON.stringify(at), side === 'right')) as YrsTableReceipt
          )
      );
    },
    deleteRow: (range, suggesting) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.delete_row(JSON.stringify(range), suggesting?.name, suggesting?.date)
            ) as YrsTableReceipt
          )
      );
    },
    deleteColumn: (range) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.delete_column(JSON.stringify(range))) as YrsTableReceipt)
      );
    },
    deleteTable: (table) => {
      ensureUndo(table.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.delete_table(JSON.stringify(table))) as YrsTableReceipt)
      );
    },
    mergeCells: (range) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.merge_cells(JSON.stringify(range))) as YrsTableReceipt)
      );
    },
    splitCell: (at, rows, columns) => {
      ensureUndo(at.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.split_cell(JSON.stringify(at), rows, columns)) as YrsTableReceipt)
      );
    },
    setCellShading: (range, color) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_shading(JSON.stringify(range), color ?? undefined)
            ) as YrsTableReceipt
          )
      );
    },
    setCellTextFormat: (range, patch) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_text_format(JSON.stringify(range), JSON.stringify(patch))
            ) as YrsTableReceipt
          )
      );
    },
    setCellBorders: (range, borders) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_borders(JSON.stringify(range), JSON.stringify(borders))
            ) as YrsTableReceipt
          )
      );
    },
    setColumnWidth: (at, widthTwips) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.set_column_width(JSON.stringify(at), widthTwips)) as YrsTableReceipt
          )
      );
    },
    setTableWidth: (table, widthTwips) => {
      ensureUndo(table.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.set_table_width(JSON.stringify(table), widthTwips)) as YrsTableReceipt
          )
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
          ) as YrsReplaceReceipt
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
          ) as YrsReplaceReceipt
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
          ) as YrsReplaceReceipt
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
          ) as YrsReplaceReceipt
      );
    },
    setContentControlValue: (embedId, value) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.set_content_control_value(embedId, JSON.stringify(value)));
    },
    setContentControlValueAt: (at, value) => {
      ensureUndo(at.story);
      if (typeof value === 'string') markDirty('all');
      mutate(() =>
        session.set_content_control_value_at(at.story, at.paraId, at.offset, JSON.stringify(value))
      );
    },
    clearContentControlValue: (embedId) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.clear_content_control_value(embedId));
    },
    setImageGeometry: (embedId, geometry) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.set_image_geometry(embedId, JSON.stringify(geometry)));
    },
    setImageGeometryAt: (at, geometry) => {
      ensureUndo(at.story);
      mutate(() =>
        session.set_image_geometry_at(at.story, at.paraId, at.offset, JSON.stringify(geometry))
      );
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
    applyRawOps: (story, ops) => {
      const rekeys = ops.some(
        (op) =>
          (op.op === 'insertEmbed' && op.kind === 'pilcrow') ||
          (op.op === 'setEmbedAttr' && op.key === 'paraId')
      );
      markDirty(rekeys ? 'all' : story);
      mutate(() => session.apply_raw_ops(story, JSON.stringify(ops)));
    },
    applySeedRawOps: (story, ops) => {
      markDirty(story);
      mutate(() => session.apply_seed_raw_ops(story, JSON.stringify(ops)));
    },
    setParagraphAttr: (paraId, key, value) => {
      markDirty('all');
      mutate(() => session.set_paragraph_attr(paraId, key, JSON.stringify(value ?? null)));
    },
    addComment: (ranges, commentAuthor, date, body) => {
      markDirty(ranges.map((range) => range.story));
      return mutate(
        () =>
          JSON.parse(
            session.add_comment(
              wireRanges(ranges),
              commentAuthor,
              date,
              JSON.stringify(body ?? null)
            )
          ) as YrsCommentReceipt
      );
    },
    setCommentRanges: (commentId, ranges) => {
      ensureUndo();
      mutate(() => {
        session.set_comment_ranges(commentId, wireRanges(ranges));
        markDirty('all');
      });
    },
    acceptChange: (target) => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.accept_change(wireChangeTarget(target))) as YrsResolveReceipt
      );
    },
    rejectChange: (target) => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.reject_change(wireChangeTarget(target))) as YrsResolveReceipt
      );
    },

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
    listComments: () => JSON.parse(session.list_comments()) as YrsCommentInfo[],
    storyIds: () => session.story_ids(),
    hasStory: (story) => !LONE_SURROGATE.test(story) && session.has_story(story),
    storyLength: (story) => session.story_len(story),
    storyChecksum: (story) => BigInt(session.story_checksum(story)),
    yrsBlocksForStory: (story, env = {}) => {
      const json = session.yrs_blocks_for_story(story, JSON.stringify(env));
      const blocks = JSON.parse(json) as unknown[];
      residentRenderInputs.set(story, structuredClone(env));
      return blocks;
    },
    paragraphs: (story) => JSON.parse(session.paragraphs(story)) as YrsParagraph[],
    searchText: (query, options = {}) => {
      if (!query) return [];
      const limit = options.limit ?? Number.POSITIVE_INFINITY;
      if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
        throw new RangeError('search limit must be a non-negative safe integer');
      }
      return JSON.parse(
        session.search_text(
          query,
          options.caseSensitive ?? false,
          Number.isFinite(limit) ? Math.min(limit, 0xffffffff) : undefined
        )
      ) as YrsTextMatch[];
    },
    paragraphSpans: (story) => JSON.parse(session.paragraph_spans(story)) as YrsParagraphLength[],
    storySegments: (story) => JSON.parse(session.story_segments(story)) as YrsStorySegment[],
    storiesChangedSince: (since) =>
      JSON.parse(session.stories_changed_since(since)) as { revision: number; stories: string[] },
    storySegmentUnitDigests: (story) =>
      JSON.parse(session.story_segment_unit_digests(story)) as string[],
    storySegmentUnits: (story, units) =>
      JSON.parse(session.story_segment_units(story, Uint32Array.from(units))) as YrsStorySegment[][],
    tablePayload: (story, tableIndex) => {
      // No table has an index the u32 boundary would wrap; the story must still exist.
      if (!Number.isInteger(tableIndex) || tableIndex < 0 || tableIndex > 0xffffffff) {
        session.story_len(story);
        return null;
      }
      const payload = session.table_payload(story, tableIndex);
      return payload === undefined ? null : (JSON.parse(payload) as Record<string, unknown>);
    },
    locateParagraph: (story, paraId) =>
      JSON.parse(session.locate_paragraph(story, paraId)) as YrsParagraphSpan,
    paragraphIdCount: (story, paraId) => session.paragraph_id_count(story, paraId),

    paragraphIdentities: () =>
      JSON.parse(session.paragraph_identities()) as DocxParagraphIdentitySnapshot,
    persistParagraphIds: () => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.persist_paragraph_ids()) as DocxParagraphIdentityReceipt
      );
    },
    resolveParagraphAnchor: (anchor) =>
      JSON.parse(
        session.resolve_paragraph_anchor(JSON.stringify(anchor))
      ) as DocxParagraphAnchorResult,
    storyParagraphIds: (story) =>
      JSON.parse(session.story_paragraph_ids(story)) as Array<string | null>,
    paragraphSavePlan: () => JSON.parse(session.paragraph_save_plan()) as DocxParagraphSavePlan,
    recordSavedParagraphIds: (saved) =>
      mutate(
        () =>
          JSON.parse(session.record_saved_paragraph_ids(JSON.stringify(saved))) as Array<
            [string, string]
          >
      ),
    writtenParagraphIds: (bytes) =>
      JSON.parse(session.written_paragraph_ids(bytes)) as Record<string, string[]>,
    exportStructured: (options) =>
      JSON.parse(
        session.export_structured_json(JSON.stringify(options))
      ) as DocxExportResult<DocxStructuredContent>,
    exportMarkdown: (options) =>
      JSON.parse(
        session.export_markdown_json(JSON.stringify(options))
      ) as DocxExportResult<DocxMarkdownContent>,
    exportStructuredWithPages: (options) =>
      JSON.parse(
        session.export_structured_with_pages_json(JSON.stringify(options), undefined)
      ) as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>,
    exportStructuredWithPagesFor: (options, currentRequest) =>
      JSON.parse(
        session.export_structured_with_pages_json(JSON.stringify(options), currentRequest)
      ) as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>,
    exportSnapshotWithPrivateFonts: (fonts, fontLengths, request, options) =>
      JSON.parse(
        session.export_snapshot_with_private_fonts_json(
          fonts,
          fontLengths,
          request,
          JSON.stringify(options)
        )
      ) as
        | { ok: true; content: DocxPagedStructuredContent<DocxSnapshotLayoutMap> }
        | { ok: false; failure: DocxExportFailure },
    headings: (story) => JSON.parse(session.headings_json(story)) as DocxParagraphHeading[],
    listContentControls: (options = {}) =>
      JSON.parse(
        session.list_content_controls_json(JSON.stringify(options))
      ) as DocxContentControlsResult,
    findContentControls: (query, options = {}) =>
      JSON.parse(
        session.find_content_controls_json(JSON.stringify(query), JSON.stringify(options))
      ) as DocxContentControlsResult,

    version: () => workerDocumentVersion ?? session.version(),
    readParagraphs: (request) =>
      JSON.parse(session.read_paragraphs_json(JSON.stringify(request))) as DocxReadParagraphsResult,
    findText: (request) =>
      JSON.parse(session.find_text_json(JSON.stringify(request))) as DocxFindTextResult,
    validateEdits: (request) =>
      JSON.parse(session.validate_edits_json(JSON.stringify(request))) as DocxValidationResult,
    applyEdits: (request) =>
      mutate(() => {
        const result = JSON.parse(
          session.apply_edits_json(JSON.stringify(request))
        ) as DocxEditResult;
        if (result.ok && result.applied) markDirty(result.changedStories);
        return result;
      }),
    proposeChanges: (request) => mutate(() => proposals.propose(request)),
    setProposalStates: (request) => proposals.setStates(request),
    withdrawProposals: (request) => mutate(() => proposals.withdraw(request)),
    getProposals: () => proposals.snapshot(),
    mirrorWorkerDocument: (mirror) => {
      workerDocumentVersion = mirror?.version ?? null;
      proposals.mirror(mirror);
    },
    workerDocumentMirrored: () => workerDocumentVersion !== null,
    onProposalChange: (listener) => {
      if (destroyed) throw new Error('yrs session is destroyed');
      return proposals.subscribe(listener);
    },
    formatTextTarget: (target, delta) => {
      ensureUndo(targetStory(target));
      return mutate(
        () =>
          JSON.parse(
            session.format_text_target_json(JSON.stringify(target), JSON.stringify(delta))
          ) as YrsTargetEditResult
      );
    },
    commentTextTarget: (target, comment) => {
      markDirty(targetStory(target));
      return mutate(
        () =>
          JSON.parse(
            session.comment_text_target_json(
              JSON.stringify(target),
              JSON.stringify({ ...comment, body: comment.body ?? null })
            )
          ) as YrsTargetEditResult
      );
    },
    selectionText: (range) =>
      JSON.parse(
        session.selection_text_json(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset
        )
      ) as YrsSelectionText,

    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      resetMedia();
      listeners.clear();
      proposals.destroy();
      pendingUpdates.length = 0;
      if (observing) session.clear_update_observer();
      session.free();
    },
  };

  registerSessionInternals(facade, {
    sourcePackage: () =>
      docxSource && docxSourceKeys
        ? { buffer: docxSourceBuffer(docxSource), keys: docxSourceKeys }
        : null,
    compareDocx: (original, revised, options) => {
      markDirty('all');
      const json = mutate(() => {
        const compared = session.compare_docx_json(original, revised, options);
        proposals.reset();
        return compared;
      });
      docxSource = original.slice();
      docxSourceKeys = null;
      return json;
    },
    finishComparedDocx: (bytes) => session.finish_compared_docx_json(bytes),
    failComparedDocx: (message) => session.fail_compared_docx_json(message),
  });

  return facade;
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
