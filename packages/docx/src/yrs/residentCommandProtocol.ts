import type { DocxEditRequest, DocxEditResult, DocxTextRange } from './edits';
import type { DisplayListRect } from '../layout/render/displayListQueries';
import type { ResidentFontEnvironment, ResidentFontEnvironmentSnapshot } from './residentFontEnvironment';
import type { ResidentCommentRecoveryState } from './residentCommentProjection';
import type { DocxReservationsPayload, DocxResolvedReservations } from './reservations';
import type {
  ApplyContext,
  ApplyResult,
  CommandAck,
  Checkpoint,
  CommandConnection,
  CommandEnvelope,
  CommandHost,
  CommandMethods,
  EditSession,
  EchoHooks,
  FrameEnvelope,
  RecoveryHooks,
  WorkerEditSessionOptions,
} from '../../../../shared/office-session/commands/types';
import type { TransferResult } from '../../../../shared/office-session/protocol';
import type { MethodHandlers, MethodPolicies } from '../../../../shared/office-session/types';
import type { ColorValue } from '../types/colors';
import type { Comment } from '../types/content';
import type {
  YrsAuthor,
  YrsContentControlValue,
  YrsHyperlinkAttrs,
  YrsInlineFormatDelta,
  YrsLoc,
  YrsParagraphAttrs,
  YrsResidentCaretSnapshot,
  YrsResidentFontRegistration,
  YrsRunMark,
  YrsSelectionContext,
  YrsSession,
  YrsStickyPosition,
  YrsStoryRange,
  YrsCellBorders,
  YrsTableRange,
} from './index';
import type { DocxClipboardFragment } from './clipboardFragment';
import type { LayoutMetaV1 } from './layoutMeta';
import type { DecodedFrameDelta } from '../layout/render/frameDelta';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { DocxProposalRegistry, DocxProposalRegistryState, DocxProposalRequest, DocxProposalStateRequest, DocxProposalWithdrawRequest } from './proposals';
import type { DocxCommentPatch } from './residentReview';
import type { DocxCellNavigationPlan } from './residentCommandCells';
import type { DocxTableCommandTarget } from './residentCommandTables';
import type { DocxPartTarget, ResidentPartTarget, ResidentPartGeometry } from './residentParts';
import type { SectionProperties, Watermark, FootnoteProperties, EndnoteProperties } from '../types/document';
import type { DocxSectionMargins } from './residentSections';
import type { ResidentFrameGeometry } from './residentObjectGeometry';

export type DocxCommandPosition = YrsLoc & { sticky?: Uint8Array };

export interface DocxCommandSelection {
  anchor: DocxCommandPosition;
  focus: DocxCommandPosition;
}

export interface DocxStoredMarks {
  story: string;
  paraId: string;
  clear: boolean;
  delta: YrsInlineFormatDelta;
}

export interface DocxCommandTableContext {
  story: string;
  tableIndex: number;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasMultiCellSelection: boolean;
  canSplitCell: boolean;
  rowSpan?: number;
  columnSpan?: number;
  justification?: string;
  cellBorderColor?: ColorValue;
  cellBackgroundColor?: string;
}

export type DocxEffectiveFormatting = Pick<
  YrsSelectionContext,
  | 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript'
  | 'fontFamily' | 'color' | 'highlight' | 'styleId' | 'alignment'
> & { fontSizePt: number | null };

/** The comment an `addComment` command added, under the ID saves and the host use, with every comment after it. */
export interface DocxCommentResult {
  seq: number;
  id: number;
  comments: Comment[];
}

/** Every comment after an `updateComment`, `setCommentRanges` or `removeComment` command. */
export interface DocxCommentChangeResult {
  seq: number;
  comments: Comment[];
}

export interface DocxEditState {
  documentVersion?: string;
  batch?: { seq: number; result: DocxEditResult };
  comment?: DocxCommentResult;
  commentChange?: DocxCommentChangeResult;
  activePart?: ResidentPartTarget | null;
  readOnly: boolean;
  review?: { version: string; previewVersion: number };
  toolbar: YrsSelectionContext | null;
  effectiveFormatting: DocxEffectiveFormatting | null;
  storedMarks: DocxStoredMarks | null;
  canUndo: boolean;
  canRedo: boolean;
  hyperlink: YrsHyperlinkAttrs | null;
  tableContext: DocxCommandTableContext | null;
  /** The tracked change whose range holds the selection head, host proposals aside. */
  currentRevisionId?: string;
}

/** Set by a host allowed to propose on a read-only document (`allowHostProposals`). */
export interface DocxHostProposalGrant {
  hostProposal?: boolean;
}

export interface DocxSelectionIntent {
  selection?: DocxCommandSelection;
  suggesting?: YrsAuthor;
}

/** @internal */
export type DocxNavigationIntent =
  | {
    unit: 'character' | 'word' | 'line' | 'paragraph' | 'document-boundary';
    direction: 'backward' | 'forward';
    extend: boolean;
  }
  | { unit: 'line'; direction: 'up' | 'down'; extend: boolean; goalX: number };

/** @internal Captured retained-frame identity; pageId is the decimal retained page ID. */
export interface DocxCommandPageIdentity {
  pageId: string;
  pageIndex: number;
}

/** @internal */
export interface DocxPointModifiers {
  shift: boolean;
  alt: boolean;
  ctrl: boolean;
  meta: boolean;
}

/** @internal Page-local layout coordinates, never document offsets. */
export interface DocxSelectionPointIntent {
  page: DocxCommandPageIdentity;
  x: number;
  y: number;
  layoutRevision: number;
  frameEpoch: number;
  clickCount: number;
  modifiers: DocxPointModifiers;
  story?: string;
  cellId?: string;
}

/** @internal */
export interface DocxEchoRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** @internal Canvas affine transform from context-local to page-local coordinates. */
export type DocxEchoTransform = readonly [number, number, number, number, number, number];

/** @internal Display-only indices into the projected context text. */
export interface DocxEchoStyleRun {
  startUtf16: number;
  endUtf16: number;
  fontFamily: string;
  fontSizePx: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  transform: DocxEchoTransform;
  baseline: { x: number; y: number };
}

/** @internal Bidirectional boundary mapping within the projected text. */
export interface DocxEchoGraphemeBoundary {
  utf16: number;
  grapheme: number;
  caret: DocxEchoRect;
}

/** @internal Text-free restoration dependency supplied with a staged page. */
export interface DocxEchoDamageDependency {
  page: DocxCommandPageIdentity;
  rect: DocxEchoRect;
  backgroundId: string;
  /** The echoed line's runs and decorations, which the background leaves out. */
  omittedPrimitiveIds: readonly string[];
}

/**
 * @internal The echoed line as the engine places it: one cluster per character, shaped through
 * one glyph set, at pen positions `round3(x + sum of advances)`.
 */
export interface DocxEchoGlyphLine {
  /** Key of the glyph set the line shapes with. */
  glyphSet: string;
  fontId: number;
  size: number;
  color: string;
  x: number;
  y: number;
  /** The paragraph character after the line, which kerns with its last one; empty at the end. */
  next: string;
  /** Rightmost pen end a lengthened line keeps without wrapping. */
  right: number;
  canEmpty?: boolean;
  /** Free width at the end of the previous line, when this line continues a paragraph. */
  previousSlack?: number;
  /** Width of the next line's leading letters and digits, when a line follows. */
  nextWord?: number;
  /** Repainted region in device pixels, and the scale pages raster at. */
  device: DocxEchoRect;
  pixelRatio: number;
}

/** @internal Shaping facts of one font at one size, indexed by `alphabet`; see `echo_glyph_set_json`. */
export interface DocxEchoGlyphSet {
  key: string;
  fontId: number;
  alphabet: string;
  horizontal: number;
  glyphIds: readonly number[];
  advances: readonly number[];
  kerning: readonly (readonly [number, number, number])[];
  unsafePairs: readonly (readonly [number, number])[];
  /** Outline JSON per glyph id, exactly as the worker's glyph cache reads it. */
  outlines: Readonly<Record<string, string>>;
}

/** @internal Bounded, display-only active context; never a command target. */
export interface DocxCommandEchoContext {
  page: DocxCommandPageIdentity;
  frameEpoch: number;
  layoutRevision: number;
  serializedBytes: number;
  text: string;
  styleRuns: readonly DocxEchoStyleRun[];
  boundaries: readonly DocxEchoGraphemeBoundary[];
  selection: { anchorGrapheme: number; focusGrapheme: number };
  transform: DocxEchoTransform;
  damage: readonly DocxEchoDamageDependency[];
  shaping: 'simple-ltr' | 'unsupported';
  glyphs?: DocxEchoGlyphLine;
}

/** @internal Transferable authoritative pixels and optional text-free restoration pixels. */
export interface DocxStagedPageBitmap {
  page: DocxCommandPageIdentity;
  frameEpoch: number;
  bitmap: ImageBitmap;
  backgrounds?: readonly { id: string; rect: DocxEchoRect; bitmap: ImageBitmap }[];
}

/** @internal */
export interface DocxCommandStagedPresentation {
  frameEpoch: number;
  pages: readonly DocxStagedPageBitmap[];
}

export interface DocxCommandPayloads {
  applyRemoteUpdate: { update: ArrayBuffer };
  applyRemoteUpdateChunks: { updates: ArrayBuffer[] };
  applyEdits: { request: DocxEditRequest; suggesting?: YrsAuthor };
  setRenderEnvironment: {
    themeColors: Record<string, string>;
    showHiddenText: boolean;
    showTableGridlines?: boolean;
    /** Applied only while the document version is still `reservations.version`. */
    reservations?: DocxReservationsPayload;
  };
  setFontEnvironment: ResidentFontEnvironment;
  setContentControlValue: { handle: string; value: YrsContentControlValue; expectVersion: string };
  addRepeatingItem: { handle: string; expectVersion: string };
  removeRepeatingItem: { handle: string; expectVersion: string };
  openPart: DocxPartTarget;
  ensurePart: DocxPartTarget;
  removePart: DocxPartTarget;
  setHyperlink: { selection: DocxCommandSelection; hyperlink: YrsHyperlinkAttrs | null; displayText?: string; expectVersion: string };
  setObjectProperties: { handle: string; patch: Record<string, unknown>; expectVersion: string };
  setSectionProperties: { section: string; patch: Partial<SectionProperties>; expectVersion: string };
  setNoteProperties: { section: string; footnotePr: FootnoteProperties; endnotePr: EndnoteProperties; expectVersion: string };
  setWatermark: { watermark: Watermark | null; expectVersion: string };
  setSectionMargins: { section: string; margins: DocxSectionMargins; expectVersion: string };
  setTabStops: DocxSelectionIntent & { tabs: YrsParagraphAttrs['tabs']; expectVersion: string };
  removeTabStop: DocxSelectionIntent & { positionTwips: number; expectVersion: string };
  formatSelection: DocxSelectionIntent & { delta: YrsInlineFormatDelta };
  toggleStrike: DocxSelectionIntent;
  toggleList: DocxSelectionIntent & { kind: 'bullet' | 'numbered' };
  changeIndent: DocxSelectionIntent & { direction: 'increase' | 'decrease' };
  insertImage: DocxSelectionIntent & { image: Readonly<Record<string, unknown>> };
  insertTable: DocxSelectionIntent & { rows: number; columns: number };
  insertTableRow: DocxSelectionIntent & { side: 'above' | 'below' };
  insertTableColumn: DocxSelectionIntent & { side: 'left' | 'right' };
  deleteTableRow: DocxSelectionIntent;
  deleteTableColumn: DocxSelectionIntent;
  deleteTable: DocxSelectionIntent;
  mergeTableCells: DocxSelectionIntent;
  splitTableCell: DocxSelectionIntent & { rows: number; columns: number };
  setTableCellShading: DocxSelectionIntent & { color: string | null };
  setTableBorders: DocxSelectionIntent & { borders: YrsCellBorders };
  setTableProperties: DocxSelectionIntent & { op: 'patch'; patch: { width?: number | null; widthType?: string | null; justification?: 'left' | 'center' | 'right' | null }; expectVersion?: string };
  pasteFragment: DocxSelectionIntent & { fragment: DocxClipboardFragment };
  insertPageBreak: DocxSelectionIntent;
  insertSectionBreak: DocxSelectionIntent & { type: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage' };
  resizeObject: { handle: string; expectVersion: string } &
    ({ widthEmu: number; heightEmu: number } | { widthTwips: number });
  moveObject: { handle: string; hOffsetEmu: number; vOffsetEmu: number; expectVersion: string };
  replaceMatches: { targets: DocxTextRange[]; text: string; expectVersion: string; suggesting?: YrsAuthor };
  undo: Record<string, never>;
  redo: Record<string, never>;
  addComment: { ranges: YrsStoryRange[]; author: string; date: string; body: unknown; expectVersion: string; parentId?: string };
  updateComment: { id: string; patch: DocxCommentPatch; expectVersion: string };
  setCommentRanges: { id: string; ranges: YrsStoryRange[]; expectVersion: string };
  removeComment: { id: string; expectVersion: string };
  acceptRevision: { revisionId: string; expectVersion: string };
  rejectRevision: { revisionId: string; expectVersion: string };
  propose: DocxProposalRequest & DocxHostProposalGrant;
  setProposalStates: DocxProposalStateRequest & DocxHostProposalGrant;
  withdrawProposals: DocxProposalWithdrawRequest & DocxHostProposalGrant;
  navigateCell: DocxSelectionIntent & { direction: 'forward' | 'backward'; extend: boolean; createLastRow: boolean };
  selectTableCells: DocxSelectionIntent & { target: 'table' | 'row' | 'column' };
  select: { selection: DocxCommandSelection };
  navigate: DocxNavigationIntent;
  selectPoint: DocxSelectionPointIntent;
  extendSelection: DocxSelectionPointIntent;
  insertText: DocxSelectionIntent & { text: string };
  replaceSelection: DocxSelectionIntent & { text: string };
  deleteBackward: DocxSelectionIntent & { count?: number };
  deleteForward: DocxSelectionIntent & { count?: number };
  paragraphBreak: DocxSelectionIntent;
  lineBreak: DocxSelectionIntent;
  toggleMark: DocxSelectionIntent & { mark: YrsRunMark };
  formatRange: DocxSelectionIntent & { selection: DocxCommandSelection; delta: YrsInlineFormatDelta };
  clearFormatting: DocxSelectionIntent;
  setParagraphAttrs: DocxSelectionIntent & { attrs: YrsParagraphAttrs };
  applyParagraphStyle: DocxSelectionIntent & { styleId: string };
  setReadOnly: { readOnly: boolean };
}

export type DocxCommandKind = keyof DocxCommandPayloads;
export type DocxCommandPayload = DocxCommandPayloads[DocxCommandKind];
export type DocxCommandEnvelope = CommandEnvelope<DocxCommandPayload>;
export type DocxCommandIntent = {
  [K in DocxCommandKind]: CommandEnvelope<DocxCommandPayloads[K]> & { kind: K };
}[DocxCommandKind];

/**
 * @internal The reservation set a layout includes: the `version` and `key` of the payload it was
 * resolved from, or a null `key` for entries kept from an earlier set.
 */
export interface DocxFrameReservations {
  version: string;
  key: string | null;
}

export interface DocxCommandFramePayload {
  fontEnvironment?: ResidentFontEnvironmentSnapshot;
  reservations?: DocxFrameReservations;
  parts?: ResidentPartGeometry[];
  geometry?: ResidentFrameGeometry;
  frame: ArrayBuffer;
  pageFrames?: ArrayBuffer[];
  caret: YrsResidentCaretSnapshot;
  layoutRevision: number;
  selectionRects?: DisplayListRect[];
  documentVersion?: string;
  layoutMeta?: LayoutMetaV1;
  layoutProvisional?: boolean;
  finalizedPages?: number;
  documentWidth?: number;
  echoContext?: DocxCommandEchoContext;
  /** Glyph sets the main thread has not been sent yet in this session. */
  echoGlyphSets?: DocxEchoGlyphSet[];
  stagedPresentation?: DocxCommandStagedPresentation;
}

export type DocxCommandFrame = FrameEnvelope<DocxCommandFramePayload>;
export type DocxCommandAck = CommandAck<DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload>;
export type DocxCommandApplyResult = ApplyResult<DocxCommandSelection | null, DocxEditState>;
export type DocxCommandReply = DocxCommandAck | TransferResult<DocxCommandAck>;

/** @internal */
export interface ResidentCommandRecoveryState {
  selection: DocxCommandSelection | null;
  storedMarks: DocxStoredMarks | null;
  activePart?: ResidentPartTarget | null;
  stickyPositions: [string, YrsStickyPosition][];
  cellSelection: YrsTableRange | null;
  proposals: DocxProposalRegistryState;
  selections: [number, DocxCommandSelection | null][];
  comments?: ResidentCommentRecoveryState;
  layoutReservations?: { values: DocxResolvedReservations; set?: DocxFrameReservations };
}

/** @internal */
export interface ResidentCommandCheckpoint extends Checkpoint {
  recoveryState?: ResidentCommandRecoveryState;
}

export interface ResidentCommandInitRequest {
  session: string;
  incarnation: number;
  epoch: number;
  version?: number;
  throughSeq?: number;
  refusedSeqs?: readonly number[];
  replayAppliedSeqs?: readonly number[];
  recoveryState?: ResidentCommandRecoveryState;
  readOnly: boolean;
  selection?: DocxCommandSelection | null;
  fonts: YrsResidentFontRegistration[];
  fontRegistrationSlots?: boolean | { fontChains: Record<string, number[]> };
  layoutInput: string;
  layoutWithRegions: boolean;
  layoutRevision: number;
  extras: string;
  displayWindow?: [number, number];
  retainBuiltPages?: boolean;
  provisionalPages?: number;
  rasterScale?: number;
}

export type ResidentCommandInitReply = DocxCommandAck &
  Required<Pick<DocxCommandAck, 'selection' | 'editState' | 'frame'>>;
export type DocxCommandMethods = CommandMethods<
  DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
>;

export interface DocxCommandApplyContext extends ApplyContext {
  predecessorSelection?: DocxCommandSelection | null;
}

export interface PreparedDocxCommand {
  command: DocxCommandIntent;
  selection: DocxCommandSelection | null;
  range: YrsStoryRange | null;
  cellNavigation?: DocxCellNavigationPlan;
  cells?: YrsTableRange | null;
  table?: DocxTableCommandTarget;
  paragraphChanges?: Array<{ range: YrsStoryRange; attrs: YrsParagraphAttrs }>;
  sectionBreakProperties?: SectionProperties;
}

export type DocxCommandPreflightResult =
  | { status: 'ready'; prepared: PreparedDocxCommand }
  | Extract<DocxCommandApplyResult, { status: 'refused' }>;

export interface ResidentCommandHooks {
  preflight(command: DocxCommandEnvelope, context: DocxCommandApplyContext):
    DocxCommandPreflightResult | Promise<DocxCommandPreflightResult>;
  apply(command: PreparedDocxCommand, context: DocxCommandApplyContext):
    DocxCommandApplyResult | Promise<DocxCommandApplyResult>;
  selection(): DocxCommandSelection | null;
  editState(): DocxEditState;
  frame(input: { coversThroughSeq: number; version: number; full: boolean }):
    DocxCommandFrame | TransferResult<DocxCommandFrame> |
    Promise<DocxCommandFrame | TransferResult<DocxCommandFrame>>;
  snapshot(): Uint8Array | Promise<Uint8Array>;
  recoveryState?(): Omit<ResidentCommandRecoveryState, 'selections'>;
}

/** @internal Worker-owned layout access for ordered selection resolution. */
export interface ResidentCommandHooksEnvironment {
  supportsFontEnvironment?: boolean;
  proposals?: DocxProposalRegistry;
  /** The session merges remote updates; its collaboration wrap applies them. */
  collaboration?: boolean;
  readSelectionLayout(): DocxCommandSelectionLayout;
  getLayoutReservations?(): DocxResolvedReservations | null;
  /**
   * Omits `reservations` to keep the current ones; `set` names the payload they were resolved
   * from, null when they were kept from an earlier set.
   */
  setRenderEnvironment?(environment: {
    themeColors: Record<string, string>;
    showHiddenText: boolean;
    showTableGridlines?: boolean;
    reservations?: DocxResolvedReservations;
  }, set?: DocxFrameReservations | null): void;
}

export type ResidentCommandHooksFactory = (
  engine: ResidentEngineSession,
  init: ResidentCommandInitRequest,
  environment?: ResidentCommandHooksEnvironment
) => Pick<ResidentCommandHooks, 'preflight' | 'apply' | 'selection' | 'editState' | 'recoveryState'> & {
  /** Runs after the collaboration wrap applies a remote merge at `context.currentVersion`. */
  afterRemoteMerge?(context: DocxCommandApplyContext): void;
};

export interface ResidentCommandRegistry {
  handlers: Partial<MethodHandlers<DocxCommandMethods, unknown>>;
  policies: Partial<MethodPolicies<DocxCommandMethods>>;
}

export interface ResidentCommandHostOptions {
  init: ResidentCommandInitRequest;
  hooks: ResidentCommandHooks;
  registry: ResidentCommandRegistry;
  retainedOutcomes?: number;
}

export interface ResidentCommandHost extends CommandHost<
  DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
> {
  withFrameWork<T>(run: () => Promise<T>): Promise<T>;
  initial(): Promise<ResidentCommandInitReply | TransferResult<ResidentCommandInitReply>>;
  requestFullFrame(): Promise<DocxCommandReply>;
}

export interface BuildCommandFrameInput extends Omit<DocxCommandFramePayload, 'frame' | 'pageFrames'> {
  frame: Uint8Array;
  metadata: Pick<DecodedFrameDelta, 'full' | 'frameEpoch' | 'baseFrameEpoch'>;
  pageFrames?: readonly Uint8Array[];
  coversThroughSeq: number;
}

export interface ResidentCommandViewportRequest {
  start: number;
  end: number;
  frameEpoch: number;
  rasterScale?: number;
}

export interface ResidentCommandClient {
  readonly commandGeneration: number;
  readonly commandFailure: Error | undefined;
  commandInitialGeneration(initial: ResidentCommandInitReply): number | undefined;
  commandInit(request: ResidentCommandInitRequest): Promise<ResidentCommandInitReply>;
  edit(command: DocxCommandEnvelope): Promise<DocxCommandAck>;
  checkpoint(): Promise<ReturnType<DocxCommandMethods['checkpoint']>>;
  requestFullFrame(): Promise<DocxCommandAck>;
  requestViewport?(request: ResidentCommandViewportRequest): Promise<void>;
  onFrame?(listener: (frame: DocxCommandFrame) => void): () => void;
  onFailure(listener: ((error: Error) => void) | null): void;
  destroy(): void;
}

export interface ResidentCommandConnection extends CommandConnection<
  DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
> {
  initial: ResidentCommandInitReply;
  requestFullFrame(): Promise<DocxCommandAck>;
  requestViewport?(request: ResidentCommandViewportRequest): Promise<void>;
}

export type DocxWorkerEditSession<X = never> = EditSession<
  DocxCommandPayload, DocxCommandSelection | null, DocxEditState, X
>;

export type DocxWorkerEditSessionOptions<O, X> = Omit<WorkerEditSessionOptions<
  O, DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload, X
>, 'connection' | 'echo' | 'requestFullFrame' | 'recovery'> & {
  connection: ResidentCommandConnection;
  echo: EchoHooks<X, DocxCommandFramePayload>;
  recovery: Omit<RecoveryHooks<
    O, DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
  >, 'reopen'> & {
    reopen(input: Parameters<RecoveryHooks<
      O, DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
    >['reopen']>[0]): Promise<ResidentCommandConnection>;
  };
};

export interface ResidentDocxCommandState {
  activePart?: ResidentPartTarget | null;
  readOnly: boolean;
  selection: DocxCommandSelection | null;
  storedMarks: DocxStoredMarks | null;
  stickyPositions: Map<string, YrsStickyPosition>;
  proposals?: DocxProposalRegistry;
  /** Host versions in `(from, to]`, all produced by applied remote merges and none by a local command. */
  remoteMerges?: { from: number; to: number };
}

export type ResidentDocxCommandEngine = ResidentEngineSession & Pick<YrsSession,
  | 'paragraphs' | 'selectionContext' | 'storySegments' | 'cellSelection' | 'tablePayload'
  | 'insertText' | 'insertLineBreak' | 'replaceRange' | 'deleteRange' | 'splitParagraph'
  | 'toggleMark' | 'formatRange' | 'clearFormatting' | 'setParagraphAttrs' | 'applyParagraphStyle'
  | 'insertColumn' | 'deleteRow' | 'deleteColumn' | 'deleteTable'
  | 'mergeCells' | 'splitCell' | 'setCellShading'
  | 'beginUndoCapture' | 'canUndo' | 'canRedo'
>;

export type ResolvedDocxCommandSelection =
  | { ok: true; selection: DocxCommandSelection; range: YrsStoryRange }
  | { ok: false; reason: string };
