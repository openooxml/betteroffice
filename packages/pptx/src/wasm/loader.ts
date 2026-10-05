import initWasmModule, {
  decodeTiffPng,
  exportPptxMarkdownJson,
  exportPptxStructuredJson,
  parsePptxJson,
  PptxDocument,
  PptxRenderer,
  renderPptxMarkdownJson,
  rendererVersion,
} from './generated/pptx_wasm.js';
import type { InitInput } from './generated/pptx_wasm.js';
import { wasmAssetUrl } from './asset';
import { StaleProposalError } from '../proposals';
import { PptxExportError } from '../structuredExport';
import type {
  PptxExportFailure,
  PptxExportOptions,
  PptxExportResult,
  PptxMarkdownContent,
  PptxMarkdownOptions,
  PptxStructuredContent,
} from '../structuredExport';
import type { Proposal, ProposalAcceptance, ProposalDiffSlide, ProposalEdit, ProposalPreview } from '../proposals';
import type {
  PptxEditRequest,
  PptxEditResult,
  PptxFindRequest,
  PptxFindResult,
  PptxReadRequest,
  PptxReadResult,
  PptxValidationResult,
} from '../edits';
import type {
  CollaborationReplica,
  CollaborationUpdateOrigin,
} from '../collaboration/types';
import type {
  CommentFlavor,
  CommentReceipt,
  CommentSnapshot,
  DeckSnapshot,
  HistoryProfile,
  HistoryResult,
  HitTestResult,
  ParagraphAlignment,
  PictureDraft,
  PresetShapeDraft,
  Profiled,
  ProfiledLayout,
  PptxFontFace,
  PptxCaretAnchor,
  PptxTextMatch,
  PptxTextSearchOptions,
  ShapeAdjustReceipt,
  ShapeDraft,
  ShapeFillReceipt,
  ShapeReceipt,
  ShapeRect,
  ShapeStroke,
  ShapeStrokeReceipt,
  ShapeZOrderReceipt,
  SlideDisplayList,
  SlideReceipt,
  StorySnapshot,
  TextReceipt,
  TextStyle,
  TextStylePatch,
  TransformReceipt,
} from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;

export interface OpenPresentationOptions {
  clientId?: number;
  fonts?: ReadonlyArray<PptxFontFace>;
  /** Faces drawn only for characters the run's own face has no glyph for, e.g. a script's Noto face. */
  fallbackFonts?: ReadonlyArray<PptxFontFace>;
  /**
   * Opens from a collaboration update instead of parsing the file bytes.
   * When the bytes are the file the update was seeded from, the session
   * keeps them and `save()` works; any other bytes open without a source
   * and `save()` throws.
   */
  initialUpdate?: Uint8Array;
}

export type UndoCaptureMode = 'auto' | 'manual';

export interface PresentationHandle extends CollaborationReplica {
  isProposalsAvailable(): boolean;
  propose(agentId: string, note: string | null, edits: readonly ProposalEdit[]): Proposal;
  listProposals(): Proposal[];
  previewProposal(id: string): ProposalPreview;
  layoutProposalSlide(id: string, slideIndex: number): SlideDisplayList;
  layoutProposalDiffSlide(id: string, slideIndex: number): ProposalDiffSlide;
  acceptProposal(id: string, options?: { force?: boolean }): ProposalAcceptance;
  rejectProposal(id: string): boolean;
  readonly clientId: number;
  snapshot(): DeckSnapshot;
  story(storyId: string): StorySnapshot;
  /** Anchors the UTF-16 caret `index` of a story so later edits move it along. */
  anchorCaret(storyId: string, index: number): PptxCaretAnchor;
  /** The anchor's current offset, or `null` once its story is gone. */
  resolveCaretAnchor(anchor: PptxCaretAnchor): number | null;
  /**
   * The session-scoped version token. It changes with every committed change, local or remote,
   * undo and redo included; compare tokens only within this session.
   */
  version(): string;
  /** Slides and their stories' text, with the version they were read at. */
  readContent(request?: PptxReadRequest): PptxReadResult;
  /** Exact, case-sensitive, paragraph-local search; overlapping matches count separately. */
  findText(request: PptxFindRequest): PptxFindResult;
  /**
   * Structured slide content with the version it was read at, read from committed state: pending
   * editor input is not flushed and nothing changes. Option refusals are returned; malformed
   * options throw.
   */
  exportStructured(options?: PptxExportOptions): PptxExportResult<PptxStructuredContent>;
  /** `exportStructured` rendered as Markdown from the same read. */
  exportMarkdown(options?: PptxExportOptions): PptxExportResult<PptxMarkdownContent>;
  /** Runs every check of `applyEdits`, staging included, without changing anything. */
  validateEdits(request: PptxEditRequest): PptxValidationResult;
  /**
   * Applies every step or none against `expectVersion`, as one transaction, one update and, for
   * `history: 'separate'`, one undo step. Policy failures are returned; malformed requests throw.
   */
  applyEdits(request: PptxEditRequest): PptxEditResult;
  /** Literal search in slide order. */
  searchText(query: string, options?: PptxTextSearchOptions): PptxTextMatch[];
  registerFont(face: PptxFontFace): number;
  /** Registers a face drawn only where a run's own face has no glyph. */
  registerFallbackFont(face: PptxFontFace): number;
  layoutSlide(slideIndex: number): SlideDisplayList;
  /** `layoutSlide` with scope, layout and serialize time measured inside the renderer. */
  layoutSlideProfiled(slideIndex: number): ProfiledLayout;
  hitTest(x: number, y: number): HitTestResult | null;
  mediaBytes(partPath: string): Uint8Array;
  /** serialize the presentation back to .pptx bytes, edits included. */
  save(): Uint8Array;
  insertText(storyId: string, index: number, text: string, style?: TextStyle): TextReceipt;
  /** `insertText` with the boundary's stage timings attached. */
  insertTextProfiled(
    storyId: string,
    index: number,
    text: string,
    style?: TextStyle
  ): Profiled<TextReceipt>;
  deleteText(storyId: string, start: number, end: number): TextReceipt;
  deleteTextProfiled(storyId: string, start: number, end: number): Profiled<TextReceipt>;
  formatText(storyId: string, start: number, end: number, patch: TextStylePatch): TextReceipt;
  insertParagraphBreak(storyId: string, index: number): TextReceipt;
  /** Sets the alignment of every paragraph the range touches; `null` restores
   *  the inherited value. */
  setParagraphAlignment(
    storyId: string,
    start: number,
    end: number,
    alignment: ParagraphAlignment | null
  ): TextReceipt;
  insertSlide(index: number, layoutPartPath?: string): SlideReceipt;
  insertSlideProfiled(index: number, layoutPartPath?: string): Profiled<SlideReceipt>;
  deleteSlide(slideId: string): SlideReceipt;
  moveSlide(slideId: string, toIndex: number): SlideReceipt;
  /** Sets a slide's speaker notes; empty text clears them. */
  setSlideNotes(slideId: string, text: string): void;
  addTextBox(slideId: string, draft: ShapeDraft): ShapeReceipt;
  addTextBoxProfiled(slideId: string, draft: ShapeDraft): Profiled<ShapeReceipt>;
  addShape(slideId: string, draft: PresetShapeDraft): ShapeReceipt;
  addPicture(slideId: string, draft: PictureDraft): ShapeReceipt;
  setShapeFill(slideId: string, shapeId: string, color: string | null): ShapeFillReceipt;
  setShapeStroke(
    slideId: string,
    shapeId: string,
    stroke: ShapeStroke
  ): ShapeStrokeReceipt;
  setShapeAdjust(
    slideId: string,
    shapeId: string,
    adjustments: Record<string, number>
  ): ShapeAdjustReceipt;
  removeShape(slideId: string, shapeId: string): ShapeReceipt;
  /** Moves a shape to the top of its slide's paint order (drawn last). */
  bringShapeToFront(slideId: string, shapeId: string): ShapeZOrderReceipt;
  /** Moves a shape to the bottom of its slide's paint order (drawn first). */
  sendShapeToBack(slideId: string, shapeId: string): ShapeZOrderReceipt;
  /** Swaps a shape one step later in its slide's paint order. */
  bringShapeForward(slideId: string, shapeId: string): ShapeZOrderReceipt;
  /** Swaps a shape one step earlier in its slide's paint order. */
  sendShapeBackward(slideId: string, shapeId: string): ShapeZOrderReceipt;
  /** Adds a slide comment; coordinates are EMU. */
  addComment(
    slideId: string,
    comment: {
      author: string;
      initials?: string;
      text: string;
      created: string;
      xEmu?: number;
      yEmu?: number;
    }
  ): CommentReceipt;
  /** Modern decks only; legacy `p:cm` has no reply list. */
  replyToComment(
    commentId: string,
    reply: { author: string; initials?: string; text: string; created: string }
  ): CommentReceipt;
  /** Resolves or reopens a modern comment. */
  setCommentStatus(commentId: string, resolved: boolean): CommentReceipt;
  /** Moves a root comment on its slide; coordinates are safe integer EMU. */
  setCommentPosition(commentId: string, position: { xEmu: number; yEmu: number }): CommentReceipt;
  removeComment(commentId: string): CommentReceipt;
  /** Only legal while the deck has no comments. */
  setCommentFlavor(flavor: CommentFlavor): CommentFlavor;
  comments(): CommentSnapshot[];
  moveShape(slideId: string, shapeId: string, x: number, y: number): TransformReceipt;
  moveShapeProfiled(
    slideId: string,
    shapeId: string,
    x: number,
    y: number
  ): Profiled<TransformReceipt>;
  resizeShape(slideId: string, shapeId: string, width: number, height: number): TransformReceipt;
  setShapeRect(slideId: string, shapeId: string, rect: ShapeRect): TransformReceipt;
  canUndo(): boolean;
  canRedo(): boolean;
  undoCaptureMode(): UndoCaptureMode;
  setUndoCaptureMode(mode: UndoCaptureMode): void;
  addUndoBoundary(): void;
  undo(): HistoryResult;
  /** `undo` with undo, snapshot and serialize time measured at the boundary. */
  undoProfiled(): Profiled<HistoryResult, HistoryProfile>;
  redo(): HistoryResult;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  encodeDiff(remoteStateVector: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): DeckSnapshot;
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;
  dispose(): void;
}

let initialized = false;

/** @internal */
export class PresentationPeerError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = 'PresentationPeerError';
  }
}

const replayMethods = [
  'insertText', 'deleteText', 'formatText', 'insertParagraphBreak', 'setParagraphAlignment',
  'insertSlide', 'deleteSlide', 'moveSlide', 'setSlideNotes', 'addTextBox', 'addShape', 'addPicture',
  'removeShape', 'moveShape', 'resizeShape', 'setShapeRect', 'setShapeFill', 'setShapeStroke',
  'setShapeAdjust', 'bringShapeToFront', 'sendShapeToBack', 'bringShapeForward', 'sendShapeBackward',
  'addComment', 'replyToComment', 'setCommentStatus', 'setCommentPosition', 'removeComment',
  'setCommentFlavor', 'propose', 'acceptProposal', 'rejectProposal', 'applyEdits',
  'addUndoBoundary', 'undo', 'redo',
] as const satisfies readonly (keyof PresentationHandle)[];

/** @internal */
export type PresentationReplayOp = {
  [K in typeof replayMethods[number]]: { method: K; args: Parameters<PresentationHandle[K]> }
}[typeof replayMethods[number]];

/** @internal */
export interface PresentationReplayOutcome {
  result: unknown;
  applied: boolean;
  changedTargets: string[];
  canUndo: boolean;
  canRedo: boolean;
}

/** @internal */
export interface PresentationReplayEnvelope {
  sequence: number;
  baseVersion: string;
  op: PresentationReplayOp;
  expectedOutcome?: PresentationReplayOutcome | null;
}

/** @internal */
export interface PresentationReplayReply {
  sequence: number;
  revision: number;
  version: string;
  engineVersion: string;
  consumed: boolean;
  outcome: PresentationReplayOutcome;
}

/** @internal */
export type PresentationPeerHandle = Pick<PresentationHandle,
  'clientId' | 'snapshot' | 'story' | 'anchorCaret' | 'resolveCaretAnchor' | 'version' |
  'readContent' | 'findText' | 'exportStructured' | 'exportMarkdown' | 'validateEdits' |
  'searchText' | 'isProposalsAvailable' | 'listProposals' | 'previewProposal' |
  'layoutProposalSlide' | 'layoutProposalDiffSlide' | 'layoutSlide' | 'hitTest' |
  'mediaBytes' | 'save' | 'canUndo' | 'canRedo' | 'undoCaptureMode' | 'onUpdate' | 'dispose'
>;

type PeerMode = { kind: 'baseline' } | { kind: 'peer'; identity: string };
type PeerInternals = {
  registerFonts(): Promise<void>;
  adopt(): void;
  hydration(): string;
  metadata(): PresentationMetadata;
  displayListJson(slideIndex: number): string;
  replay(envelope: PresentationReplayEnvelope, captured?: (reply: PresentationReplayReply) => void): PresentationReplayReply;
};
const peerInternals = new WeakMap<PresentationPeerHandle, PeerInternals>();

function peerInternal(handle: PresentationPeerHandle): PeerInternals {
  const internal = peerInternals.get(handle);
  if (!internal) throw new PresentationPeerError('stage', 'Presentation is not a retained replay handle');
  return internal;
}

/** @internal */
export function openPresentationReplayBaseline(
  bytes: Uint8Array, options: OpenPresentationOptions = {}
): PresentationPeerHandle {
  try {
    return openPresentationInternal(bytes, options, { kind: 'baseline' });
  } catch (error) {
    throw peerError(error);
  }
}

/** @internal */
export function openPresentationPeerDeck(
  bytes: Uint8Array, identity: string, options: Omit<OpenPresentationOptions, 'clientId'> = {}
): PresentationPeerHandle {
  try {
    return openPresentationInternal(bytes, options, { kind: 'peer', identity });
  } catch (error) {
    throw peerError(error);
  }
}

/** @internal */
export function registerPresentationPeerFonts(handle: PresentationPeerHandle): Promise<void> {
  return peerInternal(handle).registerFonts();
}

/** @internal */
export function adoptPresentationPeerIdentity(handle: PresentationPeerHandle): void {
  peerInternal(handle).adopt();
}

/** @internal */
export function presentationPeerHydration(handle: PresentationPeerHandle): string {
  return peerInternal(handle).hydration();
}

/** @internal */
export function presentationPeerMetadata(handle: PresentationPeerHandle): PresentationMetadata {
  return peerInternal(handle).metadata();
}

/** @internal */
export function presentationPeerDisplayListJson(handle: PresentationPeerHandle, slideIndex: number): string {
  return peerInternal(handle).displayListJson(slideIndex);
}

/** @internal */
export function replayPresentation(
  handle: PresentationPeerHandle, envelope: PresentationReplayEnvelope,
  captured?: (reply: PresentationReplayReply) => void
): PresentationReplayReply {
  return peerInternal(handle).replay(envelope, captured);
}

const displayListJsonReaders = new WeakMap<PresentationHandle, (slideIndex: number) => string>();

type SessionMetadataDocument = PptxDocument & { sessionMetadataJson(): string };
type PresentationMetadata = {
  slides: { id: string; index: number; name: string | null; layoutPartPath: string | null }[];
  size: { width: number; height: number };
};
const metadataReaders = new WeakMap<PresentationHandle, () => PresentationMetadata>();

export function presentationMetadata(handle: PresentationHandle): PresentationMetadata {
  const read = metadataReaders.get(handle);
  if (!read) throw new Error('Presentation metadata is unavailable');
  return read();
}

/** @experimental */
export function presentationDisplayListJson(handle: PresentationHandle, slideIndex: number): string {
  const read = displayListJsonReaders.get(handle);
  if (!read) throw new Error('Presentation display list is unavailable');
  return read(slideIndex);
}

export function isProposalsAvailable(): boolean {
  return typeof PptxDocument.prototype.proposeJson === 'function'
    && typeof PptxRenderer.prototype.layoutProposalSlideJson === 'function';
}
let initialization: Promise<void> | undefined;

export function initWasm(
  input: WasmInitInput = wasmAssetUrl()
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

export function wasmVersion(): string {
  requireInitialized();
  return rendererVersion();
}

export function inspectPresentation(bytes: Uint8Array): unknown {
  requireInitialized();
  return call(() => parsePptxJson(bytes));
}

type ExportOutcome<T> = { ok: true; content: T } | { ok: false; failure: PptxExportFailure };

function exported<T>(operation: () => string): T {
  const outcome = jsonCall<ExportOutcome<T>>(operation);
  if (!outcome.ok) throw new PptxExportError(outcome.failure);
  return outcome.content;
}

/**
 * Exports PPTX bytes as structured content whose anchors address the returned snapshot. No
 * session, DOM or font is involved. Throws `PptxExportError` for refused options and an `Error`
 * for bytes that are not a readable PPTX.
 */
export async function exportPptxStructured(
  bytes: Uint8Array,
  options: PptxExportOptions = {}
): Promise<PptxStructuredContent> {
  await initWasm();
  const json = requestJson(options);
  return exported(() => exportPptxStructuredJson(bytes, json));
}

/** `exportPptxStructured` rendered as Markdown with anchor markers. */
export async function exportPptxMarkdown(
  bytes: Uint8Array,
  options: PptxExportOptions = {}
): Promise<PptxMarkdownContent> {
  await initWasm();
  const json = requestJson(options);
  return exported(() => exportPptxMarkdownJson(bytes, json));
}

/**
 * Renders structured content (schema version 1) as Markdown. Content the renderer cannot trust
 * refuses with `invalid-content`.
 */
export async function renderPptxMarkdown(
  content: PptxStructuredContent,
  options: PptxMarkdownOptions = {}
): Promise<PptxMarkdownContent> {
  await initWasm();
  const json = requestJson(options);
  return exported(() => renderPptxMarkdownJson(JSON.stringify(content), json));
}

export function decodeTiffImage(bytes: Uint8Array): Uint8Array {
  requireInitialized();
  return construct(() => decodeTiffPng(bytes));
}

export function openPresentation(
  bytes: Uint8Array,
  options: OpenPresentationOptions = {}
): PresentationHandle {
  return openPresentationInternal(bytes, options);
}

function openPresentationInternal(
  bytes: Uint8Array, options: OpenPresentationOptions, peer?: PeerMode
): PresentationHandle {
  requireInitialized();
  const faces = peer ? [
    ...(options.fonts ?? []).map((face) => ({ ...face, bytes: new Uint8Array(face.bytes), fallback: false })),
    ...(options.fallbackFonts ?? []).map((face) => ({ ...face, bytes: new Uint8Array(face.bytes), fallback: true })),
  ] : [];
  const peerUpdate = peer && options.initialUpdate !== undefined
    ? new Uint8Array(options.initialUpdate) : undefined;
  const collaborationClientId = peer ? options.clientId : options.clientId ?? clientId();
  const doc = construct(() => peer
    ? peer.kind === 'baseline'
      ? PptxDocument.openReplayBaseline(new Uint8Array(bytes), collaborationClientId, peerUpdate)
      : PptxDocument.openPeerDeckJson(new Uint8Array(bytes), peer.identity, peerUpdate)
    : options.initialUpdate === undefined
      ? PptxDocument.openCollaborative(bytes, collaborationClientId!)
      : PptxDocument.openCollaborativeFromUpdate(
          options.initialUpdate.slice(),
          collaborationClientId!,
          bytes.slice()
        )
  );
  let renderer: PptxRenderer;
  try {
    renderer = construct(() => new PptxRenderer());
  } catch (error) {
    if (peer) {
      try { doc.free(); } catch {}
      throw peerError(error);
    }
    throw error;
  }
  if (!peer) {
    for (const face of options.fonts ?? []) registerFont(renderer, face);
    for (const face of options.fallbackFonts ?? []) registerFont(renderer, face, true);
  }
  const listeners = new Map<
    number,
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  const pendingUpdates: Array<{
    update: Uint8Array;
    origin: CollaborationUpdateOrigin;
  }> = [];
  let nextListenerId = 0;
  let disposed = false;
  let observerInstalled = false;
  let wasmCallDepth = 0;
  let flushingUpdates = false;

  const assertAlive = (): void => {
    if (disposed) throw new Error('presentation handle is disposed');
  };

  const flushUpdates = (): void => {
    if (disposed || flushingUpdates || wasmCallDepth !== 0) return;
    flushingUpdates = true;
    try {
      while (!disposed && pendingUpdates.length > 0) {
        const event = pendingUpdates.shift();
        if (!event) break;
        for (const [id, listener] of [...listeners]) {
          if (disposed) return;
          if (listeners.get(id) !== listener) continue;
          try {
            listener(event.update.slice(), event.origin);
          } catch {}
        }
      }
    } finally {
      flushingUpdates = false;
      if (disposed) pendingUpdates.length = 0;
    }
  };

  const drainWasmUpdates = (): void => {
    if (!observerInstalled || disposed) return;
    while (true) {
      const encoded = doc.drainUpdateEvent();
      if (encoded.byteLength === 0) return;
      const origin = encoded[0];
      if (origin !== 0 && origin !== 1) {
        throw new Error(`pptx wasm returned unknown update origin ${origin}`);
      }
      pendingUpdates.push({
        update: encoded.subarray(1),
        origin: origin === 0 ? 'local' : 'remote',
      });
    }
  };

  const wasmCall = <T,>(operation: () => T, drainUpdates = false): T => {
    assertAlive();
    wasmCallDepth += 1;
    try {
      let result: T | undefined;
      let failure: unknown;
      let failed = false;
      try {
        result = operation();
      } catch (error) {
        failure = error;
        failed = true;
      }
      if (drainUpdates) {
        try {
          drainWasmUpdates();
        } catch (error) {
          if (!failed) {
            failure = error;
            failed = true;
          }
        }
      }
      if (failed) throw toError(failure);
      return result as T;
    } finally {
      wasmCallDepth -= 1;
      if (wasmCallDepth === 0) flushUpdates();
    }
  };

  const jsonWasmCall = <T,>(operation: () => string, drainUpdates = false): T =>
    wasmCall(() => JSON.parse(operation()) as T, drainUpdates);

  const ensureUpdateObserver = (): void => {
    if (observerInstalled) return;
    wasmCall(() => doc.startUpdateObservation());
    observerInstalled = true;
  };

  const clearUnusedUpdateObserver = (): void => {
    if (!observerInstalled || listeners.size > 0 || disposed) return;
    pendingUpdates.length = 0;
    wasmCall(() => doc.clearUpdateObservation());
    observerInstalled = false;
  };

  const handle: PresentationHandle = {
    isProposalsAvailable,
    propose(agentId, note, edits) {
      return jsonWasmCall(() => doc.proposeJson(JSON.stringify({ agentId, note, edits })));
    },
    listProposals() {
      return isProposalsAvailable() ? jsonWasmCall(() => doc.listProposalsJson()) : [];
    },
    previewProposal(id) {
      return jsonWasmCall(() => doc.previewProposalJson(JSON.stringify({ id })));
    },
    layoutProposalDiffSlide(id, slideIndex) {
      return jsonWasmCall(() => renderer.layoutProposalDiffSlideJson(doc, id, slideIndex));
    },
    layoutProposalSlide(id, slideIndex) {
      return jsonWasmCall(() => renderer.layoutProposalSlideJson(doc, id, slideIndex));
    },
    acceptProposal(id, options) {
      return jsonWasmCall(() => doc.acceptProposalJson(JSON.stringify({ id, force: options?.force ?? false })), true);
    },
    rejectProposal(id) {
      return jsonWasmCall(() => doc.rejectProposalJson(JSON.stringify({ id })));
    },
    get clientId(): number {
      return wasmCall(() => doc.clientId);
    },
    snapshot(): DeckSnapshot {
      return jsonWasmCall(() => doc.snapshotJson());
    },
    story(storyId: string): StorySnapshot {
      return jsonWasmCall(() => doc.storyJson(JSON.stringify({ storyId })));
    },
    anchorCaret(storyId, index): PptxCaretAnchor {
      return jsonWasmCall(() => doc.anchorCaretJson(JSON.stringify({ storyId, index })));
    },
    resolveCaretAnchor(anchor): number | null {
      return jsonWasmCall(() =>
        doc.resolveCaretAnchorJson(
          JSON.stringify({ storyId: anchor.storyId, position: anchor.position })
        )
      );
    },
    version(): string {
      return wasmCall(() => doc.documentVersion());
    },
    readContent(request = {}) {
      const json = requestJson(request);
      return jsonWasmCall(() => doc.readContentJson(json));
    },
    findText(request) {
      const json = requestJson(request);
      return jsonWasmCall(() => doc.findTextJson(json));
    },
    exportStructured(options = {}) {
      const json = requestJson(options);
      return jsonWasmCall(() => doc.exportStructuredJson(json));
    },
    exportMarkdown(options = {}) {
      const json = requestJson(options);
      return jsonWasmCall(() => doc.exportMarkdownJson(json));
    },
    validateEdits(request) {
      const json = requestJson(request);
      return jsonWasmCall(() => doc.validateEditsJson(json));
    },
    applyEdits(request) {
      const json = requestJson(request);
      return jsonWasmCall(() => doc.applyEditsJson(json), true);
    },
    searchText(query, options = {}) {
      if (!query) return [];
      const limit = options.limit ?? Number.POSITIVE_INFINITY;
      if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
        throw new RangeError('search limit must be a non-negative safe integer');
      }
      return jsonWasmCall(() =>
        doc.searchTextJson(JSON.stringify({
          query,
          caseSensitive: options.caseSensitive ?? false,
          limit: Number.isFinite(limit) ? Math.min(limit, 0xffffffff) : undefined,
        }))
      );
    },
    registerFont(face: PptxFontFace): number {
      return wasmCall(() => registerFont(renderer, face));
    },
    registerFallbackFont(face: PptxFontFace): number {
      return wasmCall(() => registerFont(renderer, face, true));
    },
    layoutSlide(slideIndex: number): SlideDisplayList {
      return jsonWasmCall(() => renderer.layoutSlideJson(doc, slideIndex));
    },
    layoutSlideProfiled(slideIndex: number): ProfiledLayout {
      return jsonWasmCall(() => renderer.layoutSlideProfiledJson(doc, slideIndex));
    },
    hitTest(x: number, y: number): HitTestResult | null {
      return jsonWasmCall(() => renderer.hitTestJson(x, y));
    },
    mediaBytes(partPath: string): Uint8Array {
      return wasmCall(() => doc.mediaBytes(partPath));
    },
    save(): Uint8Array {
      return wasmCall(() => doc.saveBytes());
    },
    insertText(storyId, index, text, style = {}): TextReceipt {
      return jsonWasmCall(
        () => doc.insertTextJson(JSON.stringify({ storyId, index, text, style })),
        true
      );
    },
    insertTextProfiled(storyId, index, text, style = {}): Profiled<TextReceipt> {
      return jsonWasmCall(
        () => doc.insertTextProfiledJson(JSON.stringify({ storyId, index, text, style })),
        true
      );
    },
    deleteText(storyId, start, end): TextReceipt {
      return jsonWasmCall(
        () => doc.deleteTextJson(JSON.stringify({ storyId, start, end })),
        true
      );
    },
    deleteTextProfiled(storyId, start, end): Profiled<TextReceipt> {
      return jsonWasmCall(
        () => doc.deleteTextProfiledJson(JSON.stringify({ storyId, start, end })),
        true
      );
    },
    formatText(storyId, start, end, patch): TextReceipt {
      return jsonWasmCall(
        () => doc.formatTextJson(JSON.stringify({ storyId, start, end, patch })),
        true
      );
    },
    insertParagraphBreak(storyId, index): TextReceipt {
      return jsonWasmCall(
        () => doc.insertParagraphBreakJson(JSON.stringify({ storyId, index })),
        true
      );
    },
    setParagraphAlignment(storyId, start, end, alignment): TextReceipt {
      return jsonWasmCall(
        () =>
          doc.setParagraphAlignmentJson(JSON.stringify({ storyId, start, end, alignment })),
        true
      );
    },
    insertSlide(index, layoutPartPath): SlideReceipt {
      return jsonWasmCall(
        () =>
          doc.insertSlideJson(JSON.stringify({ index, layoutPartPath: layoutPartPath ?? null })),
        true
      );
    },
    insertSlideProfiled(index, layoutPartPath): Profiled<SlideReceipt> {
      return jsonWasmCall(
        () =>
          doc.insertSlideProfiledJson(
            JSON.stringify({ index, layoutPartPath: layoutPartPath ?? null })
          ),
        true
      );
    },
    deleteSlide(slideId): SlideReceipt {
      return jsonWasmCall(() => doc.deleteSlideJson(JSON.stringify({ slideId })), true);
    },
    moveSlide(slideId, toIndex): SlideReceipt {
      return jsonWasmCall(
        () => doc.moveSlideJson(JSON.stringify({ slideId, toIndex })),
        true
      );
    },
    setSlideNotes(slideId, text): void {
      jsonWasmCall(() => doc.setSlideNotesJson(JSON.stringify({ slideId, text })), true);
    },
    addComment(slideId, comment): CommentReceipt {
      return jsonWasmCall(
        () =>
          doc.addCommentJson(
            JSON.stringify({
              slideId,
              author: comment.author,
              initials: comment.initials ?? '',
              text: comment.text,
              created: comment.created,
              xEmu: comment.xEmu ?? 0,
              yEmu: comment.yEmu ?? 0,
            })
          ),
        true
      );
    },
    replyToComment(commentId, reply): CommentReceipt {
      return jsonWasmCall(
        () =>
          doc.replyToCommentJson(
            JSON.stringify({
              commentId,
              author: reply.author,
              initials: reply.initials ?? '',
              text: reply.text,
              created: reply.created,
            })
          ),
        true
      );
    },
    setCommentPosition(commentId, position): CommentReceipt {
      if (!Number.isSafeInteger(position.xEmu) || !Number.isSafeInteger(position.yEmu)) {
        throw new Error('Comment coordinates must be safe integer EMU');
      }
      return jsonWasmCall(
        () => doc.setCommentPositionJson(JSON.stringify({ commentId, ...position })), true
      );
    },
    setCommentStatus(commentId, resolved): CommentReceipt {
      return jsonWasmCall(
        () => doc.setCommentStatusJson(JSON.stringify({ commentId, resolved })),
        true
      );
    },
    removeComment(commentId): CommentReceipt {
      return jsonWasmCall(() => doc.removeCommentJson(JSON.stringify({ commentId })), true);
    },
    setCommentFlavor(flavor): CommentFlavor {
      return jsonWasmCall(() => doc.setCommentFlavorJson(JSON.stringify({ flavor })), true);
    },
    comments(): CommentSnapshot[] {
      return jsonWasmCall(() => doc.commentsJson());
    },
    addTextBox(slideId, draft): ShapeReceipt {
      return jsonWasmCall(() => doc.addTextBoxJson(JSON.stringify({ slideId, draft })), true);
    },
    addTextBoxProfiled(slideId, draft): Profiled<ShapeReceipt> {
      return jsonWasmCall(
        () => doc.addTextBoxProfiledJson(JSON.stringify({ slideId, draft })),
        true
      );
    },
    addShape(slideId, draft): ShapeReceipt {
      return jsonWasmCall(() => doc.addShapeJson(JSON.stringify({ slideId, draft })), true);
    },
    addPicture(slideId, draft): ShapeReceipt {
      return jsonWasmCall(() => doc.addPictureJson(JSON.stringify({ slideId, ...draft })), true);
    },
    setShapeFill(slideId, shapeId, color): ShapeFillReceipt {
      return jsonWasmCall(
        () => doc.setShapeFillJson(JSON.stringify({ slideId, shapeId, color })),
        true
      );
    },
    setShapeStroke(slideId, shapeId, stroke): ShapeStrokeReceipt {
      return jsonWasmCall(
        () => doc.setShapeStrokeJson(JSON.stringify({ slideId, shapeId, stroke })),
        true
      );
    },
    setShapeAdjust(slideId, shapeId, adjustments): ShapeAdjustReceipt {
      return jsonWasmCall(
        () => doc.setShapeAdjustJson(JSON.stringify({ slideId, shapeId, adjustments })),
        true
      );
    },
    removeShape(slideId, shapeId): ShapeReceipt {
      return jsonWasmCall(
        () => doc.removeShapeJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    bringShapeToFront(slideId, shapeId): ShapeZOrderReceipt {
      return jsonWasmCall(
        () => doc.bringShapeToFrontJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    sendShapeToBack(slideId, shapeId): ShapeZOrderReceipt {
      return jsonWasmCall(
        () => doc.sendShapeToBackJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    bringShapeForward(slideId, shapeId): ShapeZOrderReceipt {
      return jsonWasmCall(
        () => doc.bringShapeForwardJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    sendShapeBackward(slideId, shapeId): ShapeZOrderReceipt {
      return jsonWasmCall(
        () => doc.sendShapeBackwardJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    moveShape(slideId, shapeId, x, y): TransformReceipt {
      return jsonWasmCall(
        () => doc.moveShapeJson(JSON.stringify({ slideId, shapeId, x, y })),
        true
      );
    },
    moveShapeProfiled(slideId, shapeId, x, y): Profiled<TransformReceipt> {
      return jsonWasmCall(
        () => doc.moveShapeProfiledJson(JSON.stringify({ slideId, shapeId, x, y })),
        true
      );
    },
    resizeShape(slideId, shapeId, width, height): TransformReceipt {
      return jsonWasmCall(
        () => doc.resizeShapeJson(JSON.stringify({ slideId, shapeId, width, height })),
        true
      );
    },
    setShapeRect(slideId, shapeId, rect): TransformReceipt {
      return jsonWasmCall(
        () => doc.setShapeRectJson(JSON.stringify({ slideId, shapeId, rect })),
        true
      );
    },
    canUndo(): boolean {
      return wasmCall(() => doc.canUndo());
    },
    canRedo(): boolean {
      return wasmCall(() => doc.canRedo());
    },
    undoCaptureMode(): UndoCaptureMode {
      return wasmCall(() => doc.undoCaptureMode()) as UndoCaptureMode;
    },
    setUndoCaptureMode(mode): void {
      wasmCall(() => doc.setUndoCaptureMode(mode));
    },
    addUndoBoundary(): void {
      wasmCall(() => doc.addUndoBoundary());
    },
    undo(): HistoryResult {
      return jsonWasmCall(() => doc.undoJson(), true);
    },
    undoProfiled(): Profiled<HistoryResult, HistoryProfile> {
      return jsonWasmCall(() => doc.undoProfiledJson(), true);
    },
    redo(): HistoryResult {
      return jsonWasmCall(() => doc.redoJson(), true);
    },
    encodeStateVector(): Uint8Array {
      return wasmCall(() => doc.encodeStateVector());
    },
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
      return wasmCall(() =>
        remoteStateVector === undefined
          ? doc.encodeStateAsUpdate()
          : doc.encodeDiff(remoteStateVector)
      );
    },
    encodeDiff(remoteStateVector): Uint8Array {
      return wasmCall(() => doc.encodeDiff(remoteStateVector));
    },
    applyUpdate(update): DeckSnapshot {
      return jsonWasmCall(() => doc.applyUpdateJson(update), true);
    },
    onUpdate(listener): () => void {
      assertAlive();
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      const id = nextListenerId++;
      listeners.set(id, listener);
      try {
        ensureUpdateObserver();
      } catch (error) {
        listeners.delete(id);
        throw error;
      }
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        listeners.delete(id);
        clearUnusedUpdateObserver();
      };
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      faces.length = 0;
      listeners.clear();
      pendingUpdates.length = 0;
      let disposalError: unknown;
      if (observerInstalled) {
        try {
          doc.clearUpdateObservation();
        } catch (error) {
          disposalError = error;
        }
        observerInstalled = false;
      }
      try {
        renderer.free();
      } catch (error) {
        disposalError ??= error;
      }
      try {
        doc.free();
      } catch (error) {
        disposalError ??= error;
      }
      if (disposalError !== undefined) throw toError(disposalError);
    },
  };
  displayListJsonReaders.set(handle, (slideIndex) =>
    wasmCall(() => renderer.layoutSlideJson(doc, slideIndex))
  );
  metadataReaders.set(handle, () =>
    jsonWasmCall(() => (doc as SessionMetadataDocument).sessionMetadataJson())
  );
  Object.defineProperty(handle, Symbol.for('@betteroffice/pptx/slide-layout-cache'), {
    value: {
      snapshot: (): { snapshot: DeckSnapshot; keys: Record<string, string> } =>
        jsonWasmCall(() => renderer.snapshotWithLayoutKeysJson(doc)),
      key: (index: number): string =>
        wasmCall(() => renderer.slideLayoutKey(doc, index)),
      activate: (slideId: string, key: string): boolean =>
        wasmCall(() => renderer.setActiveSlide(doc, slideId, key)),
      hitTest: (slideId: string, x: number, y: number): HitTestResult | null =>
        jsonWasmCall(() => renderer.hitTestSlideJson(doc, slideId, x, y)),
    },
  });
  if (peer) {
    let stage: 'deck' | 'registering' | 'fonts' | 'ready' | 'failed' = 'deck';
    let manifest = '';
    let hydration: string | undefined;
    let replaying = false;

    const failStage = (error: unknown): never => {
      stage = 'failed';
      faces.length = 0;
      try { handle.dispose(); } catch {}
      throw peerError(error);
    };
    const requireStage = (expected: 'deck' | 'registering' | 'fonts' | 'ready' | 'failed'): void => {
      if (disposed || stage !== expected) {
        throw new PresentationPeerError('stage', `Presentation peer requires ${expected} stage`);
      }
    };

    peerInternals.set(handle, {
      metadata() {
        requireStage('ready');
        return jsonWasmCall(() => (doc as SessionMetadataDocument).sessionMetadataJson());
      },
      displayListJson(slideIndex) {
        requireStage('ready');
        return wasmCall(() => renderer.layoutSlideJson(doc, slideIndex));
      },
      async registerFonts() {
        try {
          requireStage('deck');
          stage = 'registering';
          const entries = [];
          for (const face of faces) {
            const digest = await globalThis.crypto.subtle.digest('SHA-256', face.bytes);
            requireStage('registering');
            entries.push({
              family: face.family, bold: face.bold ?? false, italic: face.italic ?? false,
              fallback: face.fallback,
              fingerprint: Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join(''),
            });
          }
          manifest = JSON.stringify(entries);
          wasmCall(() => {
            for (const face of faces) registerFont(renderer, face, face.fallback);
            if (peer.kind === 'peer') doc.registerPeerFontsJson(manifest);
          });
          faces.length = 0;
          stage = 'fonts';
        } catch (error) {
          return failStage(error);
        }
      },
      adopt() {
        try {
          requireStage('fonts');
          if (peer.kind !== 'peer') {
            throw new PresentationPeerError('stage', 'Only a peer can adopt identity');
          }
          wasmCall(() => doc.adoptPeerIdentity());
          stage = 'ready';
        } catch (error) {
          return failStage(error);
        }
      },
      hydration() {
        if (peer.kind !== 'baseline') {
          throw new PresentationPeerError('stage', 'Only the worker baseline can capture hydration');
        }
        if (hydration !== undefined) {
          requireStage('ready');
          return hydration;
        }
        try {
          requireStage('fonts');
          hydration = wasmCall(() => doc.peerHydrationJson(manifest));
          stage = 'ready';
          return hydration;
        } catch (error) {
          return failStage(error);
        }
      },
      replay(envelope, captured) {
        requireStage('ready');
        if (replaying || flushingUpdates) {
          throw new PresentationPeerError('stage', 'Presentation replay cannot reenter');
        }
        replaying = true;
        try {
          const ownedJson = requestJson(envelope);
          return wasmCall(() => {
            const reply = JSON.parse(doc.replayJson(ownedJson)) as PresentationReplayReply;
            freezeReplayValue(reply);
            captured?.(reply);
            return reply;
          }, true);
        } finally {
          replaying = false;
        }
      },
    });
    for (const method of [
      ...replayMethods, 'insertTextProfiled', 'deleteTextProfiled', 'insertSlideProfiled',
      'addTextBoxProfiled', 'moveShapeProfiled', 'undoProfiled', 'applyUpdate',
      'setUndoCaptureMode', 'registerFont', 'registerFallbackFont',
    ] as const) {
      Object.defineProperty(handle, method, {
        value: () => { throw new PresentationPeerError('stage', 'Retained presentations require ordered replay'); },
      });
    }
  }
  return handle;
}

function freezeReplayValue(value: unknown): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freezeReplayValue(child);
  Object.freeze(value);
}

function peerError(error: unknown): PresentationPeerError {
  const mapped = toError(error);
  return mapped instanceof PresentationPeerError
    ? mapped : new PresentationPeerError('engine', mapped.message);
}

/** JSON for a host request; `JSON.stringify` would turn NaN and infinities into `null`. */
function requestJson(request: unknown): string {
  return JSON.stringify(request, (_key, value: unknown) => {
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new RangeError('host requests must not contain NaN or infinite numbers');
    }
    return value;
  });
}

function registerFont(renderer: PptxRenderer, face: PptxFontFace, fallback = false): number {
  try {
    const register = fallback ? renderer.registerFallbackFont : renderer.registerFont;
    return register.call(renderer, face.family, face.bold ?? false, face.italic ?? false, face.bytes);
  } catch (error) {
    throw toError(error);
  }
}

function requireInitialized(): void {
  if (!initialized) throw new Error('pptx wasm is not initialized; call initWasm() first');
}

function clientId(): number {
  const random = globalThis.crypto;
  if (!random || typeof random.getRandomValues !== 'function') {
    throw new Error('crypto.getRandomValues is required to generate a collaboration client ID');
  }
  const values = new Uint32Array(2);
  let value: number;
  do {
    random.getRandomValues(values);
    value = (values[0] & 0x1fffff) * 0x1_0000_0000 + values[1];
  } while (value === 0);
  return value;
}

function construct<T>(operation: () => T): T {
  try {
    return operation();
  } catch (error) {
    throw toError(error);
  }
}

function jsonCall<T>(operation: () => string): T {
  try {
    return JSON.parse(operation()) as T;
  } catch (error) {
    throw toError(error);
  }
}

function call<T>(operation: () => string): T {
  return jsonCall(operation);
}

function toError(error: unknown): Error {
  if (error instanceof PresentationPeerError) return error;
  if (error !== null && typeof error === 'object' && 'name' in error &&
      error.name === 'PresentationPeerError' && 'code' in error && typeof error.code === 'string' &&
      'message' in error && typeof error.message === 'string') {
    return new PresentationPeerError(error.code, error.message);
  }
  if (error instanceof Error) return error;
  if (typeof error === 'string') {
    try {
      const parsed = JSON.parse(error);
      if (parsed.code === 'staleProposal' && Array.isArray(parsed.targets)
          && parsed.targets.every((target: unknown) => typeof target === 'string')) {
        return new StaleProposalError(parsed.targets);
      }
    } catch {}
  }
  return new Error(typeof error === 'string' ? error : String(error));
}
