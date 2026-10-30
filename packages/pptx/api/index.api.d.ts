// @betteroffice/pptx "." public API surface
// Regenerate: bun scripts/api-surface.ts --write pptx
// ---- collaboration/asyncReplica.d.ts ----
import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
import type { AsyncCollaborationReplica, CollaborationReplica, PptxPresenceCursor, PptxPresencePeer } from './types';
/** @internal */
export declare function attachCollaborationPipe(replica: AsyncCollaborationReplica, channel: CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer>): void;
/** @internal */
export declare function collaborationPipeFor(replica: AsyncCollaborationReplica | CollaborationReplica): CollaborationPipeChannel<PptxPresenceCursor, PptxPresencePeer> | undefined;

// ---- collaboration/index.d.ts ----
export { CollaborationProvider } from './provider';
export { CollaborationError, type CollaborationUser, type CollaborationErrorCode, type CollaborationErrorListener, type CollaborationProviderOptions, type AsyncCollaborationReplica, type CollaborationReplica, type CollaborationStatus, type CollaborationStatusChange, type CollaborationStatusListener, type CollaborationTransport, type CollaborationTransportEvent, type CollaborationUpdateOrigin, type PptxPresence, type PptxPresenceCursor, type PptxPresenceListener, type PptxPresencePeer, type PptxPresenceState, type PptxPresenceUser, } from './types';
export { PRESENCE_LABEL_DURATION_MS, presenceColorForClientId } from './presence';
/** @internal */
export { attachCollaborationPipe, collaborationPipeFor } from './asyncReplica';

// ---- collaboration/presence.d.ts ----
import type { PptxPresenceCursor, PptxPresencePeer, PptxPresenceState, PptxPresenceUser } from './types';
export declare const PRESENCE_CURSOR_INTERVAL_MS = 80;
export declare const PRESENCE_HEARTBEAT_MS = 20000;
export declare const PRESENCE_EXPIRY_MS = 45000;
export declare const PRESENCE_CLOCK_RETENTION_MS: number;
export declare const PRESENCE_LABEL_DURATION_MS = 3000;
export declare const MAX_AWARENESS_STRING_LENGTH = 1024;
/** Full stores evict the least recently accepted client id first. */
export declare const MAX_TRACKED_PRESENCE_IDS = 4096;
export declare const PRESENCE_COLORS: readonly ["#B3261E", "#7B1FA2", "#3949AB", "#00796B", "#2E7D32", "#A84400", "#C2185B", "#00695C"];
export interface AwarenessUpdateEntry {
    clientId: number;
    clock: number;
    state: PptxPresenceState | null;
}
export interface PresencePeersOptions {
    maxTrackedIds?: number;
}
export declare function presenceColorForClientId(clientId: number): string;
export declare function sanitizePresenceColor(clientId: number, color: unknown): string;
export declare function presenceUser(clientId: number, user?: {
    name: string;
    color?: string;
}): PptxPresenceUser;
export declare function encodablePresenceCursor(cursor: PptxPresenceCursor): boolean;
export declare function validPresenceCursor(value: unknown, maxFrameBytes: number): value is PptxPresenceCursor;
export declare function samePresenceCursor(left: PptxPresenceCursor | null, right: PptxPresenceCursor | null): boolean;
export declare class PresencePeers {
    private readonly localClientId;
    private readonly states;
    private readonly clocks;
    private readonly maxTrackedIds;
    constructor(localClientId: number, options?: PresencePeersOptions);
    get peers(): readonly PptxPresencePeer[];
    get trackedIdCount(): number;
    get nextExpiryAt(): number | undefined;
    apply(entries: readonly AwarenessUpdateEntry[], now: number): boolean;
    expire(now: number, maxAge?: number): boolean;
    clear(): boolean;
    private evictOldest;
}

// ---- collaboration/provider.d.ts ----
import type { PipeArrival } from '../../../../shared/collaboration-pipe/worker';
import { type CollaborationErrorListener, type CollaborationProviderOptions, type AsyncCollaborationReplica, type CollaborationReplica, type CollaborationStatus, type CollaborationStatusListener, type CollaborationTransport, type PptxPresenceCursor, type PptxPresenceListener, type PptxPresencePeer } from './types';
/** @internal */
export declare function providerArrival(provider: CollaborationProvider): PipeArrival;
/** @internal */
export declare function providerLeave(provider: CollaborationProvider): {
    clock: number;
    frame: Uint8Array;
} | null;
export declare class CollaborationProvider {
    private readonly piped;
    private openEpoch;
    private readonly replica;
    private readonly transport;
    private readonly maxFrameBytes;
    private readonly maxMessagesPerFrame;
    private readonly maxPendingBytes;
    private readonly localClientId;
    private readonly localUser;
    private readonly remotePresence;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly presenceListeners;
    private readonly pendingFrames;
    private unsubscribeReplica;
    private unsubscribeTransport;
    private hasTransportSubscription;
    private connectionStatus;
    private isSynced;
    private wantsConnection;
    private isOpen;
    private isDestroyed;
    private epoch;
    private connectAttempt;
    private nextListenerId;
    private queuedBytes;
    private readonly incoming;
    private incomingBytes;
    private incomingBusy;
    private asyncReplica;
    private incomingGeneration;
    private isFlushing;
    private localClock;
    private reportedClockExhaustion;
    private localCursor;
    private applyingArrival;
    private lastAwarenessSentAt;
    private cursorTimer;
    private heartbeatTimer;
    private expiryTimer;
    constructor(replica: CollaborationReplica, transport: CollaborationTransport, options?: CollaborationProviderOptions);
    constructor(replica: AsyncCollaborationReplica, transport: CollaborationTransport, options?: CollaborationProviderOptions);
    get status(): CollaborationStatus;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly PptxPresencePeer[];
    setCursor(cursor: PptxPresenceCursor | null): void;
    connect(): void;
    disconnect(): void;
    destroy(): void;
    onStatus(listener: CollaborationStatusListener): () => void;
    onError(listener: CollaborationErrorListener): () => void;
    onPresence(listener: PptxPresenceListener): () => void;
    private beginArrival;
    private leaveFrame;
    private isCurrent;
    private failConnection;
    private takeTransportSubscription;
    private cleanupTransport;
    private handleTransportEvent;
    private handleOpen;
    private sendFullUpdate;
    private replicaResult;
    private handleMessage;
    private answerSyncStep1;
    private enqueueMessage;
    private drainIncoming;
    /** Takes the queued frames that carry only remote updates, up to the frame and worker byte limits, for one chunked apply. */
    private takeUpdateBatch;
    private remoteUpdates;
    private applyUpdateBatch;
    private forwardLocalUpdate;
    private scheduleCursorBroadcast;
    private broadcastLocalPresence;
    private sendLeaveFrameBestEffort;
    private startPresenceTimers;
    private stopPresenceTimers;
    private schedulePresenceExpiry;
    private clearRemotePresence;
    private emitPresence;
    private sendFrame;
    private queueFrame;
    private flushPending;
    private clearPending;
    private setStatus;
    private report;
}

// ---- collaboration/types.d.ts ----
export type CollaborationUpdateOrigin = 'local' | 'remote';
export interface CollaborationReplica {
    readonly clientId: number;
    encodeStateVector(): Uint8Array;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    applyUpdate(update: Uint8Array): unknown;
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
    /** Replicas that wrap a PresentationHandle must forward this to it. */
    setCollaborative?(): void;
}
export interface AsyncCollaborationReplica {
    readonly clientId: number;
    encodeStateVector(): Promise<Uint8Array>;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Promise<Uint8Array>;
    /** Optional. Encodes the update for `remoteStateVector` as chunks of at most `maxChunkBytes`. */
    encodeStateAsUpdateChunks?(remoteStateVector: Uint8Array, maxChunkBytes: number): Promise<Uint8Array[]>;
    applyUpdate(update: Uint8Array): Promise<unknown>;
    /** Optional. Resolves once the worker has merged the updates as a single change. */
    applyUpdateChunks?(updates: readonly Uint8Array[]): Promise<unknown>;
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}
export type CollaborationTransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
} | {
    type: 'drain';
};
export interface CollaborationTransport {
    /** Starts one connection attempt; transports may also reopen while still subscribed. */
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    /** Returns true only when the owned frame was accepted; false requires a later drain event. */
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: CollaborationTransportEvent) => void): () => void;
}
export type CollaborationStatus = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
export interface CollaborationStatusChange {
    status: CollaborationStatus;
    synced: boolean;
}
export type CollaborationErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export declare class CollaborationError extends Error {
    readonly code: CollaborationErrorCode;
    readonly cause: unknown;
    constructor(code: CollaborationErrorCode, message: string, cause?: unknown);
}
export interface CollaborationProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    maxPendingBytes?: number;
    user?: CollaborationUser;
}
export type CollaborationStatusListener = (change: CollaborationStatusChange) => void;
export type CollaborationErrorListener = (error: CollaborationError) => void;
export interface CollaborationUser {
    name: string;
    color?: string;
}
export interface PptxPresenceUser {
    name: string;
    color: string;
}
export interface PptxPresenceCursor {
    slideId: string;
    shapeId?: string;
}
export interface PptxPresenceState {
    clientId: number;
    clock: number;
    user: PptxPresenceUser;
    cursor: PptxPresenceCursor | null;
}
export interface PptxPresencePeer {
    state: PptxPresenceState;
    lastSeen: number;
    cursorMovedAt: number;
}
export type PptxPresenceListener = (peers: readonly PptxPresencePeer[]) => void;
export interface PptxPresence {
    readonly peers: readonly PptxPresencePeer[];
    setCursor(cursor: PptxPresenceCursor | null): void;
    onPresence(listener: PptxPresenceListener): () => void;
}

// ---- edits.d.ts ----
/**
 * Versioned reads and version-checked edit batches over one presentation session. Every shape
 * here is plain JSON so it can cross a worker boundary unchanged; the semantics live in Rust.
 *
 * Text offsets are half-open, story-local UTF-16 positions. A story reads as its paragraphs'
 * text joined by `\n` with no trailing separator; soft line breaks read as `\n` too, and
 * `PptxParagraphText` tells them apart. Slide, shape, story and paragraph ids anchor targets
 * within one session; they are not promised to survive save and reopen.
 */
import type { EditSuccess, OperationFailure, OperationRefusal, ValidationSuccess } from '../../../shared/host-contracts/edits';
import type { ParagraphAlignment, ShapeFill, ShapeOutline, ShapeRect, ShapeStroke, SlideSnapshot, TextStylePatch } from './types';
export type { EditSuccess, OperationFailure, OperationRefusal, ValidationSuccess, } from '../../../shared/host-contracts/edits';
/** Provenance only: it selects neither permissions nor history. */
export type PptxEditSource = 'host' | 'agent';
/** `separate`: exactly one undo step. `none`: outside undo history, keeping existing entries. */
export type PptxEditHistory = 'separate' | 'none';
export interface PptxSlideTarget {
    slideId: string;
}
/** A shape addressed through the slide that owns it. */
export interface PptxShapeTarget extends PptxSlideTarget {
    shapeId: string;
}
/**
 * A text story addressed through its slide and owning shape. Targets reject unknown fields, so
 * copy the ids off a read record rather than passing the record itself.
 */
export interface PptxStoryTarget extends PptxShapeTarget {
    storyId: string;
}
export interface PptxTextRange extends PptxStoryTarget {
    start: number;
    end: number;
}
/** An explicit range, or the one exact, case-sensitive, paragraph-local match of `text`. */
export type PptxTextTarget = ({
    kind: 'range';
} & PptxTextRange) | {
    kind: 'search';
    within: PptxStoryTarget;
    text: string;
};
/** Refuses the step unless its target currently reads exactly `text`. */
export interface PptxTextGuard {
    text: string;
}
/**
 * One batch step. Insertion, replacement and deletion stay inside one paragraph; formatting and
 * alignment may span paragraphs of one story. Geometry, fill and outline steps address shapes at
 * the top of a slide, and fill and outline only preset shapes.
 */
export type PptxEditStep = {
    op: 'insertText';
    target: PptxTextTarget;
    at: 'start' | 'end';
    text: string;
    expect?: PptxTextGuard;
} | {
    op: 'replaceText';
    target: PptxTextTarget;
    text: string;
    expect?: PptxTextGuard;
} | {
    op: 'deleteText';
    target: PptxTextTarget;
    expect?: PptxTextGuard;
} | {
    op: 'formatText';
    target: PptxTextTarget;
    patch: TextStylePatch;
    expect?: PptxTextGuard;
} | {
    op: 'setParagraphAlignment';
    target: PptxTextTarget;
    /** `null` restores the inherited alignment. */
    alignment: ParagraphAlignment | null;
    expect?: PptxTextGuard;
} | {
    op: 'setSlideNotes';
    target: PptxSlideTarget;
    text: string;
    expect?: PptxTextGuard;
} | {
    op: 'setSlideBackground';
    target: PptxSlideTarget;
    /** A solid or gradient fill; `null` removes the slide's own background. */
    background: ShapeFill | null;
    /** The effective background, as `snapshot()` reports it. */
    expect?: {
        background: ShapeFill | null;
    };
} | {
    op: 'setShapeRect';
    target: PptxShapeTarget;
    rect: ShapeRect;
    expect?: {
        rect: ShapeRect;
    };
} | {
    op: 'setShapeAdjust';
    target: PptxShapeTarget;
    adjustments: Record<string, number>;
} | {
    op: 'setShapeFill';
    target: PptxShapeTarget;
    /** A `#RRGGBB` solid fill; `null` removes the fill. */
    color: string | null;
    /** The authored fill, as `snapshot()` reports it. */
    expect?: {
        fill: ShapeFill | null;
    };
} | {
    op: 'setShapeStroke';
    target: PptxShapeTarget;
    stroke: ShapeStroke;
    /** The authored outline, as `snapshot()` reports it. */
    expect?: {
        outline: ShapeOutline | null;
    };
};
export interface PptxEditRequest {
    /** The version the targets were read at; a changed deck refuses with `stale-version`. */
    expectVersion: string;
    /** Defaults to `host`. */
    source?: PptxEditSource;
    /** Defaults to `separate`. */
    history?: PptxEditHistory;
    steps: readonly PptxEditStep[];
}
/** `read-only` comes only from an editor host that does not accept edits, never from a session. */
export type PptxEditFailureCode = 'stale-version' | 'missing-target' | 'ambiguous-target' | 'content-mismatch' | 'overlapping-steps' | 'unsupported' | 'invalid-step' | 'limit-exceeded' | 'read-only';
export type PptxEditTarget = PptxTextTarget | ({
    kind: 'story';
} & PptxStoryTarget) | ({
    kind: 'shape';
} & PptxShapeTarget) | ({
    kind: 'slide';
} & PptxSlideTarget);
export type PptxEditFailure = OperationFailure<PptxEditFailureCode, PptxEditTarget>;
export type PptxEditRefusal = OperationRefusal<PptxEditFailure>;
/** What one step did. Text targets describe the final state; a deletion reports its collapsed boundary. */
export interface PptxEditReceipt {
    stepIndex: number;
    changed: boolean;
    target: PptxEditTarget;
}
/** What one step would do at its resolved pre-batch target; it cannot be committed later. */
export interface PptxEditPreview {
    stepIndex: number;
    target: PptxEditTarget;
    wouldChange: boolean;
}
export type PptxEditResult = (EditSuccess<PptxEditReceipt> & {
    source: PptxEditSource;
    changedSlides: string[];
    changedStories: string[];
}) | PptxEditRefusal;
export type PptxValidationResult = ValidationSuccess<PptxEditPreview> | PptxEditRefusal;
export interface PptxReadRequest {
    /** Restricts the read to these slides, returned in deck order; every slide when omitted. */
    slideIds?: readonly string[];
}
/** The story offsets of a field result, such as a slide number. */
export interface PptxTextField {
    start: number;
    end: number;
    fieldType?: string;
}
export interface PptxParagraphText {
    list?: import('./structuredExport').PptxExportList | null;
    paragraphId: string;
    /** Story offsets of the paragraph's text; a following separator sits at `end`. */
    start: number;
    end: number;
    /** Story offsets of soft line breaks, which read as `\n` inside the paragraph. */
    lineBreaks: number[];
    /** Field results that text steps leave whole. */
    fields: PptxTextField[];
    /**
     * False when saving could not keep a field once the paragraph changes: the field is empty or
     * left the paragraph's unchanged leading or trailing text. Steps changing it refuse.
     */
    editable: boolean;
}
export interface PptxStoryText extends PptxStoryTarget {
    /** Paragraph texts joined by `\n`. */
    text: string;
    paragraphs: PptxParagraphText[];
}
export type PptxReadResult = {
    ok: true;
    version: string;
    slides: SlideSnapshot[];
    stories: PptxStoryText[];
} | PptxEditRefusal;
/** A slide, one shape and its descendants, or one story; `storyId` needs `shapeId`. */
export interface PptxFindScope {
    slideId: string;
    shapeId?: string;
    storyId?: string;
}
export interface PptxFindRequest {
    text: string;
    /** The whole deck when omitted. */
    within?: PptxFindScope;
    /** Defaults to 100; at most 10,000. */
    limit?: number;
}
/** One search hit; `range` is reusable as a `{ kind: 'range' }` target. */
export interface PptxFindMatch {
    text: string;
    range: PptxTextRange;
}
export type PptxFindResult = {
    ok: true;
    version: string;
    matches: PptxFindMatch[];
    truncated: boolean;
} | PptxEditRefusal;

// ---- index.d.ts ----
export { paintSlide, sizeCanvasForSlide } from './render/canvas';
export { decodePresentationImage, needsElementDecode, presentationImageBlob } from './render/image';
export type { PresentationImageDecodeOptions } from './render/image';
export { StaleProposalError } from './proposals';
export type { Proposal, ProposalAcceptance, ProposalChange, ProposalDiffSlide, ProposalEdit, ProposalPreview, ProposalTextChange } from './proposals';
export type { CanvasImageResolver, PaintSlideOptions, SlideCanvasLike } from './render/canvas';
export { slideToPng } from './render/png';
export type { SlideToPngOptions } from './render/png';
export { CollaborationError, CollaborationProvider } from './collaboration';
export type { CollaborationUser, CollaborationErrorCode, CollaborationErrorListener, CollaborationProviderOptions, AsyncCollaborationReplica, CollaborationReplica, CollaborationStatus, CollaborationStatusChange, CollaborationStatusListener, CollaborationTransport, CollaborationTransportEvent, CollaborationUpdateOrigin, PptxPresence, PptxPresenceCursor, PptxPresenceListener, PptxPresencePeer, PptxPresenceState, PptxPresenceUser, } from './collaboration';
export { PRESENCE_LABEL_DURATION_MS, presenceColorForClientId } from './collaboration';
export { exportPptxMarkdown, exportPptxStructured, initWasm, decodeFontFace, inspectPresentation, isWasmAvailable, isProposalsAvailable, openPresentation, renderPptxMarkdown, wasmVersion, engineBackend, } from './wasm/loader';
export { PptxExportError } from './structuredExport';
export type { ExportCompletion, ExportDiagnostic, ExportSeverity, MarkdownAnchor, PptxAnchor, PptxAnchorScope, PptxExportComment, PptxExportDiagnostic, PptxExportDiagnosticCode, PptxExportFailure, PptxExportFailureCode, PptxExportList, PptxExportMark, PptxExportNotes, PptxExportObject, PptxExportObjectKind, PptxExportOptions, PptxExportParagraph, PptxExportRefusal, PptxExportResult, PptxExportRun, PptxExportShape, PptxExportShapeKind, PptxExportSlide, PptxExportStory, PptxExportTable, PptxExportTableCell, PptxMarkdownContent, PptxMarkdownOptions, PptxSourceProvenance, PptxStructuredContent, PptxTextSpan, } from './structuredExport';
export type { InspectPresentationOptions, OpenPresentationOptions, SavePresentationOptions, PresentationHandle, UndoCaptureMode, WasmInitInput, EngineBackend, } from './wasm/loader';
export type { EditSuccess, OperationFailure, OperationRefusal, PptxEditFailure, PptxEditFailureCode, PptxEditHistory, PptxEditPreview, PptxEditReceipt, PptxEditRefusal, PptxEditRequest, PptxEditResult, PptxEditSource, PptxEditStep, PptxEditTarget, PptxFindMatch, PptxFindRequest, PptxFindResult, PptxFindScope, PptxParagraphText, PptxReadRequest, PptxReadResult, PptxShapeTarget, PptxSlideTarget, PptxStoryTarget, PptxStoryText, PptxTextField, PptxTextGuard, PptxTextRange, PptxTextTarget, PptxValidationResult, ValidationSuccess, } from './edits';
export type { BlipEffect, CaretStop, ColorTransform, ColorValue, CommentFlavor, CommentReceipt, CommentSnapshot, DeckSnapshot, EditProfile, GeometryPathCommand, HistoryProfile, HistoryResult, HitTestResult, ImageEffect, ImageCrop, ImagePrimitive, InheritedGeometry, LayoutProfile, Paint, ParagraphAlignment, ParagraphSnapshot, PictureDraft, PlaceholderPrimitive, PositionedGlyph, PositionedTextLine, PositionedTextRun, PresetShapeDraft, Profiled, ProfiledLayout, PptxFontFace, PptxCaretAnchor, PptxTextMatch, PptxTextSearchOptions, PrimitiveTransform, ShapeAdjustReceipt, ShapeDraft, ShapeFill, ShapeFillReceipt, ShapeGradient, ShapeKind, ShapeLineEnd, ShapeOutline, ShapePrimitive, ShapeReceipt, ShapeRect, ShapeSnapshot, ShapeStroke, ShapeStrokeReceipt, ShapeZOrderReceipt, SlideDisplayList, SlidePrimitive, SlideReceipt, SlideSnapshot, SlideLayoutSnapshot, HeaderFooterSettings, SlideHeaderFooterText, SlideDateText, SlideHeaderFooter, StorySnapshot, Stroke, StrokeEnd, TextBoxPrimitive, TextReceipt, TextParagraphDraft, TextRunSnapshot, TextStyle, TextStylePatch, TextStyleSnapshot, TransformReceipt, TableReceipt, UpdateEvent, } from './types';
/** @internal */
export { PptxPeerNotReadyError, PptxWorkerEditorDisposedError } from './session/peerHydrationError';
/** @internal */
export { openPresentationSession } from './session/client';
/** @internal */
export { createPptxWorkerEditorSession } from './session/editorSession';

// ---- proposals.d.ts ----
import type { DeckSnapshot, ParagraphAlignment, ShapeRect, ShapeSnapshot, ShapeStroke, SlideDisplayList, TextStyleSnapshot } from './types';
export type ProposalEdit = {
    type: 'replaceText';
    storyId: string;
    start: number;
    end: number;
    text: string;
    style?: Partial<TextStyleSnapshot> | null;
} | {
    type: 'formatText';
    storyId: string;
    start: number;
    end: number;
    patch: Partial<TextStyleSnapshot>;
} | {
    type: 'setParagraphAlignment';
    storyId: string;
    start: number;
    end: number;
    alignment: ParagraphAlignment | 'justLow' | 'dist' | 'thaiDist' | null;
} | {
    type: 'setShapeRect';
    slideId: string;
    shapeId: string;
    rect: ShapeRect;
} | {
    type: 'setShapeFill';
    slideId: string;
    shapeId: string;
    color: string | null;
} | {
    type: 'setShapeStroke';
    slideId: string;
    shapeId: string;
    stroke: {
        [K in keyof ShapeStroke]: ShapeStroke[K] | null;
    };
} | {
    type: 'setShapeAdjust';
    slideId: string;
    shapeId: string;
    adjustments: Record<string, number>;
} | {
    type: 'setSlideNotes';
    slideId: string;
    text: string;
};
export interface ProposalChange {
    slideId: string;
    shapeId: string | null;
    before: ShapeSnapshot | null;
    after: ShapeSnapshot | null;
    oldText: string;
    newText: string;
}
export interface Proposal {
    id: string;
    agentId: string;
    note: string | null;
    edits: ProposalEdit[];
    changes: ProposalChange[];
    staleTargets: string[];
}
export interface ProposalPreview {
    proposal: Proposal;
    snapshot: DeckSnapshot;
}
export interface ProposalTextChange {
    storyId: string;
    start: number;
    end: number;
    kind: 'insertion' | 'deletion';
}
export interface ProposalDiffSlide extends ProposalPreview {
    frame: SlideDisplayList;
    textChanges: ProposalTextChange[];
}
export interface ProposalAcceptance {
    proposalId: string;
    applied: boolean;
    snapshot: DeckSnapshot;
}
export declare class StaleProposalError extends Error {
    readonly targets: string[];
    constructor(targets: string[]);
}

// ---- render/browserImage.d.ts ----
export interface PresentationImageDecodeOptions {
    maxDimension?: number;
}
export declare function imageDecodeScale(source: CanvasImageSource): {
    x: number;
    y: number;
};
export declare function setImageDecodeScale(source: CanvasImageSource, scale: {
    x: number;
    y: number;
}): void;
export declare function browserPresentationImageBlob(bytes: Uint8Array): Blob;
export declare function needsElementDecode(blob: Blob): boolean;
export declare function decodeBrowserPresentationImage(bytes: Uint8Array, errorMessage: string, options?: PresentationImageDecodeOptions): Promise<CanvasImageSource>;
export declare function decodeBrowserImageBlob(blob: Blob, errorMessage: string, options?: PresentationImageDecodeOptions): Promise<CanvasImageSource>;

// ---- render/canvas.d.ts ----
import type { ImageEffect, SlideDisplayList } from '../types';
import type { ProposalTextChange } from '../proposals';
export type CanvasImageResolver = ((assetId: string) => CanvasImageSource | Promise<CanvasImageSource | null> | null) & {
    acquire?(): CanvasImageResolver;
    release?(): void;
};
export interface PaintSlideOptions {
    resolveImage?: CanvasImageResolver;
    maxShadowPixels?: number;
    textChanges?: readonly ProposalTextChange[];
}
export interface SlideCanvasLike {
    width: number;
    height: number;
    style: {
        width: string;
        height: string;
    };
}
/** The backing store a slide needs: a fractional extent covers its last pixel. */
export declare function slideBackingStore(list: Pick<SlideDisplayList, 'width' | 'height'>, dpr: number, scale?: number): {
    width: number;
    height: number;
};
export declare function sizeCanvasForSlide(canvas: SlideCanvasLike, list: Pick<SlideDisplayList, 'width' | 'height'>, dpr: number, scale?: number): void;
export declare function paintSlide(ctx: CanvasRenderingContext2D, list: SlideDisplayList, dpr?: number, scale?: number, options?: PaintSlideOptions): Promise<void>;
/** `getImageData` hands back straight alpha, which is what these are defined on. */
export declare function applyImageEffects(data: Uint8ClampedArray, effects: ImageEffect[]): void;

// ---- render/image.d.ts ----
import type { PresentationImageDecodeOptions } from './browserImage';
export { imageDecodeScale, needsElementDecode, setImageDecodeScale } from './browserImage';
export type { PresentationImageDecodeOptions } from './browserImage';
export declare function presentationImageBlob(bytes: Uint8Array): Blob;
export declare function decodePresentationImage(bytes: Uint8Array, errorMessage: string, options?: PresentationImageDecodeOptions): Promise<CanvasImageSource>;

// ---- render/png.d.ts ----
import type { SlideDisplayList } from '../types';
import { type CanvasImageResolver } from './canvas';
export interface SlideToPngOptions {
    /** Output scale, e.g. 2 for hidpi. */
    scale?: number;
    resolveImage?: CanvasImageResolver;
}
/**
 * Rasterize one slide to a PNG blob through the same canvas painter the editor
 * draws with, so an export matches what is on screen. The server-side twin is
 * the `betteroffice-pptx-raster` crate.
 */
export declare function slideToPng(list: SlideDisplayList, options?: SlideToPngOptions): Promise<Blob>;

// ---- session/client.d.ts ----
import { type SessionTransport } from '../../../../shared/office-session/transport';
import type { Promisified, SessionFailure } from '../../../../shared/office-session/types';
import type { OpenPresentationOptions } from '../wasm/loader';
import { type EditorBaseline, type PptxWorkerEditorFrame } from './replay';
import { type PresentationFrame, type PresentationSessionEvents, type PresentationSessionMethods, type PresentationSessionState } from './methods';
/**
 * Options for opening a presentation in a dedicated worker.
 * @experimental
 */
export interface OpenPresentationSessionOptions extends OpenPresentationOptions {
    worker?: () => Worker;
    wasm?: ArrayBuffer | WebAssembly.Module;
}
/**
 * Async presentation access and its current local projection.
 * @experimental
 */
export interface PresentationSession {
    readonly state: PresentationSessionState;
    readonly call: Promisified<Omit<PresentationSessionMethods, 'open' | 'dispose' | 'frame'>> & {
        /**
         * @experimental A newer frame call replaces a queued one, which rejects with an error
         * named SessionSuperseded.
         */
        frame(slideIndex: number): Promise<PresentationFrame>;
    };
    save(): Promise<Uint8Array>;
    on<K extends keyof PresentationSessionEvents>(name: K, listener: (payload: PresentationSessionEvents[K]) => void): () => void;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    readonly failure: SessionFailure | undefined;
    dispose(): Promise<void>;
}
export interface PresentationEditorConnection {
    readonly failure: SessionFailure | undefined;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    begin(): Promise<EditorBaseline>;
    frame(slideId: string): Promise<PptxWorkerEditorFrame>;
    dispose(): Promise<void>;
}
export declare function createPresentationEditorClient(bytes: Uint8Array | ArrayBuffer, options: OpenPresentationSessionOptions, transport: SessionTransport): PresentationEditorConnection;
/**
 * Opens a worker session, transferring copies so caller buffers stay usable.
 * @experimental
 */
export declare function openPresentationSession(bytes: Uint8Array | ArrayBuffer, options?: OpenPresentationSessionOptions): Promise<PresentationSession>;
/** Creates a presentation session over an internal transport. */
export declare function createPresentationSession(bytes: Uint8Array | ArrayBuffer, options: OpenPresentationSessionOptions, transport: SessionTransport): Promise<PresentationSession>;

// ---- session/editPeer.d.ts ----
import type { PresentationEditorConnection } from './client';
import type { DeckSnapshot, HitTestResult } from '../types';
import { type PresentationHandle, type PresentationPeerHandle } from '../wasm/loader';
import { PptxWorkerEditorFailedError } from './peerHydrationError';
import { type PptxWorkerEditorOperation, type PptxWorkerEditorReply } from './replay';
import type { PresentationSessionState } from './methods';
declare const readMethods: readonly ["snapshot", "story", "anchorCaret", "resolveCaretAnchor", "version", "readContent", "findText", "validateEdits", "searchText", "isProposalsAvailable", "listProposals", "layoutSlide", "hitTest", "mediaBytes", "comments", "canUndo", "canRedo", "undoCaptureMode", "onUpdate"];
export interface PptxWorkerEditorInteractionAccess {
    snapshot(): {
        snapshot: DeckSnapshot;
        keys: Record<string, string>;
    };
    key(index: number): string;
    activate(slideId: string, key: string): boolean;
    hitTest(slideId: string, x: number, y: number): HitTestResult | null;
}
export type PptxWorkerEditorAccess = Pick<PresentationHandle, typeof readMethods[number] | PptxWorkerEditorOperation['method'] | 'clientId'> & {
    readonly interaction: PptxWorkerEditorInteractionAccess;
};
export interface PresentationEditPeer {
    readonly access: PptxWorkerEditorAccess;
    readonly sequence: number;
    readonly acknowledgedSequence: number;
    readonly version: string;
    apply(op: PptxWorkerEditorOperation): PptxWorkerEditorReply;
    flush(): Promise<void>;
    saveAsync(): Promise<Uint8Array>;
    dispose(): void;
}
export declare function createPresentationEditPeer(options: {
    peer: PresentationPeerHandle;
    session: PresentationEditorConnection;
    guard(): void;
    onFailure(error: PptxWorkerEditorFailedError): void;
    onProgress(sequence: number, acknowledgedSequence: number, projection?: PresentationSessionState): void;
}): PresentationEditPeer;
export {};

// ---- session/editorSession.d.ts ----
import { type SessionTransport } from '../../../../shared/office-session/transport';
import { type OpenPresentationSessionOptions } from './client';
import { type PptxWorkerEditorAccess } from './editPeer';
import type { PresentationSessionState } from './methods';
import { PptxWorkerEditorFailedError } from './peerHydrationError';
import { type PptxWorkerEditorFrame, type PptxWorkerEditorOperation, type PptxWorkerEditorReply } from './replay';
export type PptxWorkerEditorStage = 'opening' | 'hydrating' | 'ready' | 'failed' | 'recovering' | 'disposed';
export interface PptxWorkerEditorState {
    readonly stage: PptxWorkerEditorStage;
    readonly sequence: number;
    readonly acknowledgedSequence: number;
    readonly version?: string;
    readonly projection?: PresentationSessionState;
    readonly initialFrame?: PptxWorkerEditorFrame;
    readonly failure?: PptxWorkerEditorFailedError;
}
export interface PptxWorkerEditorOptions extends OpenPresentationSessionOptions {
    collaboration?: unknown;
    signal?: AbortSignal;
    onError?(error: PptxWorkerEditorFailedError): void;
}
export interface PptxWorkerEditorRecovery {
    bytes: Uint8Array;
    recovery: true;
}
export interface PptxWorkerEditorSession {
    readonly state: PptxWorkerEditorState;
    readonly hydrated: boolean;
    readonly failure: PptxWorkerEditorFailedError | undefined;
    subscribe(listener: (state: PptxWorkerEditorState) => void): () => void;
    whenHydrated(): Promise<void>;
    handleAsync(): Promise<PptxWorkerEditorAccess>;
    apply(op: PptxWorkerEditorOperation): PptxWorkerEditorReply;
    flush(): Promise<void>;
    saveAsync(): Promise<Uint8Array>;
    frame(slideId: string): Promise<PptxWorkerEditorFrame>;
    recoverySave(): Promise<PptxWorkerEditorRecovery>;
    dispose(): Promise<void>;
}
export declare function createPptxWorkerEditorSession(bytes: Uint8Array | ArrayBuffer, options?: PptxWorkerEditorOptions): PptxWorkerEditorSession;
export declare function createPresentationEditorSession(bytes: Uint8Array | ArrayBuffer, options: PptxWorkerEditorOptions, transport: SessionTransport): PptxWorkerEditorSession;

// ---- session/methods.d.ts ----
import type { MethodPolicies, SessionState } from '../../../../shared/office-session';
import type { PptxEditRequest, PptxEditResult, PptxFindRequest, PptxFindResult, PptxReadRequest, PptxReadResult, PptxValidationResult } from '../edits';
import type { SlideDisplayList } from '../types';
import type { OpenPresentationOptions } from '../wasm/loader';
/** @experimental */
export interface PresentationFrame {
    displayList: SlideDisplayList;
    version: string;
    epoch: number;
    slideIndex: number;
    media: ReadonlyMap<string, Uint8Array>;
}
/** @experimental */
export type PresentationWireFrame = Omit<PresentationFrame, 'displayList' | 'media'> & {
    displayList: ArrayBuffer;
    media: {
        assetId: string;
        bytes: ArrayBuffer;
    }[];
};
/**
 * A transferable font face.
 * @experimental
 */
export interface PresentationSessionFont {
    family: string;
    bytes: ArrayBuffer;
    bold?: boolean;
    italic?: boolean;
}
/**
 * Structured-cloneable presentation open options.
 * @experimental
 */
export interface PresentationSessionOpenOptions extends Omit<OpenPresentationOptions, 'fonts' | 'fallbackFonts'> {
    fonts?: readonly PresentationSessionFont[];
    fallbackFonts?: readonly PresentationSessionFont[];
    wasm?: ArrayBuffer | WebAssembly.Module;
}
/**
 * Slide metadata without layout; indices are zero-based.
 * @experimental
 */
export interface PresentationSlideSummary {
    id: string;
    index: number;
    name: string | null;
    layoutPartPath: string | null;
}
/**
 * Presentation projection; version counts applied batches and size is in EMU.
 * @experimental
 */
export type PresentationSessionState = Omit<SessionState, 'stage'> & {
    format: 'pptx';
    stage: 'preview' | 'ready' | 'failed';
    slides: PresentationSlideSummary[];
    size: {
        width: number;
        height: number;
    };
};
/**
 * Presentation change notifications after applied batches.
 * @experimental
 */
export interface PresentationSessionEvents {
    changed: {
        version: number;
        dirty: boolean;
    };
}
/**
 * Presentation RPC methods; engine version tokens remain strings.
 * @experimental
 */
export type PresentationSessionMethods = {
    open(bytes: ArrayBuffer, options?: PresentationSessionOpenOptions): PresentationSessionState;
    version(): string;
    readContent(request?: PptxReadRequest): PptxReadResult;
    findText(request: PptxFindRequest): PptxFindResult;
    validateEdits(request: PptxEditRequest): PptxValidationResult;
    applyEdits(request: PptxEditRequest): PptxEditResult;
    frame(slideIndex: number): PresentationWireFrame;
    slides(): PresentationSlideSummary[];
    slideSize(): {
        width: number;
        height: number;
    };
    save(): ArrayBuffer;
    dispose(): void;
};
/** Exhaustive RPC names used by the session client. */
export declare const PRESENTATION_SESSION_METHODS: {
    open: true;
    version: true;
    readContent: true;
    findText: true;
    validateEdits: true;
    applyEdits: true;
    frame: true;
    slides: true;
    slideSize: true;
    save: true;
    dispose: true;
};
/** Non-reorderable policies preserve call order across all lanes. */
export declare const PRESENTATION_SESSION_POLICIES: MethodPolicies<PresentationSessionMethods>;

// ---- session/peerHydrationError.d.ts ----
export declare class PptxPeerHydrationError extends Error {
    readonly code: string;
    readonly cause?: unknown | undefined;
    readonly refusal: {
        code: string;
    };
    constructor(code: string, message: string, cause?: unknown | undefined);
}
export declare class PptxPeerNotReadyError extends Error {
    readonly code = "peer-not-ready";
    constructor();
}
export declare class PptxWorkerEditorFailedError extends Error {
    readonly cause: unknown;
    readonly code = "editor-failed";
    constructor(cause: unknown);
}
export declare class PptxWorkerEditorDisposedError extends Error {
    readonly code = "editor-disposed";
    constructor();
}
export declare class PptxWorkerEditorCollaborationError extends Error {
    readonly code = "collaboration-unavailable";
    constructor();
}

// ---- session/replay.d.ts ----
import type { MethodPolicies } from '../../../../shared/office-session/types';
import type { SessionClient } from '../../../../shared/office-session/client';
import type { PresentationEditorConnection } from './client';
import type { PresentationReplayEnvelope, PresentationReplayOp, PresentationReplayReply } from '../wasm/loader';
import type { PresentationFrame, PresentationSessionOpenOptions, PresentationSessionState, PresentationWireFrame } from './methods';
export type PptxWorkerEditorOperation = Exclude<PresentationReplayOp, {
    method: 'insertTableRow' | 'deleteTableRow' | 'insertTableColumn' | 'deleteTableColumn';
}>;
export type PptxWorkerEditorReply = PresentationReplayReply;
export type PptxWorkerEditorFrame = PresentationFrame & {
    slideId: string;
    sequence: number;
};
export type EditorPosition = {
    version: string;
    sequence: number;
};
export type EditorBaseline = EditorPosition & {
    state: PresentationSessionState;
    hydration: string;
    module: WebAssembly.Module;
};
export type EditorReplayReply = PresentationReplayReply & {
    projection: PresentationSessionState;
};
export type PresentationEditorMethods = {
    beginEditor(bytes: ArrayBuffer, options: PresentationSessionOpenOptions): EditorBaseline;
    attachPeer(version: string, sequence: number): EditorPosition;
    replay(envelope: PresentationReplayEnvelope): EditorReplayReply;
    flush(sequence: number, version: string): EditorPosition;
    editorFrame(slideId: string): PresentationWireFrame & {
        slideId: string;
        sequence: number;
    };
    editorSave(sequence: number, version: string): ArrayBuffer;
};
export declare const presentationEditorSessionInternals: WeakMap<PresentationEditorConnection, {
    client: SessionClient<PresentationEditorMethods, {}>;
    peerAttached: boolean;
}>;
export declare const PRESENTATION_EDITOR_METHODS: {
    beginEditor: true;
    attachPeer: true;
    replay: true;
    flush: true;
    editorFrame: true;
    editorSave: true;
};
export declare const PRESENTATION_EDITOR_POLICIES: MethodPolicies<PresentationEditorMethods>;
export declare const PRESENTATION_REPLAY_MUTATORS: {
    insertText: true;
    deleteText: true;
    formatText: true;
    setStoryParagraphs: true;
    insertParagraphBreak: true;
    setParagraphAlignment: true;
    insertSlide: true;
    duplicateSlide: true;
    deleteSlide: true;
    moveSlide: true;
    setSlideNotes: true;
    setSlideLayout: true;
    addTextBox: true;
    addShape: true;
    addPicture: true;
    removeShape: true;
    moveShape: true;
    resizeShape: true;
    setShapeRect: true;
    setShapeFill: true;
    setShapeStroke: true;
    setShapeAdjust: true;
    bringShapeToFront: true;
    sendShapeToBack: true;
    bringShapeForward: true;
    sendShapeBackward: true;
    addComment: true;
    replyToComment: true;
    setCommentStatus: true;
    setCommentPosition: true;
    removeComment: true;
    setCommentFlavor: true;
    propose: true;
    acceptProposal: true;
    rejectProposal: true;
    applyEdits: true;
    addUndoBoundary: true;
    undo: true;
    redo: true;
};
export declare function sameReplayValue(left: unknown, right: unknown): boolean;
export declare function matchingReplayReply(local: PresentationReplayReply, remote: PresentationReplayReply): boolean;

// ---- shared/collaboration-pipe/main.d.ts ----
import { type PipeCursor, type PipeErrorCode, type PipeOp, type PipeOutput, type PipePeer, type PipeProviderOptions } from './protocol';
type Status = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
type TransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
} | {
    type: 'drain';
};
export interface PipeHostTransport {
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: TransportEvent) => void): () => void;
}
/** One provider's line to the worker that runs it. */
export interface CollaborationPipeEndpoint<C = PipeCursor> {
    /** Runs `ops` in the worker, in order after earlier posts; settles once the worker has run them. */
    post(ops: PipeOp<C>[], transfer: ArrayBuffer[]): Promise<unknown>;
    close(): void;
}
/** What a worker replica offers a provider that can run in its worker. */
export interface CollaborationPipeChannel<C = PipeCursor, P = PipePeer> {
    /** `receive` takes the worker's output for this provider; `lost` reports the worker is gone for it. */
    open(receive: (out: readonly PipeOutput<P>[]) => void, lost: (reason: string) => void): CollaborationPipeEndpoint<C>;
}
export interface PipedCollaborationOptions<E, P = PipePeer> {
    worker: PipeProviderOptions;
    maxPendingBytes: number;
    hasUser: boolean;
    /** Moves a worker peer's timestamps onto this context's clock; omitted when they are wall-clock times. */
    rebasePeer?(peer: P, offset: number): P;
    makeError(code: PipeErrorCode, message: string, cause?: unknown): E;
    /** `cause` itself when it already is an `E`, else a new `E` with `context` and the cause's message. */
    normalize(code: PipeErrorCode, context: string, cause: unknown): E;
}
/**
 * The main-thread side of a provider whose protocol runs in the document worker. It owns the host transport's
 * lifecycle as a provider does, forwards the transport's events with frame bytes untouched, one post per burst,
 * and sends what the worker answers. Worker output older than main's latest lifecycle call is dropped.
 */
export declare class PipedCollaboration<E extends Error, C = PipeCursor, P = PipePeer> {
    private readonly transport;
    private readonly options;
    private readonly endpoint;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly peerListeners;
    private nextListenerId;
    private connectionStatus;
    private isSynced;
    private peerList;
    private seq;
    /** The op of main's latest lifecycle call; older worker output describes a connection main has left. */
    private floor;
    /** The op of the transport's latest open or close; sends, status and peers from before it describe a connection the host ended. */
    private sendFloor;
    /** Received frame bytes queued or in flight to the worker. */
    private inboundBytes;
    private flushingFrames;
    private queue;
    private posting;
    private flushQueued;
    private heldTimer;
    /** The queue holds an op that should not wait for held updates. */
    private urgent;
    private unsubscribeTransport;
    private subscription;
    private wantsConnection;
    private connectAttempt;
    private leaveFrame;
    private readonly pendingFrames;
    private queuedBytes;
    private isDestroyed;
    private closed;
    constructor(channel: CollaborationPipeChannel<C, P>, transport: PipeHostTransport, options: PipedCollaborationOptions<E, P>);
    get status(): Status;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly P[];
    setCursor(cursor: C | null, broadcast: boolean): void;
    onPeers(listener: (peers: readonly P[]) => void): () => void;
    onStatus(listener: (change: {
        status: Status;
        synced: boolean;
    }) => void): () => void;
    onError(listener: (error: E) => void): () => void;
    connect(): void;
    disconnect(): void;
    destroy(): void;
    private transportError;
    /** Leaves the connection main holds: drops the subscription, disconnects, clears queued bytes and peers. */
    private release;
    /** Fails main's connection with `error`, and the worker's with it. */
    private fail;
    private sendLeave;
    private enqueue;
    private schedule;
    private flush;
    private forward;
    private receive;
    private send;
    private queueFrame;
    private flushPending;
    private clearQueue;
    /**
     * The worker no longer runs this provider. Hands the host transport back at once; reports the loss a task later,
     * so a host that destroys the provider on `onReplica(null)` hears nothing.
     */
    private lose;
    private close;
    private clearPeers;
    private setStatus;
    private emitPeers;
    private report;
}
export {};

// ---- shared/collaboration-pipe/protocol.d.ts ----
/**
 * Wire format between a main-thread collaboration facade and the provider its document worker runs.
 * Main owns the host transport's lifecycle and forwards its bytes untouched; the worker runs the protocol and
 * answers with frames to send, state and the transport calls its own failures need.
 */
export type PipeErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export interface PipeError {
    code: PipeErrorCode;
    message: string;
}
export interface PipeUser {
    name: string;
    color?: string;
}
export interface PipeProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    syncChunkBytes?: number;
    maxSyncBytes?: number;
    user?: PipeUser;
}
/** The docx cursor; the default cursor type `C` of the pipe types. */
export interface PipeCursor {
    story: string;
    anchor: Uint8Array;
    head: Uint8Array;
}
export type PipeTransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: ArrayBuffer;
} | {
    type: 'close';
    reason?: string;
};
/**
 * Main to worker, in order; `seq` increases. `connect`, `disconnect`, `destroy` and `fail` replay what main already
 * did to the host transport, so the worker sends nothing and asks for no transport call while it runs them.
 */
export type PipeOp<C = PipeCursor> = {
    op: 'create';
    seq: number;
    options: PipeProviderOptions;
} | {
    op: 'connect';
    seq: number;
} | {
    op: 'disconnect';
    seq: number;
} | {
    op: 'destroy';
    seq: number;
} | {
    op: 'fail';
    seq: number;
    error: PipeError;
} | {
    op: 'cursor';
    seq: number;
    cursor: C | null;
    broadcast: boolean;
} | {
    op: 'event';
    seq: number;
    event: PipeTransportEvent;
};
/** The docx peer; the default peer type `P` of the pipe types. */
export interface PipePeer {
    clientId: number;
    clock: number;
    user: {
        name: string;
        color: string;
    };
    cursor: PipeCursor | null;
    lastSeenAt: number;
    cursorMovedAt: number;
    inferredCursor: {
        clientId: number;
        story: string;
        paraId: string;
        endOffset: number;
    } | null;
}
/**
 * Worker to main, in order. `through` is the last op the worker had run; main drops output older than its own
 * latest lifecycle call.
 */
export type PipeOutput<P = PipePeer> = {
    do: 'unsubscribe';
    through: number;
} | {
    do: 'disconnect';
    through: number;
} | {
    do: 'send';
    data: ArrayBuffer;
    through: number;
}
/** The awareness leave main sends when the host disconnects first; null while there is none to send. */
 | {
    do: 'leave';
    data: ArrayBuffer | null;
    through: number;
} | {
    do: 'status';
    status: 'disconnected' | 'connecting' | 'connected' | 'destroyed';
    synced: boolean;
    through: number;
} | {
    do: 'peers';
    peers: P[];
    timeOrigin: number;
    through: number;
} | {
    do: 'error';
    error: PipeError;
    through: number;
}
/** The worker has no provider under this id. */
 | {
    do: 'lost';
};
export interface PipeOutputBatch<P = PipePeer> {
    provider: number;
    out: PipeOutput<P>[];
}
export interface PipeCommandPayload<C = PipeCursor> {
    provider: number;
    ops: PipeOp<C>[];
}
export interface PipePayloadShape {
    /** Checks a non-null cursor; its byte fields stay within `maxFrameBytes`. */
    validCursor(cursor: unknown, maxFrameBytes: number): boolean;
    /** The keys a `create` op's options may carry. */
    optionKeys: ReadonlySet<string>;
}
/** Buffers a pipe payload holds, for transfer. */
export declare function pipeOpBuffers<C>(ops: readonly PipeOp<C>[]): ArrayBuffer[];
export declare function ownedBuffer(bytes: Uint8Array): ArrayBuffer;
export declare const PIPE_OPTION_KEYS: ReadonlySet<string>;
/** The docx cursor check. */
export declare function validPipeCursor(cursor: unknown, maxFrameBytes: number): boolean;
/** Shape check for a pipe command payload; frames past `maxFrameBytes` fail it. */
export declare function validPipePayload<C = PipeCursor>(payload: unknown, maxFrameBytes: number, shape: PipePayloadShape): payload is PipeCommandPayload<C>;
/** Epoch time of this context's provider clock zero: performance.now() when present, else Date.now(). */
export declare function clockOrigin(): number;

// ---- shared/collaboration-pipe/worker.d.ts ----
import { type PipeCommandPayload, type PipeCursor, type PipeError, type PipeOutput, type PipeOutputBatch, type PipePeer, type PipeProviderOptions } from './protocol';
type UpdateOrigin = 'local' | 'remote';
type TransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
};
/** The synchronous replica the worker's provider runs against; merges wait for the end of the burst. */
export interface PipeReplica {
    readonly clientId: number;
    encodeStateVector(): Uint8Array;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    encodeStateAsUpdateChunks?(remoteStateVector: Uint8Array, maxChunkBytes: number): Uint8Array[];
    applyUpdate(update: Uint8Array): unknown;
    applyUpdateChunks(updates: readonly Uint8Array[]): unknown;
    onUpdate(listener: (update: Uint8Array, origin: UpdateOrigin) => void): () => void;
}
export interface PipeTransport {
    connect(): void;
    disconnect(): void;
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: TransportEvent) => void): () => void;
}
export interface PipeProvider<C = PipeCursor, P = PipePeer> {
    connect(): void;
    disconnect(): void;
    destroy(): void;
    setCursor(cursor: C | null, broadcast?: boolean): void;
    onStatus(listener: (change: {
        status: Extract<PipeOutput, {
            do: 'status';
        }>['status'];
        synced: boolean;
    }) => void): () => void;
    onError(listener: (error: {
        code: string;
        message: string;
    }) => void): () => void;
    onPeers(listener: (peers: readonly P[]) => void): () => void;
}
/** A merge the provider handed over: settled once the burst's updates merge, failed when they do not. */
export interface PipeArrival {
    settle(applied: unknown): void;
    fail(cause: unknown): void;
}
export type PipeMergeResult = {
    applied: unknown;
} | {
    failed: unknown;
};
export interface CollaborationPipeHostOptions<Prov extends PipeProvider<C, P>, C = PipeCursor, P = PipePeer> {
    clientId: number;
    read: Pick<PipeReplica, 'encodeStateVector' | 'encodeStateAsUpdate' | 'encodeStateAsUpdateChunks'>;
    onUpdate: PipeReplica['onUpdate'];
    /**
     * Merges a burst's updates in one change: `single` is one peer update, which may report a typing inference
     * as `applied`. Returns `failed` for updates the replica refused without changing; throws when it cannot go on.
     */
    merge(updates: Uint8Array[], single: boolean): PipeMergeResult;
    createProvider(replica: PipeReplica, transport: PipeTransport, options: PipeProviderOptions): Prov;
    arrive(provider: Prov): PipeArrival;
    /** The awareness leave `provider` would send now, by the clock it carries; null when it would send none. */
    leave(provider: Prov): {
        clock: number;
        frame: Uint8Array;
    } | null;
    /** The error object a transport `error` event carries for `error`. */
    transportError(error: PipeError): unknown;
    /** Output became pending outside `run`, from the providers' timers. */
    onOutput?(): void;
}
export interface CollaborationPipeHost<C = PipeCursor, P = PipePeer> {
    /** Runs one burst of ops: transport events in order, then one merge of every update they carried. */
    run(payload: PipeCommandPayload<C>): void;
    /** Pending output, in order, and the buffers it holds. */
    drain(): {
        batches: PipeOutputBatch<P>[];
        transfer: ArrayBuffer[];
    } | undefined;
    /** Forgets every provider without sending anything. */
    dispose(): void;
}
export declare function createCollaborationPipeHost<Prov extends PipeProvider<C, P>, C = PipeCursor, P = PipePeer>(options: CollaborationPipeHostOptions<Prov, C, P>): CollaborationPipeHost<C, P>;
export {};

// ---- shared/host-contracts/edits.d.ts ----
/**
 * Result envelopes shared by the formats' host edit APIs. Targets, steps, receipts and failure
 * codes stay format-owned; these shapes only fix how every format reports success and refusal.
 */
/** Why a host operation was refused, as data. The document is unchanged. */
export interface OperationFailure<Code extends string, Target> {
    code: Code;
    /** The request step that failed, when one did. */
    stepIndex?: number;
    /** The earlier step an `overlapping-steps` failure conflicts with. */
    conflictingStepIndex?: number;
    /** The failing step's target as requested or resolved. */
    target?: Target;
    message: string;
}
/** A policy refusal carrying the version the document is still at. */
export interface OperationRefusal<Failure> {
    ok: false;
    version: string;
    failure: Failure;
}
/**
 * An accepted edit request. `applied` is false when every step was a no-op; the version then
 * stays `baseVersion` and no history or change notification was produced.
 */
export interface EditSuccess<Receipt> {
    ok: true;
    baseVersion: string;
    version: string;
    applied: boolean;
    /** One receipt per request step, in request order. */
    receipts: Receipt[];
}
/** A validated edit request. Validation reserves nothing and cannot be committed later. */
export interface ValidationSuccess<Preview> {
    ok: true;
    baseVersion: string;
    wouldApply: boolean;
    previews: Preview[];
}

// ---- shared/host-contracts/exports.d.ts ----
/**
 * Shapes shared by the formats' read-only structured exports. Content trees, anchors, story or
 * sheet selectors, diagnostic codes and limits stay format-owned; these only fix how every format
 * reports diagnostics, truncation and Markdown markers.
 */
export type ExportSeverity = 'info' | 'warning' | 'error';
/** Something an export omitted or could not represent, anchored where it applies. */
export interface ExportDiagnostic<Code extends string, Anchor> {
    code: Code;
    severity: ExportSeverity;
    anchor: Anchor | null;
    message: string;
}
/** Whether an export stopped at its limits before the end of the content. */
export interface ExportCompletion {
    truncated: boolean;
}
/** A Markdown marker comment and the anchor of the content it precedes. */
export interface MarkdownAnchor<Anchor> {
    marker: string;
    anchor: Anchor;
}

// ---- shared/office-session/client.d.ts ----
import { type HostMessage } from './protocol';
import type { SessionTransport } from './transport';
import { SessionFailure, type Promisified, type SessionEvents, type SessionMethods } from './types';
export declare const OFFICE_SESSION_SILENCE_MS = 60000;
export declare function requestWasmCompile(transport: SessionTransport): void;
export interface SessionClientOptions<M extends SessionMethods> {
    methods: {
        readonly [K in keyof M]-?: true;
    };
    silenceMs?: number;
    now?(): number;
    timer?(callback: () => void, ms: number): () => void;
    onWasmModule?(url: string, module: WebAssembly.Module): void;
    onWasmModuleMessage?(message: Extract<HostMessage, {
        kind: 'wasm-module';
    }>): void;
}
export interface SessionClient<M extends SessionMethods, E extends SessionEvents> {
    readonly call: Promisified<M>;
    callWithTransfer<K extends keyof M & string>(method: K, args: Parameters<M[K]>, transfer: Transferable[]): Promise<Awaited<ReturnType<M[K]>>>;
    on<K extends keyof E & string>(name: K, listener: (payload: E[K]) => void): () => void;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    readonly failure: SessionFailure | undefined;
    dispose(): Promise<void>;
}
/** Creates a typed client over an existing transport. */
export declare function createSessionClient<M extends SessionMethods, E extends SessionEvents>(transport: SessionTransport, options: SessionClientOptions<M>): SessionClient<M, E>;

// ---- shared/office-session/fontDecoder.d.ts ----
export type FontDecoder = (bytes: ArrayBuffer | Uint8Array, options: {
    codec: string;
    dict?: ArrayBuffer | Uint8Array;
    maxOutput?: number;
}) => Promise<{
    font: ArrayBuffer;
    payload?: ArrayBuffer;
}>;
export type FontDecodeResult = Awaited<ReturnType<FontDecoder>>;
export type FontDecoderMethods = {
    decodeFontFace(bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], wasm?: ArrayBuffer | WebAssembly.Module): FontDecodeResult;
};
export declare function rpcFontDecoder(call: (bytes: ArrayBuffer, options: Parameters<FontDecoder>[1], transfer: Transferable[]) => Promise<FontDecodeResult>): FontDecoder;
export declare function fontBytes(bytes: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer>;
export declare function copyDecodedFace(face: {
    font: Uint8Array;
    payload?: Uint8Array;
}): FontDecodeResult;

// ---- shared/office-session/host.d.ts ----
import { type TaskExecutor } from './scheduler';
import type { SessionTransport } from './transport';
import { SessionFailure, type MethodHandlers, type MethodPolicies, type SessionEvents, type SessionMethods, type SessionScheduler } from './types';
export interface SessionHostOptions<M extends SessionMethods, C> {
    handlers: MethodHandlers<M, C>;
    policies: MethodPolicies<M>;
    context: C;
    executor?: TaskExecutor;
    onDispose?(): void;
    /** Names a trap's cause, raw or already wrapped by a handler; without one it fails the session as `trap`. */
    trapFailure?(trap: SessionFailure): SessionFailure | undefined;
}
export interface SessionHost<E extends SessionEvents> {
    readonly scheduler: SessionScheduler;
    emit<K extends keyof E & string>(name: K, payload: E[K], transfer?: Transferable[]): void;
    /** Tells the client the running call is still working; safe to call from a long synchronous engine call. */
    progress(): void;
}
/** Runs method handlers serially through the resident scheduler. */
export declare function createSessionHost<M extends SessionMethods, E extends SessionEvents, C>(transport: SessionTransport, options: SessionHostOptions<M, C>): SessionHost<E>;

// ---- shared/office-session/index.d.ts ----
export * from './types';
export * from './protocol';
export * from './transport';
export * from './client';
export * from './host';
export * from './scheduler';

// ---- shared/office-session/protocol.d.ts ----
import type { SessionFailureCode } from './types';
export type ClientMessage = {
    protocol: 1;
    kind: 'wasm-compile';
} | {
    protocol: 1;
    kind: 'call';
    id: number;
    method: string;
    args: unknown[];
} | {
    protocol: 1;
    kind: 'dispose';
};
export interface ReplyError {
    name: string;
    message: string;
    refusal?: unknown;
}
export type HostMessage = {
    protocol: 1;
    kind: 'wasm-module';
    url: string;
    module: WebAssembly.Module;
    hydration?: string;
    version?: string;
    sequence?: number;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: true;
    value: unknown;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: false;
    error: ReplyError;
} | {
    protocol: 1;
    kind: 'event';
    name: string;
    payload: unknown;
} | {
    protocol: 1;
    kind: 'progress';
    id: number;
} | {
    protocol: 1;
    kind: 'failure';
    code: SessionFailureCode;
    message: string;
    diagnostics?: string;
};
declare const TRANSFER: unique symbol;
declare const DEFERRED: unique symbol;
export interface DeferredReply<T> {
    readonly [DEFERRED]: true;
    readonly promise: Promise<T | TransferResult<T>>;
}
/** Releases the foreground lane while a reply is pending. */
export declare function deferReply<T>(promise: Promise<T | TransferResult<T>>): DeferredReply<T>;
export declare function isDeferredReply(value: unknown): value is DeferredReply<unknown>;
export interface TransferResult<T> {
    readonly [TRANSFER]: true;
    readonly value: T;
    readonly transfer: Transferable[];
}
/** Transfers ownership of a handler result when replying. */
export declare function transferable<T>(value: T, transfer: Transferable[]): TransferResult<T>;
export declare function isTransferResult(value: unknown): value is TransferResult<unknown>;
/** `Object.hasOwn` for ES2020 targets. */
export declare function hasOwn(value: object, key: PropertyKey): boolean;
export declare function isClientMessage(value: unknown): value is ClientMessage;
export declare function isHostMessage(value: unknown): value is HostMessage;
export {};

// ---- shared/office-session/scheduler.d.ts ----
export type SchedulerLane = 'input' | 'collab' | 'interactive' | 'background';
declare const FOREGROUND_LANES: readonly ["input", "collab", "interactive"];
type ForegroundLane = (typeof FOREGROUND_LANES)[number];
export declare const INPUT_SLICE_MS = 8;
export declare const IDLE_SLICE_MS = 24;
export declare const INPUT_ACTIVE_MS = 1000;
export interface SchedulerMessage {
    lane: ForegroundLane;
    run(): Promise<void> | void;
    userInput?: boolean;
    holdsIdleTasks?: boolean;
    reorderable?: boolean;
    mutates?: boolean;
    reframes?: boolean;
    key?: string;
    replaceableBy?: string;
    supersede?(): void;
}
export type TaskStep = 'done' | 'yield';
export interface SchedulerTask {
    kind: string;
    version: number;
    generation: number;
    run(budgetMs: number): Promise<TaskStep> | TaskStep;
    idleAfterInputMs?: number;
    onStale?(): 'cancel' | 'continue';
    cancel?(): void;
    fail?(error: unknown): void;
}
export interface PureTask<Input = unknown, Result = unknown> {
    kind: string;
    version: number;
    generation: number;
    input: Input;
    transfer?: Transferable[];
    compute(input: Input): Result;
    install(result: Result): Promise<void> | void;
    cancel?(reason?: unknown): void;
}
export interface TaskExecutor {
    run<Input, Result>(job: {
        kind: string;
        input: Input;
        transfer?: Transferable[];
    }): Promise<Result>;
}
export interface SchedulerHost {
    now(): number;
    turn(callback: () => void): void;
    timer(callback: () => void, ms: number): () => void;
    executor?: TaskExecutor;
    failed?(error: unknown): void;
}
export interface ResidentScheduler {
    submit(message: SchedulerMessage): void;
    schedule(task: SchedulerTask): void;
    dispatch<Input, Result>(task: PureTask<Input, Result>): void;
    budget(): number;
    readonly version: number;
    readonly generation: number;
    bump(changes: {
        version?: boolean;
        generation?: boolean;
    }): void;
    pending(): {
        foreground: number;
        background: number;
    };
}
export declare function createResidentScheduler(host: SchedulerHost): ResidentScheduler;
export {};

// ---- shared/office-session/transport.d.ts ----
/** Endpoint seam: ordered reliable delivery; ports and sockets may copy instead of transfer. */
export interface SessionTransport {
    post(message: unknown, transfer?: Transferable[]): void;
    listen(listener: (message: unknown) => void): () => void;
    onError(listener: (error: unknown) => void): () => void;
    close(): void;
}
export interface SessionScope {
    postMessage(message: unknown, transfer: Transferable[]): void;
    addEventListener(type: string, listener: EventListener): void;
    removeEventListener(type: string, listener: EventListener): void;
    close?(): void;
    start?(): void;
}
/** Adapts an existing worker without constructing one. */
export declare function createWorkerTransport(worker: Worker): SessionTransport;
/** Adapts a worker scope or message port. */
export declare function createScopeTransport(scope: SessionScope): SessionTransport;

// ---- shared/office-session/types.d.ts ----
import type { DeferredReply, TransferResult } from './protocol';
import type { ResidentScheduler } from './scheduler';
export declare const SESSION_SUPERSEDED = "SessionSuperseded";
export type SessionScheduler = Pick<ResidentScheduler, 'schedule' | 'dispatch' | 'bump' | 'budget' | 'pending' | 'version' | 'generation'>;
/** Format-owned, structured-cloneable methods. */
export type SessionMethods = Record<string, (...args: any[]) => unknown>;
export type Promisified<M extends SessionMethods> = {
    [K in keyof M]: (...args: Parameters<M[K]>) => Promise<Awaited<ReturnType<M[K]>>>;
};
export type MethodHandlers<M extends SessionMethods, C> = {
    [K in keyof M]: (context: C, ...args: Parameters<M[K]>) => Awaited<ReturnType<M[K]>> | TransferResult<Awaited<ReturnType<M[K]>>> | DeferredReply<Awaited<ReturnType<M[K]>>> | Promise<Awaited<ReturnType<M[K]>> | TransferResult<Awaited<ReturnType<M[K]>>> | DeferredReply<Awaited<ReturnType<M[K]>>>>;
};
export interface MethodPolicy {
    lane: 'input' | 'collab' | 'interactive';
    mutates?: boolean;
    reframes?: boolean;
    reorderable?: boolean;
    key?: string;
    replaceableBy?: string;
    userInput?: boolean;
    holdsIdleTasks?: boolean;
}
export type MethodPolicies<M extends SessionMethods> = {
    [K in keyof M]: MethodPolicy | ((...args: Parameters<M[K]>) => MethodPolicy);
};
export type SessionEvents = Record<string, unknown>;
export interface SessionState {
    format: string;
    stage: 'preview' | 'ready';
    version: number;
    dirty: boolean;
}
export type SessionFailureCode = 'crash' | 'trap' | 'out-of-memory' | 'message' | 'silence' | 'disposed';
/** A terminal session failure. */
export declare class SessionFailure extends Error {
    readonly code: SessionFailureCode;
    readonly diagnostics?: string | undefined;
    constructor(code: SessionFailureCode, message: string, diagnostics?: string | undefined);
}

// ---- structuredExport.d.ts ----
/**
 * Read-only structured export of PPTX content: slides in deck order, each with its shape tree,
 * text stories, tables and placeholders for content the export cannot represent, every record
 * with the location it was read from, and diagnostics for everything omitted. Plain JSON,
 * produced by one Rust walker for sessions and bytes.
 *
 * Reading order is the current shape tree, depth first: slides in deck order, shapes in tree
 * order, a group's descendants at the group's position, table cells row by row. It is the
 * authored order, not a reading order inferred from geometry.
 *
 * `range` anchors are batch text targets in the story offsets of `readContent()`, so a session
 * export's `range` anchor can be a step's `target` at the version it was read at; `notes` and
 * `comment` anchors index the plain text they carry; `sourcePart` anchors address retained source
 * XML by element-child ordinals. Session anchors belong to the version they were read at; ids are
 * deterministic export-tree paths.
 */
import type { OperationRefusal } from '../../../shared/host-contracts/edits';
import type { PptxTextTarget } from './edits';
import type { ExportCompletion, ExportDiagnostic, MarkdownAnchor } from '../../../shared/host-contracts/exports';
export type { ExportCompletion, ExportDiagnostic, ExportSeverity, MarkdownAnchor, } from '../../../shared/host-contracts/exports';
/** `session` anchors resolve against the returned version; `snapshot` ones against the content. */
export type PptxAnchorScope = 'session' | 'snapshot';
/** A half-open range of UTF-16 offsets. */
export interface PptxTextSpan {
    start: number;
    end: number;
}
export type PptxAnchor = {
    kind: 'slide';
    slideId: string;
} | {
    kind: 'shape';
    slideId: string;
    shapeId: string;
} | Extract<PptxTextTarget, {
    kind: 'range';
}> | {
    kind: 'notes';
    slideId: string;
    range: PptxTextSpan;
} | {
    kind: 'comment';
    slideId: string;
    commentId: string;
    range: PptxTextSpan;
} | {
    kind: 'sourcePart';
    part: string;
    partSha256: string;
    path: number[];
};
/**
 * The retained source XML a record was seeded from: `sldId` is the owning slide's `p:sldId/@id`,
 * `sourceId` a shape's `cNvPr/@id`.
 */
export interface PptxSourceProvenance {
    part: string;
    partSha256: string;
    path: number[];
    sldId: number | null;
    sourceId: number | null;
}
export interface PptxStructuredContent extends ExportCompletion {
    schemaVersion: 1;
    anchorScope: PptxAnchorScope;
    readingOrder: 'shapeTree';
    included: {
        hiddenSlides: boolean;
        hiddenShapes: boolean;
        notes: boolean;
        comments: boolean;
        formatting: boolean;
    };
    slides: PptxExportSlide[];
    diagnostics: PptxExportDiagnostic[];
}
export interface PptxExportSlide {
    id: string;
    /** Zero-based position in the current deck, hidden slides counted. */
    index: number;
    anchor: PptxAnchor;
    name: string | null;
    /**
     * `null` when the deck does not record it: such slides are exported with a
     * `visibility-unknown` diagnostic.
     */
    hidden: boolean | null;
    provenance: PptxSourceProvenance | null;
    shapes: PptxExportShape[];
    notes: PptxExportNotes | null;
    comments: PptxExportComment[];
}
/** `unknown` is a shape-tree element only the source holds, anchored by its source location. */
export type PptxExportShapeKind = 'shape' | 'picture' | 'graphicFrame' | 'group' | 'unknown';
export interface PptxExportShape {
    id: string;
    anchor: PptxAnchor;
    kind: PptxExportShapeKind;
    name: string;
    /** Alternative-text title and description. */
    title: string | null;
    description: string | null;
    /** Hidden itself or through a hidden group. */
    hidden: boolean;
    placeholder: {
        type: string | null;
        index: number | null;
    } | null;
    provenance: PptxSourceProvenance | null;
    stories: PptxExportStory[];
    table: PptxExportTable | null;
    object: PptxExportObject | null;
    children: PptxExportShape[];
}
export interface PptxExportStory {
    id: string;
    /** Spans the whole story. */
    anchor: PptxAnchor;
    paragraphs: PptxExportParagraph[];
}
export interface PptxExportParagraph {
    id: string;
    anchor: PptxAnchor;
    paragraphId: string;
    level: number;
    /** The authored alignment, else the inherited one. */
    alignment: string | null;
    /** The paragraph's own bullet, as the session stores it. */
    bulletJson: string | null;
    /** The marker after inheritance and numbering; `null` for a paragraph that is no list item. */
    list: PptxExportList | null;
    runs: PptxExportRun[];
}
/**
 * `character` is the authored `a:buChar` and `font` its typeface, a symbol font when the character
 * names a glyph slot. `marker` is the formatted number, `null` for a scheme that cannot be
 * formatted.
 */
export type PptxExportList = {
    kind: 'bullet';
    character: string;
    font: string | null;
} | {
    kind: 'number';
    scheme: string;
    startAt: number;
    value: number;
    marker: string | null;
};
export type PptxExportMark = {
    kind: 'bold' | 'italic' | 'superscript' | 'subscript' | 'smallCaps' | 'allCaps';
} | {
    kind: 'underline';
    style: string;
};
/** Line breaks are one `\n` unit, fields cover their cached result; unsupported runs are zero-width. */
export type PptxExportRun = {
    anchor: PptxAnchor;
    /** `null` when the export excludes formatting. */
    marks: PptxExportMark[] | null;
    /** A click hyperlink: an external URL, or the part an internal link jumps to. */
    link: {
        href: string;
        external: boolean;
    } | null;
} & ({
    kind: 'text';
    text: string;
} | {
    kind: 'lineBreak';
}
/** A field's cached result; fields are never evaluated. */
 | {
    kind: 'field';
    fieldType: string | null;
    text: string;
} | {
    kind: 'unsupported';
    element: string;
});
/**
 * A table on its grid. Every source cell appears once; `merged` marks a continuation covered by
 * the cell at `mergeOrigin`, whose content PowerPoint does not show.
 */
export interface PptxExportTable {
    columns: number;
    rows: Array<{
        cells: PptxExportTableCell[];
    }>;
}
export interface PptxExportTableCell {
    id: string;
    anchor: PptxAnchor;
    row: number;
    column: number;
    gridSpan: number;
    rowSpan: number;
    merged: boolean;
    mergeOrigin: {
        row: number;
        column: number;
    } | null;
    story: PptxExportStory | null;
}
export type PptxExportObjectKind = 'picture' | 'video' | 'audio' | 'chart' | 'smartArt' | 'embeddedObject' | 'unknown';
/** A placeholder for content the export does not represent; its data is never included. */
export interface PptxExportObject {
    id: string;
    kind: PptxExportObjectKind;
    /** The source element's qualified name, such as `p:graphicFrame`. */
    element: string;
    uri: string | null;
    relationshipIds: string[];
    /** The package parts the object references. */
    parts: string[];
    /** Targets of its external relationships, such as a linked video. */
    externalTargets: string[];
}
/** A slide's speaker notes, as plain text. */
export interface PptxExportNotes {
    id: string;
    anchor: PptxAnchor;
    text: string;
    provenance: PptxSourceProvenance | null;
}
export interface PptxExportComment {
    id: string;
    anchor: PptxAnchor;
    commentId: string;
    author: string | null;
    date: string | null;
    parentId: string | null;
    resolved: boolean;
    text: string;
}
export type PptxExportDiagnosticCode = 'unsupported-content' | 'image-data-omitted' | 'unsupported-numbering' | 'field-cached-result' | 'provenance-unavailable' | 'visibility-unknown' | 'hidden-content-excluded' | 'stories-omitted' | 'formatting-omitted' | 'inherited-content-omitted' | 'notes-structure-omitted' | 'merge-continuation-content-omitted' | 'truncated' | 'markdown-lossy';
export type PptxExportDiagnostic = ExportDiagnostic<PptxExportDiagnosticCode, PptxAnchor>;
/**
 * What to export. Hidden slides and shapes, notes and comments are excluded unless requested, and
 * formatting is included unless turned off. `maxBlocks` defaults to 10,000 (at most 1,000,000;
 * every slide, shape, paragraph, notes story and comment counts) and `maxBytes` to 8,388,608
 * (1,024 to 67,108,864, measured on the compact JSON).
 */
export interface PptxExportOptions {
    includeHiddenSlides?: boolean;
    includeHiddenShapes?: boolean;
    includeNotes?: boolean;
    includeComments?: boolean;
    includeFormatting?: boolean;
    maxBlocks?: number;
    maxBytes?: number;
}
/** `maxBytes` bounds the Markdown text; the default and range are the export's. */
export interface PptxMarkdownOptions {
    maxBytes?: number;
}
/**
 * Markdown rendered from structured content. Each `<!-- pptx-export:N -->` marker precedes the
 * record `anchors` maps it to. Markdown keeps no slide layout, geometry or typography.
 */
export interface PptxMarkdownContent extends ExportCompletion {
    markdown: string;
    anchors: MarkdownAnchor<PptxAnchor>[];
    diagnostics: PptxExportDiagnostic[];
}
/**
 * Options out of range or above a hard maximum, or structured content handed to the renderer that
 * contradicts itself. Unsupported deck content is diagnosed, never refused.
 */
export type PptxExportFailureCode = 'invalid-options' | 'limit-exceeded' | 'invalid-content';
/** Why an export was refused; `target` is the anchor the refusal concerns, if any. */
export interface PptxExportFailure {
    code: PptxExportFailureCode;
    target: PptxAnchor | null;
    message: string;
}
export type PptxExportRefusal = OperationRefusal<PptxExportFailure>;
/** A session export: the content with the version it was read at, or a refusal. */
export type PptxExportResult<T> = {
    ok: true;
    version: string;
    content: T;
} | PptxExportRefusal;
/** A headless export or rendering the engine refused as data. */
export declare class PptxExportError extends Error {
    readonly failure: PptxExportFailure;
    constructor(failure: PptxExportFailure);
}

// ---- types.d.ts ----
export interface TextStyle {
    bold?: boolean;
    italic?: boolean;
    fontSizePt?: number;
    color?: string;
    fontFamily?: string;
    underline?: string;
    spacingPt?: number;
    baselinePct?: number;
}
export type TextStylePatch = TextStyle;
export interface TextStyleSnapshot {
    bold: boolean | null;
    italic: boolean | null;
    fontSizePt: number | null;
    color: string | null;
    fontFamily: string | null;
    underline: string | null;
    spacingPt?: number | null;
    baselinePct?: number | null;
    /** `a:rPr@kern`: the smallest size, in points, the run is kerned at. */
    kernPt?: number | null;
    /** `a:rPr@cap`: how the run is cased when drawn, never in the stored text. */
    caps?: 'none' | 'small' | 'all' | null;
}
export interface TextRunSnapshot {
    text: string;
    style: TextStyleSnapshot;
}
/** OOXML `a:pPr@algn` token. */
export type ParagraphAlignment = 'l' | 'ctr' | 'r' | 'just';
export interface ParagraphSnapshot {
    id: string;
    alignment: string | null;
    level: number;
    bulletJson: string | null;
    runs: TextRunSnapshot[];
}
export interface StorySnapshot {
    id: string;
    length: number;
    paragraphs: ParagraphSnapshot[];
}
export interface PptxTextSearchOptions {
    /** Defaults to false. */
    caseSensitive?: boolean;
    /** Maximum matches; unlimited by default. */
    limit?: number;
}
/**
 * A caret position in a story that later edits, undo, redo and remote updates
 * move along with the text. Plain data; `position` is opaque.
 */
export interface PptxCaretAnchor {
    storyId: string;
    position: string;
}
/** Zero-based slide index; story-local UTF-16 offsets. */
export interface PptxTextMatch {
    slideIndex: number;
    slideId: string;
    shapeId: string;
    storyId: string;
    start: number;
    end: number;
    text: string;
}
export type ShapeKind = 'shape' | 'picture' | 'graphicFrame' | 'group';
/** Hue and hue offsets use turns; other values use fractions. */
export type ColorTransform = {
    kind: 'tint' | 'shade' | 'hue' | 'hueMod' | 'hueOff' | 'sat' | 'satMod' | 'satOff' | 'lum' | 'lumMod' | 'lumOff' | 'red' | 'redMod' | 'redOff' | 'green' | 'greenMod' | 'greenOff' | 'blue' | 'blueMod' | 'blueOff' | 'alpha' | 'alphaMod' | 'alphaOff';
    value: number;
} | {
    kind: 'comp' | 'inv' | 'gray' | 'gamma' | 'invGamma';
};
export interface ColorValue {
    rgb?: string;
    themeColor?: string;
    themeTint?: string;
    themeShade?: string;
    auto?: boolean;
    luminanceModulation?: number;
    luminanceOffset?: number;
    saturationModulation?: number;
    alpha?: number;
    /** Overrides legacy colour modifiers; legacy alpha applies without an alpha step. */
    transforms?: ColorTransform[];
}
export interface ShapeGradient {
    type: string;
    angle?: number;
    stops: Array<{
        position: number;
        color: ColorValue;
    }>;
}
/** An authored fill as the deck stores it. */
export interface ShapeFill {
    type: string;
    color?: ColorValue;
    gradient?: ShapeGradient;
}
export interface ShapeLineEnd {
    type: string;
    width: string | null;
    length: string | null;
}
/** An authored outline as the deck stores it; `width` is EMU. */
export interface ShapeOutline {
    width?: number;
    color?: ColorValue;
    gradient?: ShapeGradient;
    style?: string;
    cap?: string;
    join?: string;
    headEnd?: ShapeLineEnd;
    tailEnd?: ShapeLineEnd;
}
export type BlipEffect = {
    type: 'biLevel';
    threshold: number;
} | {
    type: 'grayscale';
} | {
    type: 'luminance';
    brightness: number;
    contrast: number;
} | {
    type: 'duotone';
    shadow: ColorValue | null;
    highlight: ColorValue | null;
} | {
    type: 'colorChange';
    from: ColorValue | null;
    to: ColorValue | null;
    useAlpha?: boolean;
};
export interface ShapeSnapshot {
    id: string;
    sourceId: number;
    kind: ShapeKind;
    name: string;
    /** `p:cNvPr/@title`, the alternative-text title. */
    title?: string;
    /** `p:cNvPr/@descr`, the alternative-text description. */
    description?: string;
    /** True when `p:cNvPr` carries the a16 decorative extension; omitted when false. */
    decorative?: boolean;
    x: number;
    y: number;
    width: number;
    height: number;
    rotationDeg: number;
    flipH: boolean;
    flipV: boolean;
    /** The geometry the shape draws at, present only while it has none of its own. */
    inherited?: InheritedGeometry | null;
    /** Hides this shape and its descendants; omitted when false. */
    hidden?: boolean;
    textBox?: boolean;
    aspectRatioLocked?: boolean;
    geometry: string;
    adjustValues: Record<string, number>;
    placeholder: unknown | null;
    fill: ShapeFill | null;
    resolvedFillColor: string | null;
    outline: ShapeOutline | null;
    resolvedOutlineColor: string | null;
    mediaPartPath: string | null;
    /** Image data added to this session, retained across saves. */
    pendingMedia?: {
        contentType: string;
        base64: string;
    } | null;
    blipEffects?: BlipEffect[];
    graphic: unknown | null;
    textStories: StorySnapshot[];
    children: ShapeSnapshot[];
}
/** The transform a placeholder takes from its layout or master. */
export interface InheritedGeometry {
    x: number;
    y: number;
    width: number;
    height: number;
    rotationDeg: number;
    flipH: boolean;
    flipV: boolean;
}
/** Stored layout or master settings; omitted boolean attributes default to true and do not gate slide placeholders. */
export interface HeaderFooterSettings {
    slideNumber: boolean;
    footer: boolean;
    date: boolean;
    header: boolean;
}
/** Session shape id and text of the first visible slide placeholder in shape order, including groups; text joins stories with newlines and preserves empty text.
 * Slide number fields use firstSlideNum (default 1) plus deck index, counting hidden slides; fixed text stays unchanged. */
export interface SlideHeaderFooterText {
    shapeId: string;
    text: string;
}
/** Stored date text; field is the first datetime* field type (updates on display), or null for fixed text. */
export interface SlideDateText extends SlideHeaderFooterText {
    field: string | null;
}
/** Visible placeholders from the slide itself, independent of stored settings; missing placeholders are null. */
export interface SlideHeaderFooter {
    footer: SlideHeaderFooterText | null;
    slideNumber: SlideHeaderFooterText | null;
    date: SlideDateText | null;
}
export interface SlideLayoutSnapshot {
    /** Layout settings, falling back to the master; null when neither stores p:hf. */
    headerFooter?: HeaderFooterSettings | null;
    name: string | null;
    type: string | null;
}
export interface SlideSnapshot {
    headerFooter?: SlideHeaderFooter | null;
    title?: string | null;
    hidden?: boolean | null;
    layout?: SlideLayoutSnapshot | null;
    id: string;
    sourcePartPath: string | null;
    layoutPartPath: string | null;
    name: string | null;
    /** Speaker notes as plain text; absent when the slide has none. */
    notes?: string;
    /** The slide's own background; absent when it inherits the layout's. */
    background?: ShapeFill;
    /** `true` once an edit changed the background from the one the file declares. */
    backgroundEdited?: boolean;
    shapes: ShapeSnapshot[];
}
export interface DeckSnapshot {
    widthEmu: number;
    heightEmu: number;
    slides: SlideSnapshot[];
    commentFlavor?: CommentFlavor;
    comments?: CommentSnapshot[];
}
/** Legacy comments or modern threads. */
export type CommentFlavor = 'legacy' | 'modern';
export interface CommentSnapshot {
    id: string;
    slideId: string;
    author: string;
    initials: string;
    text: string;
    created: string | null;
    xEmu: number;
    yEmu: number;
    /** Set on a reply; names the thread root. */
    parentId: string | null;
    resolved: boolean;
}
export interface CommentReceipt {
    commentId: string;
    slideId: string;
    parentId: string | null;
    resolved: boolean;
}
export interface SlideReceipt {
    slideId: string;
    fromIndex: number | null;
    toIndex: number | null;
}
export interface ShapeReceipt {
    slideId: string;
    shapeId: string;
    index: number;
}
export interface ShapeZOrderReceipt {
    slideId: string;
    shapeId: string;
    fromIndex: number;
    toIndex: number;
}
export interface ShapeRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
export interface TransformReceipt {
    slideId: string;
    shapeId: string;
    before: ShapeRect;
    after: ShapeRect;
}
export interface TextReceipt {
    storyId: string;
    start: number;
    end: number;
    text: string;
}
export interface ShapeDraft {
    name: string;
    rect: ShapeRect;
    text: string;
    style: TextStyle;
}
export interface PresetShapeDraft {
    name: string;
    geometry: string;
    rect: ShapeRect;
    fill?: string | null;
}
export interface PictureDraft {
    name: string;
    rect: ShapeRect;
    /** The image's MIME type, e.g. `image/png`. */
    contentType: string;
    /** The image bytes, base64-encoded. */
    mediaBase64: string;
}
export interface ShapeStroke {
    color?: string;
    widthPt?: number;
}
export interface ShapeFillReceipt {
    slideId: string;
    shapeId: string;
    before: string | null;
    after: string | null;
}
export interface ShapeStrokeReceipt {
    slideId: string;
    shapeId: string;
    before: ShapeStroke | null;
    after: ShapeStroke | null;
}
export interface ShapeAdjustReceipt {
    slideId: string;
    shapeId: string;
    before: Record<string, number>;
    after: Record<string, number>;
}
export interface HistoryResult {
    applied: boolean;
    snapshot: DeckSnapshot;
}
/** Renderer stage latencies of one profiled slide layout, in ms. */
export interface LayoutProfile {
    scopeMs: number;
    layoutMs: number;
    serializeMs: number;
}
export interface ProfiledLayout {
    layout: SlideDisplayList;
    profile: LayoutProfile;
}
/** Boundary stage latencies of one profiled edit, in ms. */
export interface EditProfile {
    parseMs: number;
    applyMs: number;
    serializeMs: number;
}
export interface HistoryProfile {
    undoMs: number;
    snapshotMs: number;
    serializeMs: number;
}
export interface Profiled<T, P = EditProfile> {
    receipt: T;
    profile: P;
}
export interface PptxFontFace {
    family: string;
    bold?: boolean;
    italic?: boolean;
    bytes: Uint8Array;
}
export type GeometryPathCommand = {
    type: 'move';
    x: number;
    y: number;
} | {
    type: 'line';
    x: number;
    y: number;
} | {
    type: 'quad';
    cpx: number;
    cpy: number;
    x: number;
    y: number;
} | {
    type: 'cubic';
    cp1x: number;
    cp1y: number;
    cp2x: number;
    cp2y: number;
    x: number;
    y: number;
} | {
    type: 'close';
};
export type Paint = {
    kind: 'solid';
    color: string;
} | {
    kind: 'gradient';
    gradientType: 'linear' | 'radial' | 'rectangular' | 'path';
    angleDeg?: number;
    stops: Array<{
        position: number;
        color: string;
    }>;
};
export interface StrokeEnd {
    kind: string;
    width: number;
    length: number;
}
export interface Stroke {
    /** Solid colour or first gradient stop. */
    color: string;
    width: number;
    dashed?: boolean;
    paint?: Paint;
    join?: 'round' | 'bevel' | 'miter';
    headEnd?: StrokeEnd;
    tailEnd?: StrokeEnd;
}
/** An `a:outerShdw`: a blurred copy of the shape's own path, offset and tinted. */
export interface Shadow {
    paths?: Array<{
        path: GeometryPathCommand[];
        fill: boolean;
        stroke?: Stroke;
    }>;
    color: string;
    blur?: number;
    dx?: number;
    dy?: number;
    scaleX?: number;
    scaleY?: number;
}
export interface PrimitiveTransform {
    rotationDeg?: number;
    flipH?: boolean;
    flipV?: boolean;
}
interface PrimitiveBase {
    objectId: number;
    shapeId?: string;
    x: number;
    y: number;
    w: number;
    h: number;
    transform?: PrimitiveTransform;
}
export interface ShapePrimitive extends PrimitiveBase {
    kind: 'shape';
    name: string;
    geometry: string;
    path: GeometryPathCommand[];
    /** A rectangle substitutes for the authored outline or clip. */
    geometryFallback?: boolean;
    clip?: GeometryPathCommand[];
    evenOdd?: boolean;
    adjustValues?: Record<string, number>;
    fill?: Paint;
    stroke?: Stroke;
    shadow?: Shadow;
}
/** An `a:blip` colour transform, colours already resolved to `#rrggbbaa`. */
export type ImageEffect = {
    kind: 'biLevel';
    threshold: number;
} | {
    kind: 'grayscale';
} | {
    kind: 'alpha';
    amount: number;
} | {
    kind: 'luminance';
    brightness: number;
    contrast: number;
} | {
    kind: 'duotone';
    shadow: string;
    highlight: string;
} | {
    kind: 'colorChange';
    from: string;
    to: string;
    useAlpha?: boolean;
};
export interface ImageCrop {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
}
export interface ImagePrimitive extends PrimitiveBase {
    kind: 'image';
    name: string;
    assetId?: string;
    /** Applied to the bitmap in order before it is drawn. */
    effects?: ImageEffect[];
    /** Fraction of the source discarded per edge, from `a:srcRect`. */
    crop?: ImageCrop;
    /** Bitmap destination within the picture frame. */
    destination?: {
        x: number;
        y: number;
        w: number;
        h: number;
    };
    /** `a:tile`: repeat the picture at its own size, scaled by these fractions. */
    tile?: {
        scaleX: number;
        scaleY: number;
    };
    /** Outline the picture is masked to, when its `spPr` gives it one. */
    path?: GeometryPathCommand[];
    /** A rectangle substitutes for the authored mask. */
    geometryFallback?: boolean;
    stroke?: Stroke;
    shadow?: Shadow;
}
export interface CaretStop {
    position: number;
    x: number;
}
export interface PositionedGlyph {
    glyphId: number;
    cluster: number;
    x: number;
    advance: number;
    xOffset: number;
    yOffset: number;
}
export interface PositionedTextRun {
    text: string;
    start: number;
    end: number;
    x: number;
    width: number;
    fontId: number;
    fontFamily: string;
    fontSizePx: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
    letterSpacingPx?: number;
    baselineOffsetPx?: number;
    glyphs: PositionedGlyph[];
}
export interface PositionedTextLine {
    columnOffset?: number;
    x: number;
    y: number;
    width: number;
    height: number;
    baseline: number;
    start: number;
    end: number;
    runs: PositionedTextRun[];
    caretStops: CaretStop[];
}
export interface TextBoxPrimitive extends PrimitiveBase {
    kind: 'textBox';
    storyId?: string;
    anchor: 'top' | 'center' | 'bottom';
    paragraphs: Array<{
        align?: 'left' | 'center' | 'right' | 'justify';
        level: number;
        runs: Array<{
            text: string;
            fontFamily: string;
            fontSizePt: number;
            bold?: boolean;
            italic?: boolean;
            underline?: boolean;
            color: string;
        }>;
    }>;
    lines: PositionedTextLine[];
    overflow?: boolean;
    /** `a:rPr/a:effectLst`: the shadow the box's glyphs are drawn with. */
    textShadow?: Shadow;
}
export interface PlaceholderPrimitive extends PrimitiveBase {
    kind: 'placeholder';
    name: string;
    label?: string;
}
/** A plotted chart. Its parts paint clipped to the chart rectangle, and
 *  `label` is the screen-reader summary of the whole chart. */
export interface ChartPrimitive extends PrimitiveBase {
    kind: 'chart';
    name: string;
    label: string;
    primitives: SlidePrimitive[];
}
/** A laid-out table. Its cells paint clipped to the table rectangle, and
 *  `label` is the screen-reader summary of the whole table. */
export interface TablePrimitive extends PrimitiveBase {
    kind: 'table';
    name: string;
    label: string;
    primitives: SlidePrimitive[];
}
export type SlidePrimitive = ShapePrimitive | ImagePrimitive | TextBoxPrimitive | PlaceholderPrimitive | ChartPrimitive | TablePrimitive;
export interface SlideDisplayList {
    contractVersion: number;
    width: number;
    height: number;
    background?: Paint;
    primitives: SlidePrimitive[];
}
export type HitTestResult = {
    kind: 'shape';
    shapeId: string;
} | {
    kind: 'text';
    shapeId: string;
    storyId: string;
    position: number;
};
export interface UpdateEvent {
    origin: 'local' | 'remote';
    update: Uint8Array;
}
export interface TextParagraphDraft {
    bullet?: boolean;
    alignment?: ParagraphAlignment | null;
    runs: Array<{
        text: string;
        style?: TextStyle;
    }>;
}
export interface TableReceipt {
    slideId: string;
    shapeId: string;
    rowCount: number;
    columnCount: number;
    newCellRefs: string[];
}
export {};

// ---- wasm/generated/pptx_wasm.d.ts ----
/* tslint:disable */
/* eslint-disable */
export interface DecodedFontFace { font: Uint8Array; payload?: Uint8Array; }


export class PptxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    acceptProposalJson(args: string): string;
    addCommentJson(args: string): string;
    addPictureJson(args: string): string;
    addShapeJson(args: string): string;
    addTextBoxJson(args: string): string;
    addTextBoxProfiledJson(args: string): string;
    addUndoBoundary(): void;
    adoptPeerIdentity(): void;
    /**
     * Anchors a caret offset so that later edits, undo and remote updates move it.
     */
    anchorCaretJson(args: string): string;
    /**
     * Applies a batch all-or-nothing: `{"ok":true,"baseVersion","version","applied","source",
     * "changedSlides","changedStories","receipts"}`.
     */
    applyEditsJson(request: string): string;
    applyUpdateJson(update: Uint8Array): string;
    bringShapeForwardJson(args: string): string;
    bringShapeToFrontJson(args: string): string;
    canRedo(): boolean;
    canUndo(): boolean;
    checkTableEditJson(args: string): void;
    clearUpdateObservation(): void;
    commentsJson(): string;
    deleteSlideJson(args: string): string;
    deleteTableColumnJson(args: string): string;
    deleteTableRowJson(args: string): string;
    deleteTextJson(args: string): string;
    deleteTextProfiledJson(args: string): string;
    /**
     * The session-scoped version token of the committed deck state.
     */
    documentVersion(): string;
    drainUpdateEvent(): Uint8Array;
    duplicateSlideJson(args: string): string;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    /**
     * `exportStructuredJson` rendered as Markdown from the same read:
     * `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
     */
    exportMarkdownJson(options: string): string;
    /**
     * Structured export of the committed deck: `{"includeHiddenSlides"?,"includeHiddenShapes"?,
     * "includeNotes"?,"includeComments"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
     * `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Nothing changes.
     */
    exportStructuredJson(options: string): string;
    /**
     * `{"text","within"?,"limit"?}` -> `{"ok":true,"version","matches","truncated"}`.
     */
    findTextJson(request: string): string;
    formatTextJson(args: string): string;
    insertParagraphBreakJson(args: string): string;
    insertSlideJson(args: string): string;
    insertSlideProfiledJson(args: string): string;
    insertTableColumnJson(args: string): string;
    insertTableRowJson(args: string): string;
    insertTextJson(args: string): string;
    insertTextProfiledJson(args: string): string;
    listProposalsJson(): string;
    mediaBytes(part_path: string): Uint8Array;
    moveShapeJson(args: string): string;
    moveShapeProfiledJson(args: string): string;
    moveSlideJson(args: string): string;
    /**
     * `source` is the file the update was seeded from; when it matches the
     * recorded fingerprint the session keeps its part bytes and can save, and
     * failing to reattach it throws.
     * Any other bytes fall back to the bare update session, whose `saveBytes`
     * fails — joining a room must not depend on carrying the right file.
     */
    static openCollaborativeFromUpdate(update: Uint8Array, client_id: number, source?: Uint8Array | null): PptxDocument;
    /**
     * Opens a session that can exchange updates. It is not marked collaborative until
     * `setCollaborative` is called.
     */
    static openCollaborative(bytes: Uint8Array, client_id: number): PptxDocument;
    static openPeerDeckJson(source: Uint8Array, identity_json: string, initial_update?: Uint8Array | null): PptxDocument;
    static openReplayBaseline(source: Uint8Array, client_id?: number | null, initial_update?: Uint8Array | null): PptxDocument;
    peerHydrationJson(fonts_json: string): string;
    previewProposalJson(args: string): string;
    proposeJson(args: string): string;
    readCommandContentJson(request: string): string;
    /**
     * `{"slideIds"?}` -> `{"ok":true,"version","slides","stories"}`.
     */
    readContentJson(request: string): string;
    redoJson(): string;
    redoReceiptJson(): string;
    registerPeerFontsJson(fonts_json: string): void;
    rejectProposalJson(args: string): string;
    removeCommentJson(args: string): string;
    removeShapeJson(args: string): string;
    replayJson(envelope: string): string;
    replyToCommentJson(args: string): string;
    resizeShapeJson(args: string): string;
    /**
     * The current offset of an anchor from `anchorCaretJson`, or `null` once its story is gone.
     */
    resolveCaretAnchorJson(args: string): string;
    /**
     * Serializes the deck back to `.pptx` bytes, edits included.
     */
    saveBytes(): Uint8Array;
    searchTextJson(args: string): string;
    sendShapeBackwardJson(args: string): string;
    sendShapeToBackJson(args: string): string;
    sessionMetadataJson(): string;
    setCollaborative(): void;
    setCommentFlavorJson(args: string): string;
    setCommentPositionJson(args: string): string;
    setCommentStatusJson(args: string): string;
    setParagraphAlignmentJson(args: string): string;
    setShapeAdjustJson(args: string): string;
    setShapeFillJson(args: string): string;
    setShapeRectJson(args: string): string;
    setShapeStrokeJson(args: string): string;
    setSlideLayoutJson(args: string): string;
    setSlideNotesJson(args: string): string;
    setStoryParagraphsJson(args: string): string;
    setUndoCaptureMode(mode: string): void;
    snapshotJson(): string;
    startUpdateObservation(): void;
    storyFormattingJson(args: string): string;
    storyJson(args: string): string;
    undoCaptureMode(): string;
    undoJson(): string;
    /**
     * `undoJson` timed at its undo, snapshot and serialize boundaries, as
     * `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
     */
    undoProfiledJson(): string;
    undoReceiptJson(): string;
    /**
     * Runs every check of `applyEditsJson`, staging included, without changing anything:
     * `{"ok":true,"baseVersion","wouldApply","previews"}`.
     */
    validateEditsJson(request: string): string;
    static version(): string;
    readonly clientId: number;
}

export class PptxRenderer {
    free(): void;
    [Symbol.dispose](): void;
    hitTestJson(x: number, y: number): string;
    hitTestSlideJson(document: PptxDocument, id: string, x: number, y: number): string;
    layoutProposalDiffSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    layoutProposalSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    layoutSlideJson(document: PptxDocument, slide_index: number): string;
    /**
     * `layoutSlideJson` with its stages timed, returned as
     * `{"layout": ..., "profile": {"scopeMs", "layoutMs", "serializeMs"}}`.
     */
    layoutSlideProfiledJson(document: PptxDocument, slide_index: number): string;
    constructor();
    registerFallbackFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    /**
     * requested font faces before layout, for one slide or the whole deck.
     */
    requestedFontFacesJson(document: PptxDocument, slide_id?: string | null): string;
    saveBytes(document: PptxDocument): Uint8Array;
    setActiveSlide(document: PptxDocument, id: string, key: string): boolean;
    /**
     * Enables browser metrics for uncovered characters. Off by default.
     */
    setBrowserCoverageFallback(enabled: boolean): void;
    slideLayoutKey(document: PptxDocument, slide_index: number): string;
    snapshotWithLayoutKeysJson(document: PptxDocument): string;
}

export function compileSlideJson(slide_json: string): string;

/**
 * Decodes one font pack face (codec "zstd" or "woff2t+zstd") with this engine's linked decoder.
 * Returns owned bytes and an optional transformed payload. Non-zstd input passes through.
 * The default output limit is 64 MiB; woff2t decoding is always capped at 64 MiB.
 */
export function decodeFontFace(bytes: Uint8Array, codec: string, dict?: Uint8Array | null, maxOutput?: number | null): DecodedFontFace;

export function decodeTiffPng(data: Uint8Array): Uint8Array;

/**
 * `exportPptxStructuredJson` rendered as Markdown.
 */
export function exportPptxMarkdownJson(data: Uint8Array, options: string): string;

/**
 * Structured export of PPTX bytes as a snapshot, with the options of
 * `PptxDocument.exportStructuredJson`: `{"ok":true,"content"}` or `{"ok":false,"failure"}`.
 * Bytes that are not a readable PPTX throw.
 */
export function exportPptxStructuredJson(data: Uint8Array, options: string): string;

export function parsePptxJson(data: Uint8Array): string;

export function parsePptxJsonWithoutMedia(data: Uint8Array): string;

/**
 * Renders schema-version-1 structured content as Markdown; `options` is `{"maxBytes"?}`.
 */
export function renderPptxMarkdownJson(content: string, options: string): string;

export function rendererVersion(): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly __externref_table_alloc: () => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbg_pptxdocument_free: (a: number, b: number) => void;
    readonly __wbg_pptxrenderer_free: (a: number, b: number) => void;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_start: () => void;
    readonly compileSlideJson: (a: number, b: number) => [number, number, number, number];
    readonly decodeFontFace: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number];
    readonly decodeTiffPng: (a: number, b: number) => [number, number, number, number];
    readonly exportPptxMarkdownJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly exportPptxStructuredJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly memory: WebAssembly.Memory;
    readonly parsePptxJson: (a: number, b: number) => [number, number, number, number];
    readonly parsePptxJsonWithoutMedia: (a: number, b: number) => [number, number, number, number];
    readonly pptxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addPictureJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addUndoBoundary: (a: number) => void;
    readonly pptxdocument_adoptPeerIdentity: (a: number) => [number, number];
    readonly pptxdocument_anchorCaretJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_applyEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeForwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeToFrontJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_canRedo: (a: number) => number;
    readonly pptxdocument_canUndo: (a: number) => number;
    readonly pptxdocument_checkTableEditJson: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_clearUpdateObservation: (a: number) => void;
    readonly pptxdocument_clientId: (a: number) => number;
    readonly pptxdocument_commentsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_deleteSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTableColumnJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTableRowJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_documentVersion: (a: number) => [number, number];
    readonly pptxdocument_drainUpdateEvent: (a: number) => [number, number];
    readonly pptxdocument_duplicateSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly pptxdocument_encodeStateVector: (a: number) => [number, number];
    readonly pptxdocument_exportMarkdownJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_exportStructuredJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_findTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_formatTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertParagraphBreakJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTableColumnJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTableRowJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly pptxdocument_openCollaborativeFromUpdate: (a: number, b: number, c: number, d: number, e: number) => [number, number, number];
    readonly pptxdocument_openPeerDeckJson: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number];
    readonly pptxdocument_openReplayBaseline: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number];
    readonly pptxdocument_peerHydrationJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_previewProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_readCommandContentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_readContentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_redoReceiptJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_registerPeerFontsJson: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_replayJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_replyToCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resizeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resolveCaretAnchorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly pptxdocument_searchTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeBackwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeToBackJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sessionMetadataJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_setCollaborative: (a: number) => [number, number];
    readonly pptxdocument_setCommentFlavorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentPositionJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentStatusJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setParagraphAlignmentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeAdjustJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeFillJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeRectJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeStrokeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setSlideLayoutJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setSlideNotesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setStoryParagraphsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setUndoCaptureMode: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly pptxdocument_storyFormattingJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_storyJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_undoCaptureMode: (a: number) => [number, number];
    readonly pptxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_undoProfiledJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_undoReceiptJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_validateEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_version: () => [number, number];
    readonly pptxrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_hitTestSlideJson: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalDiffSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_new: () => number;
    readonly pptxrenderer_registerFallbackFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly pptxrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly pptxrenderer_requestedFontFacesJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly pptxrenderer_saveBytes: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_setActiveSlide: (a: number, b: number, c: number, d: number, e: number, f: number) => number;
    readonly pptxrenderer_setBrowserCoverageFallback: (a: number, b: number) => void;
    readonly pptxrenderer_slideLayoutKey: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_snapshotWithLayoutKeysJson: (a: number, b: number) => [number, number, number, number];
    readonly renderPptxMarkdownJson: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly rendererVersion: () => [number, number];
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;

// ---- wasm/loader.d.ts ----
import type { InitInput } from './generated/pptx_wasm.js';
import { type FontDecoder } from '../../../../shared/office-session/fontDecoder';
export declare const decodeFontFace: FontDecoder;
import type { PptxExportOptions, PptxExportResult, PptxMarkdownContent, PptxMarkdownOptions, PptxStructuredContent } from '../structuredExport';
import type { Proposal, ProposalAcceptance, ProposalDiffSlide, ProposalEdit, ProposalPreview } from '../proposals';
import type { PptxEditRequest, PptxEditResult, PptxFindRequest, PptxFindResult, PptxReadRequest, PptxReadResult, PptxTextField, PptxValidationResult } from '../edits';
import type { CollaborationReplica, CollaborationUpdateOrigin } from '../collaboration/types';
import type { CommentFlavor, CommentReceipt, CommentSnapshot, DeckSnapshot, HistoryProfile, HistoryResult, HitTestResult, ParagraphAlignment, PictureDraft, PresetShapeDraft, Profiled, ProfiledLayout, PptxFontFace, PptxCaretAnchor, PptxTextMatch, PptxTextSearchOptions, ShapeAdjustReceipt, ShapeDraft, ShapeFillReceipt, ShapeReceipt, ShapeRect, ShapeStroke, ShapeStrokeReceipt, ShapeZOrderReceipt, SlideDisplayList, SlideReceipt, StorySnapshot, TextReceipt, TextParagraphDraft, TextStyle, TextStylePatch, TransformReceipt, TableReceipt } from '../types';
export type WasmInitInput = InitInput | Promise<InitInput>;
export interface InspectPresentationOptions {
    includeMedia?: boolean;
}
/**
 * Faces a save lays out text with, in a renderer of its own, so the saved
 * bytes do not depend on the faces earlier renders registered with the handle.
 */
export interface SavePresentationOptions {
    fonts: ReadonlyArray<PptxFontFace>;
    /** Faces drawn only for characters the run's own face has no glyph for. */
    fallbackFonts?: ReadonlyArray<PptxFontFace>;
}
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
    /** @internal Retains the source bytes for `fork()` and `layouts()`; worker sessions pass false. */
    retainSource?: boolean;
}
export type UndoCaptureMode = 'auto' | 'manual';
export interface PresentationHistoryReceipt {
    applied: boolean;
    changedSlideIds: string[];
    deckSizeChanged: boolean;
    slideOrderChanged: boolean;
}
export type PresentationCommandReadResult = Extract<PptxReadResult, {
    ok: false;
}> | (Extract<PptxReadResult, {
    ok: true;
}> & {
    sources: {
        storyId: string;
        paragraphs: {
            paragraphId: string;
            text: string;
            fields: PptxTextField[];
            editable: boolean;
        }[];
    }[];
});
export interface PresentationHandle extends CollaborationReplica {
    isProposalsAvailable(): boolean;
    propose(agentId: string, note: string | null, edits: readonly ProposalEdit[]): Proposal;
    listProposals(): Proposal[];
    previewProposal(id: string): ProposalPreview;
    layoutProposalSlide(id: string, slideIndex: number): SlideDisplayList;
    layoutProposalDiffSlide(id: string, slideIndex: number): ProposalDiffSlide;
    acceptProposal(id: string, options?: {
        force?: boolean;
    }): ProposalAcceptance;
    rejectProposal(id: string): boolean;
    readonly clientId: number;
    fork(): PresentationHandle;
    /** Joins a shared session: table row and column insert/delete are refused from now on. Call it before the first edit on a handle that joins a shared session. */
    setCollaborative(): void;
    isCollaborative(): boolean;
    layouts(): Array<{
        id: string;
        name: string | null;
        type: string | null;
        masterId: string | null;
        placeholders: Array<{
            id: string;
            name: string;
            type: string | null;
            idx: number | null;
        }>;
    }>;
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
    readCommandContent(request?: PptxReadRequest): PresentationCommandReadResult;
    historyStep(direction: 'undo' | 'redo'): PresentationHistoryReceipt;
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
    /**
     * serialize the presentation back to .pptx bytes, edits included. Normal
     * autofit text the edits changed is fitted with the registered faces, or
     * with `options.fonts` alone when given.
     */
    save(options?: SavePresentationOptions): Uint8Array;
    insertText(storyId: string, index: number, text: string, style?: TextStyle): TextReceipt;
    /** `insertText` with the boundary's stage timings attached. */
    insertTextProfiled(storyId: string, index: number, text: string, style?: TextStyle): Profiled<TextReceipt>;
    deleteText(storyId: string, start: number, end: number): TextReceipt;
    deleteTextProfiled(storyId: string, start: number, end: number): Profiled<TextReceipt>;
    formatText(storyId: string, start: number, end: number, patch: TextStylePatch): TextReceipt;
    setStoryParagraphs(storyId: string, paragraphs: readonly TextParagraphDraft[]): StorySnapshot;
    insertParagraphBreak(storyId: string, index: number): TextReceipt;
    /** Sets the alignment of every paragraph the range touches; `null` restores
     *  the inherited value. */
    setParagraphAlignment(storyId: string, start: number, end: number, alignment: ParagraphAlignment | null): TextReceipt;
    insertTableRow(slideId: string, shapeId: string, index: number, position: 'before' | 'after', cells?: readonly string[]): TableReceipt;
    deleteTableRow(slideId: string, shapeId: string, index: number): TableReceipt;
    insertTableColumn(slideId: string, shapeId: string, index: number, position: 'before' | 'after', cells?: readonly string[], width?: 'split' | 'copy'): TableReceipt;
    deleteTableColumn(slideId: string, shapeId: string, index: number): TableReceipt;
    /** Throws the refusal the matching table edit would throw, without changing the deck. */
    checkTableEdit(method: 'insertTableRow' | 'deleteTableRow' | 'insertTableColumn' | 'deleteTableColumn', slideId: string, shapeId: string, index: number, position?: 'before' | 'after', cells?: readonly string[], width?: 'split' | 'copy'): void;
    insertSlide(index: number, layoutPartPath?: string): SlideReceipt;
    insertSlideProfiled(index: number, layoutPartPath?: string): Profiled<SlideReceipt>;
    duplicateSlide(slideId: string, toIndex: number): SlideReceipt;
    deleteSlide(slideId: string): SlideReceipt;
    moveSlide(slideId: string, toIndex: number): SlideReceipt;
    /** Sets a slide's speaker notes; empty text clears them. */
    setSlideNotes(slideId: string, text: string): void;
    setSlideLayout(slideId: string, layoutId: string, mapping?: Array<[string, string]>): void;
    addTextBox(slideId: string, draft: ShapeDraft): ShapeReceipt;
    addTextBoxProfiled(slideId: string, draft: ShapeDraft): Profiled<ShapeReceipt>;
    addShape(slideId: string, draft: PresetShapeDraft): ShapeReceipt;
    addPicture(slideId: string, draft: PictureDraft): ShapeReceipt;
    setShapeFill(slideId: string, shapeId: string, color: string | null): ShapeFillReceipt;
    setShapeStroke(slideId: string, shapeId: string, stroke: ShapeStroke): ShapeStrokeReceipt;
    setShapeAdjust(slideId: string, shapeId: string, adjustments: Record<string, number>): ShapeAdjustReceipt;
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
    addComment(slideId: string, comment: {
        author: string;
        initials?: string;
        text: string;
        created: string;
        xEmu?: number;
        yEmu?: number;
    }): CommentReceipt;
    /** Replies to a root comment in either flavor. */
    replyToComment(commentId: string, reply: {
        author: string;
        initials?: string;
        text: string;
        created: string;
    }): CommentReceipt;
    /** Resolves or reopens a modern comment. */
    setCommentStatus(commentId: string, resolved: boolean): CommentReceipt;
    /** Moves a root comment on its slide; coordinates are safe integer EMU. */
    setCommentPosition(commentId: string, position: {
        xEmu: number;
        yEmu: number;
    }): CommentReceipt;
    removeComment(commentId: string): CommentReceipt;
    /** Only legal while the deck has no comments. */
    setCommentFlavor(flavor: CommentFlavor): CommentFlavor;
    comments(): CommentSnapshot[];
    moveShape(slideId: string, shapeId: string, x: number, y: number): TransformReceipt;
    moveShapeProfiled(slideId: string, shapeId: string, x: number, y: number): Profiled<TransformReceipt>;
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
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
    dispose(): void;
}
/** @internal */
export declare class PresentationPeerError extends Error {
    readonly code: string;
    constructor(code: string, message: string);
}
declare const replayMethods: readonly ["insertText", "deleteText", "formatText", "setStoryParagraphs", "insertParagraphBreak", "setParagraphAlignment", "insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn", "insertSlide", "duplicateSlide", "deleteSlide", "moveSlide", "setSlideNotes", "setSlideLayout", "addTextBox", "addShape", "addPicture", "removeShape", "moveShape", "resizeShape", "setShapeRect", "setShapeFill", "setShapeStroke", "setShapeAdjust", "bringShapeToFront", "sendShapeToBack", "bringShapeForward", "sendShapeBackward", "addComment", "replyToComment", "setCommentStatus", "setCommentPosition", "removeComment", "setCommentFlavor", "propose", "acceptProposal", "rejectProposal", "applyEdits", "addUndoBoundary", "undo", "redo"];
/** @internal */
export type PresentationReplayOp = {
    [K in typeof replayMethods[number]]: {
        method: K;
        args: Parameters<PresentationHandle[K]>;
    };
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
export type PresentationPeerHandle = Pick<PresentationHandle, 'clientId' | 'snapshot' | 'story' | 'anchorCaret' | 'resolveCaretAnchor' | 'version' | 'readContent' | 'findText' | 'exportStructured' | 'exportMarkdown' | 'validateEdits' | 'searchText' | 'isProposalsAvailable' | 'listProposals' | 'previewProposal' | 'layoutProposalSlide' | 'layoutProposalDiffSlide' | 'layoutSlide' | 'hitTest' | 'mediaBytes' | 'save' | 'canUndo' | 'canRedo' | 'undoCaptureMode' | 'onUpdate' | 'dispose'>;
/** @internal */
export declare function openPresentationReplayBaseline(bytes: Uint8Array, options?: OpenPresentationOptions): PresentationPeerHandle;
/** @internal */
export declare function openPresentationPeerDeck(bytes: Uint8Array, identity: string, options?: Omit<OpenPresentationOptions, 'clientId'>): PresentationPeerHandle;
/** @internal */
export declare function registerPresentationPeerFonts(handle: PresentationPeerHandle): Promise<void>;
/** @internal */
export declare function adoptPresentationPeerIdentity(handle: PresentationPeerHandle): void;
/** @internal */
export declare function presentationPeerHydration(handle: PresentationPeerHandle): string;
/** @internal */
export declare function presentationPeerMetadata(handle: PresentationPeerHandle): PresentationMetadata;
/** @internal */
export declare function presentationPeerDisplayListJson(handle: PresentationPeerHandle, slideIndex: number): string;
/** @internal */
export declare function replayPresentation(handle: PresentationPeerHandle, envelope: PresentationReplayEnvelope, captured?: (reply: PresentationReplayReply) => void): PresentationReplayReply;
type PresentationMetadata = {
    slides: {
        id: string;
        index: number;
        name: string | null;
        layoutPartPath: string | null;
    }[];
    size: {
        width: number;
        height: number;
    };
};
type PptxStoryFormattingTarget = {
    slideId: string;
    shapeId: string;
    storyId: string;
};
export declare function presentationStoryFormatting(handle: PresentationHandle, target: PptxStoryFormattingTarget): StorySnapshot;
export declare function presentationDeleteTextJoiningParagraphs(handle: PresentationHandle, storyId: string, start: number, end: number): TextReceipt;
export declare function presentationMetadata(handle: PresentationHandle): PresentationMetadata;
/** @experimental */
export declare function presentationDisplayListJson(handle: PresentationHandle, slideIndex: number): string;
export declare function isProposalsAvailable(): boolean;
export declare function initWasm(input?: WasmInitInput): Promise<void>;
export declare function isWasmAvailable(): boolean;
export type EngineBackend = 'wasm' | 'native';
/** Which engine runs this presentation code ('native' when @betteroffice/pptx-native is installed and matches this version on Node, Bun or Deno, else 'wasm'). */
export declare function engineBackend(): EngineBackend;
export declare function wasmVersion(): string;
export declare function inspectPresentation(bytes: Uint8Array, options?: InspectPresentationOptions): unknown;
/**
 * Exports PPTX bytes as structured content whose anchors address the returned snapshot. No
 * session, DOM or font is involved. Throws `PptxExportError` for refused options and an `Error`
 * for bytes that are not a readable PPTX.
 */
export declare function exportPptxStructured(bytes: Uint8Array, options?: PptxExportOptions): Promise<PptxStructuredContent>;
/** `exportPptxStructured` rendered as Markdown with anchor markers. */
export declare function exportPptxMarkdown(bytes: Uint8Array, options?: PptxExportOptions): Promise<PptxMarkdownContent>;
/**
 * Renders structured content (schema version 1) as Markdown. Content the renderer cannot trust
 * refuses with `invalid-content`.
 */
export declare function renderPptxMarkdown(content: PptxStructuredContent, options?: PptxMarkdownOptions): Promise<PptxMarkdownContent>;
export declare function decodeTiffImage(bytes: Uint8Array): Uint8Array;
export declare function openPresentation(bytes: Uint8Array, options?: OpenPresentationOptions): PresentationHandle;
export {};
