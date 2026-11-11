// @betteroffice/docx declaration yrs/residentEngineWorkerClient.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { LayoutMetaV1 } from './layoutMeta';
import type { DocxCommandAck, DocxCommandEnvelope, DocxCommandFrame, DocxCommandClient, DocxCommandCheckpoint, DocxCommandInitReply, DocxCommandInitRequest, DocxCommandViewportRequest, DocxHostImage } from './residentCommandProtocol';
import type { YrsEngineApplyProfile, YrsResidentCaretSnapshot, YrsResidentWorkerSnapshot, YrsSelection } from './index';
import type { CollaborationCursor } from '../collaboration/types';
import type { CollaborationTextInsertion } from '../collaboration/types';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type { DocxDocumentRead, DocxDocumentReadValues, DocxEngineWorkerHostModule, DocxEngineWorkerRequest, ResidentEngineWorkerRequestWithoutId, DocxEngineWorkerResponse, ResidentProposalOperation, ResidentProposalResponse } from './residentEngineWorkerProtocol';
import type { DocxProposalRegistryState } from './proposals';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { type FontDecoder } from '../../../../shared/office-session/fontDecoder';
/** @internal */
export interface ResidentProposalReply extends Omit<ResidentProposalResponse, 'updates' | 'stateVector' | 'peerDiff'> {
    updates: Uint8Array[];
    peerDiff?: Uint8Array;
    stateVector: Uint8Array;
}
export interface ResidentEngineWorkerFrame {
    commandInitial?: DocxCommandInitReply;
    frame: Uint8Array;
    pageFrames?: Uint8Array[];
    updates: Uint8Array[];
    engineMs: number;
    workerTotalMs: number;
    engineProfile?: YrsEngineApplyProfile;
    caret: YrsResidentCaretSnapshot;
    selection: YrsSelection | null;
    /** The same selection as sticky positions, for the host to resolve against its content. */
    selectionCursor?: CollaborationCursor | null;
    /** The presented frame carries the worker-painted caret line. */
    caretPainted: boolean;
    replayMs: number;
    replayedPages: number;
    layoutRevision: number;
    /** The worker document version the frame lays out. */
    documentVersion?: string;
    documentPreview?: boolean;
    documentAsOpened?: boolean;
    /** Characters an applyDelete removed. */
    deletedUnits: number;
    /** The region layout the worker ran, when the request handed it the layout. */
    layoutJson?: string;
    layoutMeta?: LayoutMetaV1 | {
        v: number;
    };
    /** `layoutJson` covers only the first pages; `completeLayout` finishes it. */
    layoutProvisional?: boolean;
}
/** A bootstrap/sync whose snapshot layout the worker runs as the only layout. */
export interface ResidentEngineWorkerLayoutOptions {
    /** Display extras minus the header/footer payload the worker's layout supplies. */
    layoutExtras?: string;
    layoutReply?: 'meta';
    headersFootersEpoch?: number;
    /** The host state vector the snapshot brings the worker to. */
    stateVector?: Uint8Array;
    /** Lay out just the body's first pages before replying. */
    provisionalPages?: number;
    /** Bootstrap only: lay out the document {@link ResidentEngineWorkerClient.open} opened. */
    opened?: boolean;
    /** Bootstrap only: the epoch of the frame the host shows; the worker's frames follow it. */
    frameEpoch?: number;
    /**
     * Bootstrap only: the most the worker's editing core may allocate at once.
     * An allocation past it stops the worker with {@link ResidentWorkerOutOfMemoryError}.
     */
    heapLimitBytes?: number;
}
/** What an {@link ResidentEngineWorkerClient.open} seeded in the worker. */
export interface ResidentEngineWorkerOpened {
    /** The package's host metadata JSON, for `decodeDocxHostJson`. */
    hostJson: string;
    /** The worker replica's state vector once seeded. */
    stateVector: Uint8Array;
}
/** How a bootstrap or sync builds its frame. */
export interface ResidentEngineWorkerSnapshotOptions {
    /** Pages `[start, end)` the frame builds; the rest stay unbuilt. */
    displayWindow?: [number, number];
}
export interface ResidentEngineOffscreenPage {
    pageId: string;
    canvas: OffscreenCanvas;
}
export interface ResidentEngineWorkerApplyResult extends ResidentEngineWorkerFrame {
    applied: true;
}
type PendingRequest = {
    type: AwaitedRequest['type'];
    resolve(response: DocxEngineWorkerResponse & {
        ok: true;
    }): void;
    reject(error: Error): void;
};
type AwaitedRequest = Exclude<ResidentEngineWorkerRequestWithoutId, {
    type: 'applyUpdate' | 'eraseCaret' | 'prepare-save' | 'destroy';
}>;
/**
 * How long the worker may stay silent while requests wait. A large document
 * legitimately takes seconds per request, and the main thread would take just
 * as long, so this only catches a worker that stopped answering altogether.
 */
export declare const RESIDENT_WORKER_SILENCE_MS = 60000;
export interface ResidentEngineWorkerPort {
    onmessage: ((event: MessageEvent<DocxEngineWorkerResponse>) => void) | null;
    onerror: ((event: ErrorEvent) => void) | null;
    onmessageerror: ((event: MessageEvent) => void) | null;
    postMessage(message: DocxEngineWorkerRequest | DocxEngineWorkerHostModule, transfer?: Transferable[]): void;
    terminate(): void;
}
/** Dedicated-worker owner for resident input, pagination, and FrameDelta output. */
export declare class ResidentEngineWorkerClient implements DocxCommandClient {
    protected readonly worker: ResidentEngineWorkerPort;
    readonly fontRuntime: {
        decoder: FontDecoder;
    };
    decodeFontFace(bytes: ArrayBuffer | Uint8Array, { codec, dict, maxOutput }: Parameters<FontDecoder>[1]): ReturnType<FontDecoder>;
    protected readonly pending: Map<number, PendingRequest>;
    private watchdog;
    private nextId;
    protected terminalError: Error | null;
    private commandSession;
    private commandGenerationValue;
    private readonly commandInitialGenerations;
    private commandReplacementPending;
    private ready;
    protected revision: number;
    protected documentSequence: number;
    private remoteVector;
    private appliedFontsRevision;
    protected bootstrapped: boolean;
    /** Set once `open` is sent, with the heap limit it opened under. */
    protected openedHeapLimit: {
        bytes?: number;
    } | null;
    private bootstraps;
    /** The font-requirements read still in flight, shared while nothing was posted after it. */
    private fontRead;
    /** Id of the last snapshot request sent; replies to earlier requests must
     * not replace the state it recorded. */
    private lastSnapshotId;
    private lastFailedFontsSnapshotId;
    private keepSurfaces;
    private bootstrapWaiters;
    private lastMemory;
    private answeredFrameEpoch;
    private retainBuiltPages;
    private failureListener;
    private readonly frameListeners;
    private readonly pipeListeners;
    private readonly collabListeners;
    protected commandIdentity: {
        session: string;
        incarnation: number;
        epoch: number;
    } | null;
    constructor(worker?: ResidentEngineWorkerPort);
    isReady(): boolean;
    get commandGeneration(): number;
    get commandFailure(): Error | undefined;
    commandInitialGeneration(initial: DocxCommandInitReply): number | undefined;
    commandInit(request: DocxCommandInitRequest): Promise<DocxCommandInitReply>;
    edit(command: DocxCommandEnvelope): Promise<DocxCommandAck>;
    checkpoint(): Promise<DocxCommandCheckpoint>;
    requestFullFrame(): Promise<DocxCommandAck>;
    /** Hands the worker pictures the host decoded after `commandInit`; the displayed pages repaint with them. */
    provideHostImages(images: DocxHostImage[]): Promise<void>;
    requestViewport(request: DocxCommandViewportRequest): Promise<void>;
    onFrame(listener: (frame: DocxCommandFrame) => void): () => void;
    /** Output of the collaboration providers the worker runs, per provider, in order. */
    onCollabPipe(listener: (batches: readonly PipeOutputBatch[]) => void): () => void;
    onCollabUpdates(listener: (event: {
        updates: Uint8Array[];
        inferences: Array<{
            seq: number;
            insertion: CollaborationTextInsertion;
        }>;
        clock?: number;
    }) => void): () => void;
    applyRemoteUpdate(update: Uint8Array): Promise<void>;
    private commandAck;
    /** @internal */
    hasFailed(): boolean;
    /** The worker's wasm memories as of its latest reply; null before one. */
    memory(): WasmModuleMemory[] | null;
    layoutRevision(): number;
    /** @internal */
    stateSequence(): number;
    /** @internal The newest frame epoch a reply carried; 0 before any frame. */
    answeredFrame(): number;
    /** @internal */
    setRetainBuiltPages(retain: boolean): void;
    /** @internal Called once when the worker fails, whether or not a request was waiting; not on `destroy`. */
    onFailure(listener: ((error: Error) => void) | null): void;
    /** @internal Whether foreground document or frame work awaits its reply; `reads: false` leaves document reads out. */
    frameRequestPending(reads?: boolean): boolean;
    /** The worker replica's last reported yrs state vector (null before any). */
    remoteStateVector(): Uint8Array | null;
    /** The fonts revision this worker last applied (null before bootstrap). */
    syncedFontsRevision(): number | null;
    /** A bootstrap was sent; later snapshots go as syncs queued behind it. */
    bootstrapSent(): boolean;
    /** Resolves once a bootstrap is sent, at once when one was. */
    whenBootstrapSent(): Promise<void>;
    /**
     * Starts over with another document in the same worker: the next request
     * is a bootstrap, which keeps the page surfaces already attached so the
     * new document's pages paint where the old ones were.
     */
    rebootstrap(): void;
    /**
     * Parses and seeds a DOCX in the worker, which then holds the document:
     * its first layout is a bootstrap with `opened`, and the main replica loads
     * {@link encodeState}. A copy of `bytes` is transferred.
     */
    open(bytes: Uint8Array, options?: {
        digest?: string;
        generation?: string;
        heapLimitBytes?: number;
    }): Promise<ResidentEngineWorkerOpened>;
    /**
     * Opens a display-only preview of the first `blocks` body blocks of a DOCX in the worker, or
     * resolves null when the package cannot open as a preview. Its first layout is a bootstrap
     * with `opened`; {@link rebootstrap} then lets {@link open} replace it with the whole document.
     * A copy of `bytes` is transferred.
     */
    openPreview(bytes: Uint8Array, blocks: number, options?: {
        heapLimitBytes?: number;
        paragraphBudget?: number;
    }): Promise<ResidentEngineWorkerOpened | null>;
    /** The font requirements of a region layout request, read from the opened document. */
    fontRequirements(layoutInput: string): Promise<string>;
    layoutJson(layoutRevision: number): Promise<{
        status: 'ok';
        layoutJson: string;
    } | {
        status: 'stale';
    }>;
    /** @internal */
    proposal(operation: ResidentProposalOperation): Promise<ResidentProposalReply>;
    /** @internal */
    documentRead<K extends DocxDocumentRead['kind']>(read: DocxDocumentRead & {
        kind: K;
    }): Promise<{
        version: string;
        value: DocxDocumentReadValues[K];
    }>;
    /**
     * Reads the document at `expectVersion`. A later message may run first; the
     * read then answers `superseded` when the document moved on meanwhile.
     * @internal
     */
    documentReadAt<K extends DocxDocumentRead['kind']>(read: DocxDocumentRead & {
        kind: K;
    }, expectVersion: string): Promise<{
        status: 'ok';
        version: string;
        value: DocxDocumentReadValues[K];
    } | {
        status: 'superseded';
    }>;
    /** @internal */
    handOver(peerMetadata?: true): Promise<{
        state: Uint8Array;
        metadata?: Uint8Array;
        metadataReason?: string;
        version: string;
        proposals: DocxProposalRegistryState;
    }>;
    /** The worker's whole document state as one yrs v1 update. */
    encodeState(): Promise<Uint8Array>;
    /** @internal */
    encodeVersionedState(peerMetadata?: true): Promise<{
        state: Uint8Array;
        version: string | undefined;
        metadata?: Uint8Array;
        metadataReason?: string;
    }>;
    /**
     * @internal Saves the opened document in the worker. With the editor copy's
     * `stateVector`, `updates` carry what the worker holds beyond it (the
     * paragraph IDs the save wrote among them) for that copy to integrate;
     * `version` is the document version after the save.
     */
    save(request: {
        comments: Comment[];
        /** @internal */
        commentsSource?: 'host';
        host?: Document;
        stateVector?: Uint8Array;
        /** @internal */
        stories?: readonly string[];
    }): Promise<{
        bytes: ArrayBuffer;
        updates: Uint8Array[];
        version: string;
    }>;
    prepareSave(): void;
    revisionCount(): Promise<number>;
    warm(hostModule?: boolean): Promise<void>;
    bootstrap(snapshot: YrsResidentWorkerSnapshot, extras: string, options?: ResidentEngineWorkerLayoutOptions & ResidentEngineWorkerSnapshotOptions): Promise<ResidentEngineWorkerFrame>;
    sync(snapshot: YrsResidentWorkerSnapshot, extras: string, expectedFrameEpoch: number, paintCaret?: boolean, options?: ResidentEngineWorkerLayoutOptions & ResidentEngineWorkerSnapshotOptions): Promise<ResidentEngineWorkerFrame>;
    /**
     * Lay out the rest of a provisional bootstrap layout: its frame and full
     * layout, or null when a later snapshot already replaced it.
     */
    completeLayout(expectedFrameEpoch: number, paintCaret?: boolean, sliceBlocks?: number): Promise<ResidentEngineWorkerFrame | null>;
    buildFrame(extras: string, expectedFrameEpoch: number, paintCaret?: boolean, displayWindow?: [number, number]): Promise<ResidentEngineWorkerFrame>;
    /** Build unbuilt display pages; the reply frame carries them. */
    buildPages(pages: number[], expectedFrameEpoch: number, paintCaret?: boolean): Promise<ResidentEngineWorkerFrame>;
    /** @internal */
    buildPages(pages: number[], expectedFrameEpoch: number, paintCaret: boolean, background: boolean): Promise<ResidentEngineWorkerFrame | null>;
    releasePages(pages: Array<{
        index: number;
        pageId: string;
    }>, expectedFrameEpoch: number, paintCaret?: boolean): Promise<ResidentEngineWorkerFrame | {
        superseded: true;
    }>;
    applyInput(text: string, selection: YrsSelection, expectedFrameEpoch: number, profile?: boolean, paintCaret?: boolean, displayWindow?: [number, number]): Promise<ResidentEngineWorkerApplyResult | {
        applied: false;
    }>;
    applyDelete(direction: 'backward' | 'forward', selection: YrsSelection, expectedFrameEpoch: number, profile?: boolean, paintCaret?: boolean, count?: number, displayWindow?: [number, number]): Promise<ResidentEngineWorkerApplyResult | {
        applied: false;
    }>;
    /** Drop the worker-painted caret line by re-presenting the caret page's
     * retained raster. Fire-and-forget and idempotent. */
    eraseCaret(): void;
    invalidate(update: Uint8Array, selection: YrsSelection | null): void;
    syncUpdate(update: Uint8Array, stateVector: Uint8Array): Promise<{
        version: string;
        stateVector: Uint8Array;
        repair: Uint8Array | null;
    }>;
    attachCanvases(pages: ResidentEngineOffscreenPage[], activePageIds: string[], devicePixelRatio: number, zoom: number, caretStyle: ResidentCaretPaintStyle): Promise<void>;
    destroy(): void;
    protected request(request: AwaitedRequest, transfer?: Transferable[]): Promise<DocxEngineWorkerResponse & {
        ok: true;
    }>;
    protected fenceCommands(): void;
    /**
     * Requests run in order, so once a snapshot is sent every later request
     * finds its state and fonts in the worker: the next sync can diff against
     * them before this one answers.
     */
    private recordSent;
    /** Restarted by every message, so a queue of slow but answered requests never trips it. */
    private armWatchdog;
    private disarmWatchdog;
    /** Record a successfully applied bootstrap/sync payload's fonts revision.
     * The state vector is tracked centrally in `onmessage`. */
    private recordSync;
    protected fail(error: Error, notify?: boolean): void;
}
declare class ResidentWorkerUnavailableError extends Error {
}
/** The worker itself failed (crash, timeout, torn-down, corrupt reply). */
export declare class ResidentWorkerFailureError extends Error {
}
/** The worker has no whole opened package to save. @internal */
export declare class DocxWorkerSaveUnavailableError extends ResidentWorkerFailureError {
}
/** A whole source failed to open and the worker still holds the source cut it was to replace. @internal */
export declare class ResidentWorkerCutKeptError extends Error {
}
/** The worker trapped because its wasm memory could not grow any further. */
export declare class ResidentWorkerOutOfMemoryError extends ResidentWorkerUnavailableError {
    /** The worker's wasm memories when it trapped. */
    readonly memory: WasmModuleMemory[];
    constructor(message: string, 
    /** The worker's wasm memories when it trapped. */
    memory: WasmModuleMemory[]);
}
export declare function canUseResidentEngineWorker(): boolean;
/** @internal */
export declare function preloadResidentEngineWorker(): Promise<void>;
/** @internal */
export declare function retainPreloadedResidentEngineWorker(): () => void;
/** @internal */
export declare function takePreloadedResidentEngineWorker(): ResidentEngineWorkerClient | null;
export {};
