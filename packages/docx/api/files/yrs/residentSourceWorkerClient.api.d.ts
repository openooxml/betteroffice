// @betteroffice/docx declaration yrs/residentSourceWorkerClient.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { ResidentEngineWorkerClient, type ResidentEngineWorkerPort, type ResidentEngineWorkerFrame } from './residentEngineWorkerClient';
import { type DocxSourceOpenRequest, type DocxSourcePreviewOptions } from './residentSource';
import type { DocxFontEnvironmentEvent } from './residentFontEnvironment';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
import type { Document } from '../types/document';
import type { DocxSvgMedia } from './residentCommandProtocol';
export interface DocxEngineWorkerSourceOpened {
    document: Document;
    stateVector: Uint8Array;
    hostJson?: string;
    preview?: ResidentEngineWorkerFrame;
    /** The client id the opened replica writes with: a fresh one when the pinned id would reuse its clocks. */
    clientId?: number;
    /** The opened state's clock for `clientId`. */
    clientClock?: number;
    /** Pictures shown as SVG, which the host decodes for the worker's `commandInit`. */
    svgMedia?: DocxSvgMedia[];
}
export declare class DocxSourceWorkerClient extends ResidentEngineWorkerClient {
    private fontListeners;
    private memoryListeners;
    private requirementsListeners;
    private presentedFontsRevision;
    private cutReservation;
    constructor(worker?: ResidentEngineWorkerPort);
    static adopt(client: ResidentEngineWorkerClient): DocxSourceWorkerClient;
    private observe;
    /** Opens a display-only cut of the first body blocks for a preview; null when the worker declines it. {@link openSource} replaces it. */
    openSourceCut(source: DocxSourceOpenRequest & {
        bytes: Uint8Array;
        cut: NonNullable<DocxSourceOpenRequest['cut']>;
    }): Promise<Document | null>;
    private releaseCut;
    openSource(source: DocxSourceOpenRequest): Promise<DocxEngineWorkerSourceOpened>;
    /** A preview frame of the opened source's first pages, laid out with the fonts its command session will use. */
    previewSource(preview: DocxSourcePreviewOptions & {
        layoutInput: string;
    }): Promise<ResidentEngineWorkerFrame>;
    prewarm(): Promise<void>;
    onFontEnvironment(listener: (event: DocxFontEnvironmentEvent) => void): () => void;
    /**
     * @internal Called when an applied edit may need fonts the last `fontRequirements` answer lacked,
     * or with `wholeFaces` when the engine refused a sliced face's outlines and needs whole faces.
     */
    onFontRequirementsHint(listener: (hint: {
        wholeFaces: boolean;
    }) => void): () => void;
    onMemory(listener: (samples: WasmModuleMemory[]) => void): () => void;
    frameRequestPending(reads?: boolean): boolean;
    protected request(...args: Parameters<ResidentEngineWorkerClient['request']>): Promise<{
        id: number;
        ok: true;
        decodedFontFace?: {
            font: ArrayBuffer;
            payload?: ArrayBuffer;
        };
        commandInitial?: import("./residentCommandProtocol").DocxCommandInitReply;
        commandAck?: import("./residentCommandProtocol").DocxCommandAck;
        commandFrame?: {
            session: string;
            incarnation: number;
            epoch: number;
            frame: import("./residentCommandProtocol").DocxCommandFrame;
        };
        fontEnvironmentEvent?: DocxFontEnvironmentEvent & {
            session: string;
            incarnation: number;
            epoch: number;
        };
        fontRequirementsHint?: {
            session: string;
            incarnation: number;
            epoch: number;
            wholeFaces?: true;
        };
        fontEnvironment?: import("./residentFontEnvironment").DocxFontEnvironmentSnapshot;
        checkpoint?: import("../../../../shared/office-session/commands/types").Checkpoint;
        frame?: ArrayBuffer;
        pageFrames?: ArrayBuffer[];
        pageBuildSuperseded?: boolean;
        superseded?: true;
        updates?: ArrayBuffer[];
        engineMs?: number;
        workerTotalMs?: number;
        engineProfile?: import(".").YrsEngineApplyProfile;
        caret?: import(".").YrsResidentCaretSnapshot;
        selection?: import(".").YrsSelection | null;
        selectionCursor?: import("../collaboration").CollaborationCursor | null;
        caretPainted?: boolean;
        replayMs?: number;
        replayedPages?: number;
        layoutRevision?: number;
        documentVersion?: string;
        documentPreview?: boolean;
        documentAsOpened?: boolean;
        deletedUnits?: number;
        stateVector?: ArrayBuffer;
        layoutJson?: string;
        layoutMeta?: import("./layoutMeta").LayoutMetaV1 | {
            v: number;
        };
        layoutJsonStatus?: "ok" | "stale";
        layoutProvisional?: boolean;
        hostJson?: string;
        sourceDocument?: Document;
        embeddedFontFaces?: import("../utils").EmbeddedFontFace[];
        svgMedia?: DocxSvgMedia[];
        clientId?: number;
        clientClock?: number;
        previewRefused?: boolean;
        requirementsJson?: string;
        state?: ArrayBuffer;
        peerMetadata?: ArrayBuffer;
        peerMetadataReason?: string;
        saved?: ArrayBuffer;
        revisionCount?: number;
        proposals?: import("./proposals").DocxProposalRegistryState;
        version?: string;
        repair?: ArrayBuffer | null;
        proposal?: import("./residentEngineWorkerProtocol").ResidentProposalResponse;
        read?: {
            version: string;
            value: unknown;
        };
        progress?: number;
    } & {
        collabUpdates?: import("./residentCollaboration").ResidentCollaborationUpdates & {
            session?: string;
            incarnation?: number;
            epoch?: number;
        };
        memory?: WasmModuleMemory[];
    } & {
        ok: true;
    }>;
    protected fenceCommands(): void;
    protected fail(error: Error, notify?: boolean): void;
}
export declare function takePreloadedDocxSourceWorker(): DocxSourceWorkerClient | null;
