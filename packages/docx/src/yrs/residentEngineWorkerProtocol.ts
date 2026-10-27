import type { DocxRangeHighlights } from './residentRangeHighlights';
import type { DocxPresencePeer, DocxPresenceGeometry, ResidentCollaborationUpdates } from './residentCollaboration';
export type { DocxPresencePeer, DocxPresenceGeometry } from './residentCollaboration';
import type { ResidentPartGeometry, ResidentPartMetadata } from './residentParts';
import type { ResidentControlSnapshot } from './residentControls';
import type { ResidentPropertyTarget, ResidentPropertySnapshot } from './residentProperties';
import type { ResidentSectionSnapshot } from './residentSections';
import type { YrsStoryRange } from './index';
import type { DocxReviewSnapshot, DocxReviewTarget, DocxReviewTargetRanges } from './residentReview';
import type { ResidentCommentInsertionRange } from './residentCommentProjection';
import type { DocxToolbarMetadata } from './residentToolbarMetadata';
import type { LayoutMetaV1 } from './layoutMeta';
import type { Checkpoint } from '../../../../shared/office-session/commands/types';
import type { ReplyError } from '../../../../shared/office-session/protocol';
import type { SessionFailureCode } from '../../../../shared/office-session/types';
import type {
  DocxCommandAck,
  DocxCommandEnvelope,
  DocxCommandFrame,
  DocxCommandSelection,
  DocxSelectionPointIntent,
  DocxCommandInitReply,
  DocxCommandInitRequest,
  DocxCommandViewportRequest,
  DocxHostImage,
  DocxSvgMedia,
} from './residentCommandProtocol';
import type {
  YrsEngineApplyProfile,
  YrsLoc,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
  YrsSelectionText,
  YrsStickyPosition,
} from './index';
import type { DocxClipboardSlice } from './clipboardFragment';
import type { PointPosition } from '../plugin-api';
import type { DocxResolvedPointPosition } from './pointPosition';
import type {
  DocxDisplayRange,
  DocxDisplaySelectionInfo,
  DocxDisplaySelectionText,
  DocxSelectionUnit,
} from './viewerSelection';
import type { FindOptions } from '../utils/findReplace';
import type { DocxFindDisplayMatch } from './findMatches';
import type { CollaborationCursor } from '../collaboration/types';
import type { DocxSidebarRead, DocxOutlineHeading } from './sidebarReads';
import type { ResidentSearchResult } from './residentSearch';
import type { DocxFindParagraphsOptions, DocxParagraphMatch } from './findParagraphs';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import type {
  DocxProposalRegistryState,
  DocxProposalRequest,
  DocxProposalResult,
  DocxProposalStateRequest,
  DocxProposalWithdrawRequest,
} from './proposals';
import type {
  DocxParagraphAnchor,
  DocxParagraphAnchorResult,
  DocxParagraphIdentitySnapshot,
} from './paragraphIdentity';
import type {
  DocxEditRequest,
  DocxValidationResult,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  DocxReadStoriesRequest,
  DocxReadStoriesResult,
} from './edits';
import type { DocxListStoriesResult } from './stories';
import type {
  AnchorGeometryTarget,
  ProposalGeometryMirror,
  ProposalGeometryTarget,
  resolveNavigationTarget,
} from './proposalGeometry';
import type { DocxPageExportOptions } from './pagedExport';
import type { DocxContentControlQuery, DocxContentControlsOptions, DocxContentControlsResult } from './contentControls';
import type { DocxSourceOpenRequest, DocxSourcePreviewOptions } from './residentSource';
import type { DocxFontEnvironmentEvent, DocxFontEnvironmentSnapshot } from './residentFontEnvironment';

/** @internal */
export interface ResidentEngineWorkerFontSync {
  fontsBaseRevision?: number;
}

/** @internal */
export type ResidentProposalOperation =
  | { kind: 'propose'; request: DocxProposalRequest; peerStateVector?: Uint8Array }
  | { kind: 'setStates'; request: DocxProposalStateRequest; peerStateVector?: Uint8Array }
  | { kind: 'withdraw'; request: DocxProposalWithdrawRequest; peerStateVector?: Uint8Array }
  | { kind: 'removeComment'; id: string }
  | { kind: 'snapshot' };

/** @internal */
export interface ResidentProposalResponse {
  result?: DocxProposalResult;
  mirror: { version: string; proposals: DocxProposalRegistryState };
  changedStories: string[];
  projectionStories?: string[];
  updates: ArrayBuffer[];
  peerDiff?: ArrayBuffer;
  stateVector: ArrayBuffer;
  geometry: ProposalGeometryMirror;
  fontRequirements?: { layoutInput: string; requirementsJson: string };
}

/** @internal */
export type DocxDocumentRead =
  | { kind: 'collabStateVector' }
  | { kind: 'collabDiff'; stateVector?: Uint8Array }
  | { kind: 'collabDiffChunks'; stateVector: Uint8Array; maxChunkBytes: number }
  | { kind: 'collabPresence'; peers: DocxPresencePeer[]; pages: [number, number]; expectVersion: string; expectFrameEpoch: number; expectLayoutRevision: number }
  | { kind: 'validateEdits'; request: DocxEditRequest }
  | { kind: 'contentControlHandles' }
  | { kind: 'partSnapshot' }
  | { kind: 'properties'; target: ResidentPropertyTarget }
  | { kind: 'sectionProperties' }
  | { kind: 'exportStructuredWithPages'; options: DocxPageExportOptions; currentRequest: string; workerFonts?: boolean }
  | { kind: 'listContentControls'; options: DocxContentControlsOptions }
  | { kind: 'findContentControls'; query: DocxContentControlQuery; options: DocxContentControlsOptions }
  | { kind: 'rangeHighlights'; ranges: YrsStoryRange[]; expectVersion: string; expectFrameEpoch?: number; expectLayoutRevision?: number }
  | { kind: 'reviewSnapshot'; expectPreviewVersion?: number }
  | { kind: 'commentInsertionRange'; paraId: string; search?: string }
  | { kind: 'commentSelectionRange'; selection: DocxCommandSelection }
  | { kind: 'toolbarMetadata' }
  | { kind: 'paragraphIdentities' }
  | { kind: 'resolveParagraphAnchors'; anchors: DocxParagraphAnchor[] }
  | { kind: 'readParagraphs'; request: DocxReadParagraphsRequest }
  | { kind: 'rawSelectionInfo'; selection: DocxCommandSelection }
  | { kind: 'clipboardSlice'; selection: DocxCommandSelection }
  | { kind: 'listStories' }
  | { kind: 'readStories'; request: DocxReadStoriesRequest }
  | { kind: 'findText'; request: DocxFindTextRequest }
  | { kind: 'findMatches'; searchText: string; options: FindOptions; expectVersion: string }
  | { kind: 'searchText'; query: string; caseSensitive: boolean; carry?: YrsStickyPosition | null }
  | ({ kind: 'findParagraphs'; query: string } & DocxFindParagraphsOptions)
  | { kind: 'stickyAnchors'; locs: YrsLoc[]; version: string }
  | { kind: 'navigationTarget'; story: string; paraId: string }
  | { kind: 'bookmarkTarget'; story: string; name: string }
  | { kind: 'anchorTargets'; targets: Exclude<AnchorGeometryTarget, { kind: 'proposal' }>[] }
  | { kind: 'reviewTarget'; target: DocxReviewTarget }
  | { kind: 'pointPosition'; hit: PointPosition; expectVersion: string; nativePoint?: DocxSelectionPointIntent }
  | {
      kind: 'selectionUnit';
      story: string;
      position: number;
      unit: DocxSelectionUnit;
      expectVersion: string;
    }
  | { kind: 'selectionText'; story: string; anchor: number; head: number; expectVersion: string }
  | { kind: 'selectionInfo'; story: string; anchor: number; head: number; expectVersion: string }
  | { kind: 'paragraphTarget'; story: string; paraId: string; expectVersion: string }
  | { kind: 'commentTarget'; story: string; commentId: string; expectVersion: string }
  | { kind: 'revisionTarget'; story: string; revisionId: string; expectVersion: string }
  | { kind: 'bookmarkPosition'; story: string; name: string; expectVersion: string }
  | { kind: 'sidebar'; commentIds: string[]; expectVersion: string }
  | { kind: 'headings'; expectVersion: string }
  | { kind: 'adjacentRevision'; selection: DocxCommandSelection | null; direction: 'next' | 'previous' };

/** @internal */
export interface DocxDocumentReadValues {
  collabStateVector: Uint8Array;
  collabDiff: Uint8Array;
  collabDiffChunks: Uint8Array[];
  collabPresence: DocxPresenceGeometry | null;
  validateEdits: DocxValidationResult;
  contentControlHandles: ResidentControlSnapshot;
  partSnapshot: { metadata: ResidentPartMetadata; geometry: ResidentPartGeometry[]; frameEpoch: number | null; layoutRevision: number };
  properties: ResidentPropertySnapshot | null;
  sectionProperties: ResidentSectionSnapshot;
  exportStructuredWithPages: string;
  listContentControls: DocxContentControlsResult;
  findContentControls: DocxContentControlsResult;
  rangeHighlights: DocxRangeHighlights | null;
  reviewSnapshot: DocxReviewSnapshot | null;
  commentInsertionRange: ResidentCommentInsertionRange;
  commentSelectionRange: YrsStoryRange | null;
  toolbarMetadata: DocxToolbarMetadata;
  paragraphIdentities: DocxParagraphIdentitySnapshot;
  resolveParagraphAnchors: { results: DocxParagraphAnchorResult[] };
  readParagraphs: DocxReadParagraphsResult;
  rawSelectionInfo: YrsSelectionText;
  clipboardSlice: DocxClipboardSlice;
  listStories: DocxListStoriesResult;
  readStories: DocxReadStoriesResult;
  findText: DocxFindTextResult;
  findMatches: DocxFindDisplayMatch[] | null;
  navigationTarget: ReturnType<typeof resolveNavigationTarget>;
  bookmarkTarget: { loc: YrsLoc; position: number } | null;
  anchorTargets: ProposalGeometryTarget[];
  reviewTarget: DocxReviewTargetRanges | null;
  searchText: ResidentSearchResult;
  findParagraphs: DocxParagraphMatch[];
  stickyAnchors: Array<YrsStickyPosition | null>;
  pointPosition: DocxResolvedPointPosition | null;
  selectionUnit: DocxDisplayRange | null;
  selectionText: DocxDisplaySelectionText | null;
  selectionInfo: DocxDisplaySelectionInfo | null;
  paragraphTarget: DocxDisplayRange | null;
  commentTarget: DocxDisplayRange | null;
  revisionTarget: DocxDisplayRange | null;
  bookmarkPosition: number | null;
  sidebar: DocxSidebarRead | null;
  headings: DocxOutlineHeading[] | null;
  adjacentRevision: { revisionId: string } | null;
}

/** How long a warm waits for the host's compiled module before loading the engine itself. */
export const RESIDENT_HOST_MODULE_WAIT_MS = 10_000;

/** @internal */
export type DocxEngineWorkerHostModule = {
  type: 'editModule';
  module: WebAssembly.Module | null;
};

export type DocxEngineWorkerRequest =
  | { id: number; type: 'warm'; hostModule?: true }
  | { id: number; type: 'openSource'; source: DocxSourceOpenRequest }
  | { id: number; type: 'previewSource'; preview: DocxSourcePreviewOptions & { layoutInput: string } }
  | { id: number; type: 'commandInit'; request: DocxCommandInitRequest }
  | { id: number; type: 'edit'; command: DocxCommandEnvelope }
  | { id: number; type: 'checkpoint' }
  | { id: number; type: 'requestFullFrame' }
  | { id: number; type: 'commandUpdate'; update: Uint8Array }
  | { id: number; type: 'commandViewport'; request: DocxCommandViewportRequest }
  | {
      id: number;
      type: 'bootstrap';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      layoutExtras?: string;
      layoutReply?: 'meta';
      headersFootersEpoch?: number;
      /** Pages `[start, end)` a full build compiles; the rest stay unbuilt. */
      displayWindow?: [number, number];
      retainBuiltPages?: boolean;
      /**
       * Lay out only as much of the body as fills this many pages; a reply
       * marked `layoutProvisional` is finished by `completeLayout`.
       */
      provisionalPages?: number;
      /**
       * The document replaces the one this worker showed, which the host
       * hands over from: keep the attached page surfaces for its pages.
       */
      keepSurfaces?: boolean;
      /** Lay out the document `open` seeded here, not the snapshot's state. */
      opened?: boolean;
      /** The most the worker's editing core may allocate at once. */
      heapLimitBytes?: number;
    }
  | {
      id: number;
      type: 'open';
      /** The DOCX package, parsed and seeded in a fresh session here. */
      bytes: ArrayBuffer;
      /** The package's SHA-256, when the caller already took it. */
      digest?: string;
      generation?: string;
      /** The most the worker's editing core may allocate at once. */
      heapLimitBytes?: number;
      /**
       * Opens a display-only preview of the first `previewBlocks` body blocks instead. A later
       * `open` of the whole document replaces it.
       */
      previewBlocks?: number;
      /** Optional paragraph weight budget for the display-only body cut. */
      previewParagraphBudget?: number;
    }
  | { id: number; type: 'fontRequirements'; layoutInput: string }
  | { id: number; type: 'layoutJson'; layoutRevision: number }
  | { id: number; type: 'encodeState'; peerMetadata?: true }
  | {
      id: number;
      type: 'save';
      comments: Comment[];
      /** `host` when `comments` is the host's list and wins over the worker's own comments. @internal */
      commentsSource?: 'host';
      host?: Document;
      stateVector?: Uint8Array;
      /** The editor peer's marked stories. @internal */
      stories?: readonly string[];
    }
  | { id: number; type: 'syncUpdate'; update: Uint8Array; stateVector: Uint8Array }
  | { id: number; type: 'revisionCount' }
  | { id: number; type: 'proposal'; operation: ResidentProposalOperation }
  | {
      id: number;
      type: 'documentRead';
      read: DocxDocumentRead;
      /** Answered `superseded` instead when the document's version differs. */
      expectVersion?: string;
    }
  | {
      id: number;
      type: 'sync';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
      /**
       * Display extras without the header/footer payload. When present, the
       * snapshot's layout is authoritative: the worker completes the extras
       * from the layout it runs and returns that layout as `layoutJson`.
       */
      layoutExtras?: string;
      layoutReply?: 'meta';
      headersFootersEpoch?: number;
      displayWindow?: [number, number];
      retainBuiltPages?: boolean;
      provisionalPages?: number;
    }
  | {
      id: number;
      type: 'buildPages';
      pages: number[];
      expectedFrameEpoch: number;
      paintCaret: boolean;
      background?: boolean;
      /** Pictures the host decoded after `commandInit`; the displayed pages repaint with them. Transferred. */
      hostImages?: DocxHostImage[];
      /** The command session `hostImages` were decoded for; any other session closes them unused. */
      hostImagesFor?: { session: string; incarnation: number; epoch: number };
    }
  | {
      id: number;
      type: 'releasePages';
      pages: Array<{ index: number; pageId: string }>;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'completeLayout';
      expectedFrameEpoch: number;
      paintCaret: boolean;
      /**
       * Measure the rest this many body blocks a step to start with, letting
       * requests that arrive meanwhile run between steps; absent, in one step.
       */
      sliceBlocks?: number;
    }
  | {
      id: number;
      type: 'buildFrame';
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
      displayWindow?: [number, number];
      retainBuiltPages?: boolean;
    }
  | {
      id: number;
      type: 'applyInput';
      text: string;
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
      displayWindow?: [number, number];
      retainBuiltPages?: boolean;
    }
  | {
      id: number;
      type: 'applyDelete';
      direction: 'backward' | 'forward';
      count: number;
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
      displayWindow?: [number, number];
      retainBuiltPages?: boolean;
    }
  | {
      id: number;
      type: 'applyUpdate';
      update: Uint8Array;
      selection: YrsSelection | null;
    }
  | {
      id: number;
      type: 'attachCanvases';
      pages: Array<{ pageId: string; canvas: OffscreenCanvas }>;
      activePageIds: string[];
      devicePixelRatio: number;
      zoom: number;
      caretStyle: ResidentCaretPaintStyle;
    }
  | { id: number; type: 'eraseCaret' }
  | { id: number; type: 'destroy' };

export type ResidentEngineWorkerRequestWithoutId = DocxEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

export type DocxEngineWorkerResponse = (
  | {
      id: number;
      ok: true;
      commandInitial?: DocxCommandInitReply;
      commandAck?: DocxCommandAck;
      commandFrame?: { session: string; incarnation: number; epoch: number; frame: DocxCommandFrame };
      fontEnvironmentEvent?: DocxFontEnvironmentEvent & { session: string; incarnation: number; epoch: number };
      fontEnvironment?: DocxFontEnvironmentSnapshot;
      checkpoint?: Checkpoint;
      frame?: ArrayBuffer;
      pageFrames?: ArrayBuffer[];
      pageBuildSuperseded?: boolean;
      superseded?: true;
      updates?: ArrayBuffer[];
      engineMs?: number;
      workerTotalMs?: number;
      engineProfile?: YrsEngineApplyProfile;
      caret?: YrsResidentCaretSnapshot;
      selection?: YrsSelection | null;
      /** The same selection as sticky positions, for the host to resolve against its content. */
      selectionCursor?: CollaborationCursor | null;
      /** The presented frame carries the worker-painted caret line. */
      caretPainted?: boolean;
      replayMs?: number;
      replayedPages?: number;
      layoutRevision?: number;
      /** The document version the frame lays out. */
      documentVersion?: string;
      documentPreview?: boolean;
      /** The frame lays out the whole document as opened, before any change. */
      documentAsOpened?: boolean;
      /** Characters an applyDelete removed. */
      deletedUnits?: number;
      /** The worker replica's yrs state vector after this operation, so the
       * next sync can ship a diff instead of the whole document state. */
      stateVector?: ArrayBuffer;
      /** The region layout the worker ran, for a request carrying `layoutExtras`. */
      layoutJson?: string;
      layoutMeta?: LayoutMetaV1 | { v: number };
      layoutJsonStatus?: 'ok' | 'stale';
      /** `layoutJson` covers only the first pages of the body. */
      layoutProvisional?: boolean;
      /** An `open` reply: the opened package's host metadata JSON. */
      hostJson?: string;
      sourceDocument?: Document;
      /** An `openSource` reply: SVG pictures for the host to decode. */
      svgMedia?: DocxSvgMedia[];
      /** An `openSource` reply: the client id the opened replica writes with. */
      clientId?: number;
      /** An `openSource` reply: the opened state's clock for `clientId`. */
      clientClock?: number;
      /** An `open` reply: the package cannot open as a preview; nothing was opened. */
      previewRefused?: boolean;
      /** A `fontRequirements` reply. */
      requirementsJson?: string;
      /** An `encodeState` reply: the document state as one yrs v1 update. */
      state?: ArrayBuffer;
      peerMetadata?: ArrayBuffer;
      peerMetadataReason?: string;
      /** @internal */
      saved?: ArrayBuffer;
      revisionCount?: number;
      /** @internal */
      proposals?: DocxProposalRegistryState;
      /** @internal */
      version?: string;
      repair?: ArrayBuffer | null;
      /** @internal */
      proposal?: ResidentProposalResponse;
      /** @internal */
      read?: { version: string; value: unknown };
      /** Sent with id 0 while the request with this id is still working. */
      progress?: number;
    }
  | {
      id: number;
      ok: false;
      error: string;
      commandError?: ReplyError;
      failureCode?: SessionFailureCode;
      /** @internal */
      code?: 'save-unavailable' | 'cut-kept';
      residentUnavailable?: boolean;
      /** A wasm trap poisoned the worker; it refuses every later request. */
      terminal?: boolean;
      /** The trap followed an allocation the worker's memory could not satisfy. */
      outOfMemory?: boolean;
    }
) & {
  collabUpdates?: ResidentCollaborationUpdates & { session?: string; incarnation?: number; epoch?: number };
  /** The worker's wasm memories as the reply left. */
  memory?: WasmModuleMemory[];
};
