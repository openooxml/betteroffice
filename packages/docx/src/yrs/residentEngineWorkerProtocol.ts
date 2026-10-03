import type {
  YrsEngineApplyProfile,
  YrsLoc,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
  YrsStickyPosition,
} from './index';
import type { CollaborationCursor } from '../collaboration/types';
import type { ResidentSearchResult } from './residentSearch';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
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
import type { DocxReadParagraphsRequest, DocxReadParagraphsResult } from './edits';
import type { ProposalGeometryMirror, resolveNavigationTarget } from './proposalGeometry';

/** @internal */
export type ResidentProposalOperation =
  | { kind: 'propose'; request: DocxProposalRequest }
  | { kind: 'setStates'; request: DocxProposalStateRequest }
  | { kind: 'withdraw'; request: DocxProposalWithdrawRequest }
  | { kind: 'snapshot' };

/** @internal */
export interface ResidentProposalResponse {
  result?: DocxProposalResult;
  mirror: { version: string; proposals: DocxProposalRegistryState };
  changedStories: string[];
  updates: ArrayBuffer[];
  stateVector: ArrayBuffer;
  geometry: ProposalGeometryMirror;
  fontRequirements?: { layoutInput: string; requirementsJson: string };
}

/** @internal */
export type ResidentDocumentRead =
  | { kind: 'paragraphIdentities' }
  | { kind: 'resolveParagraphAnchors'; anchors: DocxParagraphAnchor[] }
  | { kind: 'readParagraphs'; request: DocxReadParagraphsRequest }
  | { kind: 'searchText'; query: string; caseSensitive: boolean; carry?: YrsStickyPosition | null }
  | { kind: 'stickyAnchors'; locs: YrsLoc[]; version: string }
  | { kind: 'navigationTarget'; story: string; paraId: string };

/** @internal */
export interface ResidentDocumentReadValues {
  paragraphIdentities: DocxParagraphIdentitySnapshot;
  resolveParagraphAnchors: { results: DocxParagraphAnchorResult[] };
  readParagraphs: DocxReadParagraphsResult;
  navigationTarget: ReturnType<typeof resolveNavigationTarget>;
  searchText: ResidentSearchResult;
  stickyAnchors: Array<YrsStickyPosition | null>;
}

/** How long a warm waits for the host's compiled module before loading the engine itself. */
export const RESIDENT_HOST_MODULE_WAIT_MS = 10_000;

/** @internal */
export type ResidentEngineWorkerHostModule = {
  type: 'editModule';
  module: WebAssembly.Module | null;
};

export type ResidentEngineWorkerRequest =
  | { id: number; type: 'warm'; hostModule?: true }
  | {
      id: number;
      type: 'bootstrap';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      layoutExtras?: string;
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
    }
  | { id: number; type: 'fontRequirements'; layoutInput: string }
  | { id: number; type: 'encodeState' }
  | { id: number; type: 'revisionCount' }
  | { id: number; type: 'proposal'; operation: ResidentProposalOperation }
  | {
      id: number;
      type: 'documentRead';
      read: ResidentDocumentRead;
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

export type ResidentEngineWorkerRequestWithoutId = ResidentEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

export type ResidentEngineWorkerResponse = (
  | {
      id: number;
      ok: true;
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
      /** Characters an applyDelete removed. */
      deletedUnits?: number;
      /** The worker replica's yrs state vector after this operation, so the
       * next sync can ship a diff instead of the whole document state. */
      stateVector?: ArrayBuffer;
      /** The region layout the worker ran, for a request carrying `layoutExtras`. */
      layoutJson?: string;
      /** `layoutJson` covers only the first pages of the body. */
      layoutProvisional?: boolean;
      /** An `open` reply: the opened package's host metadata JSON. */
      hostJson?: string;
      /** An `open` reply: the package cannot open as a preview; nothing was opened. */
      previewRefused?: boolean;
      /** A `fontRequirements` reply. */
      requirementsJson?: string;
      /** An `encodeState` reply: the document state as one yrs v1 update. */
      state?: ArrayBuffer;
      revisionCount?: number;
      /** @internal */
      proposals?: DocxProposalRegistryState;
      /** @internal */
      version?: string;
      /** @internal */
      proposal?: ResidentProposalResponse;
      /** @internal */
      read?: { version: string; value: unknown };
    }
  | {
      id: number;
      ok: false;
      error: string;
      residentUnavailable?: boolean;
      /** A wasm trap poisoned the worker; it refuses every later request. */
      terminal?: boolean;
      /** The trap followed an allocation the worker's memory could not satisfy. */
      outOfMemory?: boolean;
    }
) & {
  /** The worker's wasm memories as the reply left. */
  memory?: WasmModuleMemory[];
};
