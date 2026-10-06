/// <reference lib="webworker" />

import { LAYOUT_META_VERSION, type LayoutMetaV1, type RetainedLayoutMeta } from './layoutMeta';
import type { CollaborationCursor } from '../collaboration/types';
import type { YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import { finalPreviewDisplayWindow, finalPreviewPageCount } from './previewDisplayWindow';
import { preloadEditWasm, preloadEditWasmFrom } from './wasm/index';
import {
  createProposalRegistry,
  proposalRevisionPreview,
  type DocxProposalRegistry,
  type DocxProposalRequest,
  type DocxProposalResult,
} from './proposals';
import { computeProposalGeometryMirror, resolveNavigationTarget } from './proposalGeometry';
import { findBodyMatches } from './findMatches';
import { readResidentSearch, residentBodyPositions } from './residentSearch';
import type { ResidentSaveRecord } from './residentSave';
import { proposalProjectionStories } from './dirtyProjectionStories';
import { findParagraphs } from './findParagraphs';
import { DisplayPositionIndex } from './displayPositionIndex';
import { resolveYrsPointPosition } from './pointPosition';
import {
  resolveBookmarkPosition,
  resolveCommentTarget,
  resolveParagraphTarget,
  resolveRevisionTarget,
  resolveSelectionInfo,
  resolveSelectionText,
  resolveSelectionUnit,
} from './viewerSelection';
import { readSidebar, readOutlineHeadings } from './sidebarReads';
import { createResidentScheduler, type SchedulerMessage } from './residentScheduler';
import { hasCachedYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import {
  presentOffscreenPageBackBuffer,
  presentOffscreenPageBackBufferWithCaret,
  rasterizeDisplayPageToBackBuffer,
  releaseOffscreenPageCanvas,
} from '../layout/render/canvasBackend';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  retainedFramePageById,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { GlyphCache } from '../layout/render/glyphCache';
import { wasmModuleMemories } from '../wasm/loadWasmAsset';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import {
  RESIDENT_HOST_MODULE_WAIT_MS,
  type ResidentEngineWorkerHostModule,
  type ResidentEngineWorkerRequest,
  type ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';
import {
  residentCaretDeviceRect,
  residentCaretSnapshotForFrame,
  type ResidentCaretPaintStyle,
} from './residentCaret';

const scope = self as unknown as DedicatedWorkerGlobalScope;
let resolveHostEditModule: (module: WebAssembly.Module | null) => void;
let hostEditModule = nextHostEditModule();
let session: ResidentEngineSession | null = null;
let positionIndex: { session: ResidentEngineSession; index: DisplayPositionIndex } | null = null;

function displayPositionIndex(current: ResidentEngineSession): DisplayPositionIndex {
  if (positionIndex?.session !== current) {
    positionIndex = {
      session: current,
      index: new DisplayPositionIndex({
        ...current.geometryReader,
        selectionText: current.selectionText,
        resolveComment: current.resolveComment,
      }),
    };
  }
  return positionIndex.index;
}
let proposals: DocxProposalRegistry | null = null;
let lastProposalMirrorVersion: string | null = null;
let fontRequirements: {
  version: string;
  layoutInput: string;
  requirementsJson: string;
} | null = null;
/** Font requirements by layout input, for one session at one document version. */
let requirementsCache: {
  owner: ResidentEngineSession;
  version: string;
  byInput: Map<string, string>;
  release: () => void;
} | null = null;
/** The layout input of the host's last font requirements request. */
let requestedRequirements: { owner: ResidentEngineSession; layoutInput: string } | null = null;
const REQUIREMENTS_CACHE_INPUTS = 8;
/** Set while the session holds the document `open` seeded, with the heap limit it used. */
let openedDocument: { heapLimitBytes?: number } | null = null;
let openedSource: { bytes: ArrayBuffer; hostJson: string } | null = null;
let editorSaves: ResidentSaveRecord = { full: false };
class SaveUnavailableError extends Error {}
let openedVersion: string | null | undefined;
/** A change before the whole document's first frame leaves no frame as opened. */
function noteDocumentChange(): void {
  if (openedVersion === undefined) openedVersion = null;
}
// The opened document is a display-only preview that an `open` of the whole package replaces.
let previewing = false;
/** Pages of a cut preview's layout that match the whole document's; null for a whole document. */
let previewFinalPages: number | null = null;
let provisionalFinalPages: number | null = null;
let provisionalDisplayWindow: {
  window: [number, number];
  retainBuiltPages: boolean;
} | null = null;
let unsubscribe: (() => void) | null = null;
let pendingUpdates: Uint8Array[] = [];
let layoutRevision = 0;
let retainedRegions = false;
let retainedLayoutVersion: string | null = null;
let headersFootersPayload: string | undefined;
let headersFootersEpoch = 0;
let sentHeadersFootersEpoch = 0;
// -1 = no fonts applied yet (fresh session); hydrate skips re-registration
// when the snapshot's revision matches what this session already holds.
let fontsRevision = -1;
let retainedFrame: RetainedFrame | null = null;
let glyphCache: GlyphCache | null = null;
const offscreenCanvases = new Map<string, OffscreenCanvas>();
const offscreenBackBuffers = new Map<string, OffscreenCanvas>();
const pendingOffscreenPageIds = new Set<string>();
let activeOffscreenPageIds = new Set<string>();
let offscreenDpr = 1;
let offscreenZoom = 1;
// Present-time caret painting. `caretPaintRect` is what the current frame
// wants painted, `paintedCaretPageId/Key` what is on screen. Caret-composited
// pages keep their back-buffer raster (`intactBackBuffers`) so the line can be
// dropped by re-presenting without raster or engine work; plain presents
// detach their bitmap and stay zero-copy.
let caretStyle: ResidentCaretPaintStyle = { color: '#000', width: 2 };
let caretPaintRect: YrsResidentCaretRect | null = null;
let paintedCaretPageId: string | null = null;
let paintedCaretKey: string | null = null;
let caretStage: OffscreenCanvas | null = null;
const intactBackBuffers = new Set<string>();
let trap: WebAssembly.RuntimeError | null = null;
// A bootstrap laid out only the body's first pages: the request to finish
// it, then the finished layout until `completeLayout` hands it over.
interface LayoutRequest {
  extras: string;
  layoutExtras?: string;
}
let incompleteLayout:
  | (LayoutRequest & { layoutInput: string; workerAuthoritative?: boolean })
  | null = null;
let completedLayout:
  | (LayoutRequest & { layoutJson: string; headersFootersJson: string | undefined })
  | null = null;
// A `completeLayout` measured a few blocks at a time, as operations queued
// behind the requests that arrive meanwhile.
interface SlicedCompletion {
  id: number;
  expectedFrameEpoch: number;
  paintCaret: boolean;
  /** Body blocks per step. */
  blocks: number;
  begun: boolean;
  /** Times a change in between abandoned the pass. */
  restarts: number;
}
let slicedCompletion: SlicedCompletion | null = null;
const COMPLETION_RESTARTS = 3;
const COMPLETION_IDLE_MS = 300;
const ALL_BLOCKS = 2 ** 32 - 1;

interface BackgroundPageBuild {
  request: Extract<ResidentEngineWorkerRequest, { type: 'buildPages' }>;
  owner: ResidentEngineSession;
  frameEpoch: number;
  frames: Uint8Array[];
  offset: number;
  started: number;
  engineMs: number;
  pagesPerSlice: number;
}
let backgroundPageBuild: BackgroundPageBuild | null = null;
const BACKGROUND_SLICE_PAGES = 4;

// The request being handled, and the requests answered with a trap.
let handlingId = 0;
const trappedIds = new Set<number>();

scope.onmessage = (
  event: MessageEvent<ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule>
) => {
  const message = event.data;
  if (message.type === 'editModule') {
    // The queued warm waits for this message.
    resolveHostEditModule(message.module instanceof WebAssembly.Module ? message.module : null);
    return;
  }
  scheduler.submit(classify(message));
};
scope.onmessageerror = () => resolveHostEditModule(null);

function nextHostEditModule(): Promise<WebAssembly.Module | null> {
  return new Promise((resolve) => {
    resolveHostEditModule = resolve;
  });
}

function classify(request: ResidentEngineWorkerRequest): SchedulerMessage {
  const run = () => runRequest(request);
  switch (request.type) {
    case 'applyInput':
    case 'applyDelete':
      return { lane: 'input', userInput: true, holdsIdleTasks: true, mutates: true, run };
    case 'applyUpdate':
    case 'syncUpdate':
      return { lane: 'collab', userInput: true, mutates: true, run };
    case 'proposal':
      if (request.operation.kind === 'snapshot') return { lane: 'interactive', run };
      return {
        lane: 'input', mutates: true,
        ...(request.operation.kind === 'setStates' ? { userInput: true, holdsIdleTasks: true } : {}), run,
      };
    case 'open':
    case 'bootstrap':
    case 'sync':
    case 'destroy':
      return { lane: 'input', mutates: true, run };
    case 'warm':
      return { lane: 'input', run };
    case 'documentRead':
      return {
        lane: 'interactive',
        ...(request.expectVersion !== undefined ? { reorderable: true } : {}), run,
      };
    case 'fontRequirements':
    case 'encodeState':
    case 'revisionCount':
    case 'layoutJson':
    case 'save':
      return { lane: 'interactive', run };
    case 'buildPages':
      return {
        lane: 'interactive', reframes: true, key: 'pages', run,
        ...(request.background ? {
          replaceableBy: 'pages',
          supersede: () => replyDropped({ id: request.id, ok: true, pageBuildSuperseded: true }),
        } : {}),
      };
    case 'buildFrame':
    case 'releasePages':
    case 'completeLayout':
    case 'attachCanvases':
    case 'eraseCaret':
      return { lane: 'interactive', reframes: true, run };
  }
}

async function runRequest(request: ResidentEngineWorkerRequest): Promise<void> {
  try {
    if (trap) throw trap;
    handlingId = request.id;
    await handle(request);
  } catch (error) {
    replyFailure(request.id, error);
  }
}

/** Answers a request that will not run; after a trap, as the trap's failure. */
function replyDropped(response: ResidentEngineWorkerResponse & { ok: true }): void {
  if (trap) replyFailure(response.id, trap);
  else reply(response);
}

function replyFailure(id: number, error: unknown): void {
  // Once trapped, every failure is the trap's, answered once per request.
  const failure = trap ?? error;
  if (failure instanceof WebAssembly.RuntimeError) {
    trapped(id, failure);
    return;
  }
  reply({
    id,
    ok: false,
    error: error instanceof Error ? error.message : String(error),
    ...(failure instanceof SaveUnavailableError ? { code: 'save-unavailable' as const } : {}),
  });
}

/** A trap leaves the module unusable: `id` is answered as terminal, and nothing succeeds after. */
function trapped(id: number, error: WebAssembly.RuntimeError): void {
  if (trappedIds.has(id)) return;
  trappedIds.add(id);
  trap = error;
  const failed = editFailedAllocationBytes();
  reply({
    id,
    ok: false,
    error:
      failed > 0
        ? `Resident engine worker ran out of memory allocating ${failed} bytes: ${error.message}`
        : `Resident engine worker trapped: ${error.message}`,
    terminal: true,
    ...(failed > 0 ? { outOfMemory: true } : {}),
  });
}

async function handle(request: ResidentEngineWorkerRequest): Promise<void> {
  if (session && request.type === 'sync') validateFontsBaseRevision(request.snapshot);
  if (
    request.type !== 'documentRead' && request.type !== 'fontRequirements' &&
    request.type !== 'encodeState' && request.type !== 'revisionCount' &&
    request.type !== 'save' &&
    request.type !== 'warm' &&
    !(request.type === 'proposal' && request.operation.kind === 'snapshot')
  ) supersedeBackgroundPageBuild();
  if (request.type === 'warm') {
    try {
      if (request.hostModule) {
        const timer = setTimeout(() => resolveHostEditModule(null), RESIDENT_HOST_MODULE_WAIT_MS);
        await preloadEditWasmFrom(
          hostEditModule.then((module) => {
            clearTimeout(timer);
            return module;
          })
        );
      } else {
        await preloadEditWasm();
      }
      reply({ id: request.id, ok: true });
    } catch (error) {
      // No session exists yet, so a failed load is retried by the next request.
      if (request.hostModule) hostEditModule = nextHostEditModule();
      reply({
        id: request.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    }
    return;
  }
  if (request.type === 'destroy') {
    destroySession();
    return;
  }
  if (request.type === 'open') {
    // One document per worker, so every queued request addresses the one it was sent for;
    // only a preview gives way, to the whole document.
    if (session && (!previewing || request.previewBlocks !== undefined)) {
      throw new Error('Resident engine worker already holds a document');
    }
    // The preview's memory goes before the whole package seeds; its pages stay painted.
    if (session) destroySession(true);
    openedSource = null;
    editorSaves = { full: false };
    const opening = await createResidentEngineSession(request.heapLimitBytes);
    let hostJson: string | null;
    try {
      if (request.previewBlocks === undefined) opening.setDirectBatches(true);
      hostJson =
        request.previewBlocks === undefined
          ? opening.openDocx(new Uint8Array(request.bytes), request.digest, request.generation)
          : opening.openDocxPreview(
              new Uint8Array(request.bytes),
              request.previewBlocks,
              request.previewParagraphBudget
            );
    } catch (error) {
      if (!(error instanceof WebAssembly.RuntimeError)) opening.destroy();
      throw error;
    }
    if (hostJson === null) {
      opening.destroy();
      reply({ id: request.id, ok: true, previewRefused: true });
      return;
    }
    proposals?.destroy();
    proposals = null;
    session = opening;
    if (request.previewBlocks === undefined) openedVersion = undefined;
    openedDocument = { heapLimitBytes: request.heapLimitBytes };
    previewing = request.previewBlocks !== undefined;
    if (!previewing) {
      openedSource = { bytes: request.bytes, hostJson };
    }
    const stateVector = exactBuffer(session.encodeStateVector());
    reply({ id: request.id, ok: true, hostJson, stateVector }, [stateVector]);
    return;
  }
  if (request.type === 'bootstrap') {
    if (!request.opened) {
      destroySession(request.keepSurfaces === true);
      openedVersion = null;
      // The worker is a genuine yrs peer. Reusing the main replica's client id
      // makes a fast structural input race overlap one client's clock range and
      // corrupt the update; a fresh id lets yrs merge queued/local operations
      // safely while the main replica applies worker updates with local origin.
      session = await createResidentEngineSession(request.heapLimitBytes);
    } else if (!session || !openedDocument) {
      throw new Error('Resident engine worker has no opened document');
    } else if (
      request.heapLimitBytes !== undefined &&
      request.heapLimitBytes !== openedDocument.heapLimitBytes
    ) {
      throw new Error('Resident engine worker opened its document under another heap limit');
    }
    unsubscribe?.();
    unsubscribe = null;
    previewFinalPages = null;
    clearProvisionalFinalPages();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, layoutMeta, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      request.opened !== true,
      request.layoutReply,
      request.headersFootersEpoch
    );
    if (previewFinalPages !== null || provisionalFinalPages !== null) {
      setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    }
    if (provisional) {
      incompleteLayout = {
        layoutInput: request.snapshot.layoutInput,
        workerAuthoritative: request.snapshot.workerAuthoritative,
        extras: request.extras,
        layoutExtras: request.layoutExtras,
      };
    }
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(
      frameExtras(request.extras, request.layoutExtras, layoutMeta ? headersFootersPayload : undefined),
      request.expectedFrameEpoch
    );
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      false,
      request.layoutExtras === undefined ? undefined : (layoutMeta ?? layoutJson ?? undefined),
      provisional
    );
    return;
  }
  if (request.type === 'fontRequirements') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    const requirementsJson = layoutFontRequirements(session, request.layoutInput);
    requestedRequirements = { owner: session, layoutInput: request.layoutInput };
    fontRequirements = {
      version: session.proposalEngine.version(),
      layoutInput: request.layoutInput,
      requirementsJson,
    };
    reply({ id: request.id, ok: true, requirementsJson });
    return;
  }
  if (request.type === 'encodeState') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    const state = exactBuffer(session.encodeState());
    reply(
      {
        id: request.id,
        ok: true,
        state,
        version: session.proposalEngine.version(),
        proposals: proposals?.exportState() ?? { previewVersion: 0, entries: [] },
      },
      [state]
    );
    return;
  }
  if (request.type === 'save') {
    if (previewing) throw new SaveUnavailableError('Resident engine worker is still opening');
    if (!session || !openedSource) {
      throw new SaveUnavailableError('Resident engine worker has no opened document to save');
    }
    const held = pendingUpdates;
    pendingUpdates = [];
    try {
      const saved = await session.save(
        new Uint8Array(openedSource.bytes),
        openedSource.hostJson,
        request.host,
        request.comments,
        editorSaves,
        request.stories
      );
      const bytes = saved.slice(0);
      // What the editor copy lacks: the paragraph IDs this save recorded and repairs its updates caused.
      const updates = request.stateVector
        ? [exactBuffer(session.encodeStateAsUpdate(request.stateVector))]
        : [];
      const stateVector = exactBuffer(session.encodeStateVector());
      reply(
        {
          id: request.id,
          ok: true,
          saved: bytes,
          updates,
          stateVector,
          version: session.proposalEngine.version(),
        },
        [bytes, ...updates, stateVector]
      );
    } finally {
      pendingUpdates = held;
    }
    return;
  }
  if (request.type === 'revisionCount') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    // Host proposals' revisions are not the document's own.
    const proposed = new Set(proposals?.snapshot().proposals.flatMap((p) => p.revisionIds));
    reply({ id: request.id, ok: true, revisionCount: session.revisionCount(proposed) });
    return;
  }
  if (request.type === 'eraseCaret') {
    caretPaintRect = null;
    if (paintedCaretPageId !== null) await replayOffscreen(false);
    reply({ id: request.id, ok: true });
    return;
  }
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (request.type === 'proposal') {
    if (!unsubscribe) throw new Error('Resident engine worker has not laid out its document');
    pendingUpdates = [];
    let committed = false;
    try {
      const registry = proposals ??= createProposalRegistry(session.proposalEngine);
      const previousVersion = session.proposalEngine.version();
      const known = new Set(registry.snapshot().proposals.map((proposal) => proposal.id));
      const since = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      let result: DocxProposalResult | undefined;
      let projectionStories: string[] = [];
      switch (request.operation.kind) {
        case 'propose':
          result = registry.propose(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'setStates':
          result = registry.setStates(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'withdraw':
          result = registry.withdraw(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'removeComment':
          try {
            const story = session.selection()?.head.story ?? 'body';
            session.applyRawOps(story, [{ op: 'removeComment', id: request.operation.id }]);
            projectionStories = [story];
          } catch {}
          break;
      }
      committed = request.operation.kind !== 'snapshot';
      const changedStories = session.storiesChangedSince(since).stories;
      if (result?.ok) {
        projectionStories = [...proposalProjectionStories(known, result, changedStories)];
      }
      session.markProjectionStories(projectionStories);
      if (changedStories.length > 0) completedLayout = null;
      const updates = pendingUpdates.map(exactBuffer);
      const peerDiff = 'peerStateVector' in request.operation && request.operation.peerStateVector
        ? exactBuffer(session.encodeStateAsUpdate(request.operation.peerStateVector))
        : undefined;
      const stateVector = exactBuffer(session.encodeStateVector());
      const snapshot = registry.snapshot();
      const version = snapshot.version;
      const geometry = computeProposalGeometryMirror(
        session.geometryReader,
        snapshot,
        version === previousVersion && (
          version === lastProposalMirrorVersion || hasCachedYrsSidebarProjection(session.geometryReader)
        )
      );
      lastProposalMirrorVersion = geometry.version;
      const unchangedFonts =
        fontRequirements?.version === previousVersion &&
        request.operation.kind === 'propose' &&
        result?.ok === true &&
        proposalRevisionPreview(snapshot) === undefined &&
        asciiProposalFontsUnchanged(session, request.operation.request, result);
      fontRequirements =
        fontRequirements && (unchangedFonts ||
          (request.operation.kind === 'snapshot' && fontRequirements.version === snapshot.version))
          ? { ...fontRequirements, version: snapshot.version }
          : null;
      // A decision changes only the preview: the host's next input is its last one previewing it.
      const preview = proposalRevisionPreview(snapshot);
      if (
        request.operation.kind === 'setStates' &&
        result?.ok === true &&
        changedStories.length === 0 &&
        requestedRequirements?.owner === session
      ) {
        fontRequirements = previewFontRequirements(
          session,
          requestedRequirements.layoutInput,
          preview,
          snapshot.version
        );
      }
      reply(
        {
          id: request.id,
          ok: true,
          proposal: {
            ...(result === undefined ? {} : { result }),
            mirror: { version, proposals: registry.exportState() },
            changedStories,
            projectionStories,
            updates,
            ...(peerDiff === undefined ? {} : { peerDiff }),
            stateVector,
            geometry,
            ...(fontRequirements
              ? { fontRequirements: {
                  layoutInput: fontRequirements.layoutInput,
                  requirementsJson: fontRequirements.requirementsJson,
                } }
              : {}),
          },
        },
        [...updates, stateVector, ...(peerDiff === undefined ? [] : [peerDiff])]
      );
    } catch (error) {
      if (trap) throw trap;
      if (error instanceof WebAssembly.RuntimeError) throw error;
      reply({
        id: request.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
        ...(committed || pendingUpdates.length > 0 ? { terminal: true } : {}),
      });
    } finally {
      pendingUpdates = [];
    }
    return;
  }
  if (request.type === 'documentRead') {
    const engine = session.proposalEngine;
    if (request.expectVersion !== undefined && engine.version() !== request.expectVersion) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    let value: unknown;
    switch (request.read.kind) {
      case 'exportStructuredWithPages':
        value = session.exportStructuredWithPagesJson(request.read.options, request.read.currentRequest);
        break;
      case 'listContentControls':
        value = session.listContentControls(request.read.options);
        break;
      case 'findContentControls':
        value = session.findContentControls(request.read.query, request.read.options);
        break;
      case 'paragraphIdentities':
        value = session.paragraphIdentities();
        break;
      case 'resolveParagraphAnchors':
        value = {
          results: request.read.anchors.map((anchor) => engine.resolveParagraphAnchor(anchor)),
        };
        break;
      case 'readParagraphs':
        value = engine.readParagraphs(request.read.request);
        break;
      case 'findText':
        value = engine.findText(request.read.request);
        break;
      case 'findMatches':
        value = engine.version() !== request.read.expectVersion ? null : findBodyMatches(
          session.geometryReader,
          residentBodyPositions(session.geometryReader),
          request.read.searchText,
          request.read.options
        );
        break;
      case 'navigationTarget':
        value = resolveNavigationTarget(
          session.geometryReader,
          request.read.story,
          request.read.paraId
        );
        break;
      case 'searchText':
        value = readResidentSearch(
          {
            ...session.geometryReader,
            searchText: session.searchText,
            resolveStickyPosition: session.resolveStickyPosition,
          },
          request.read.query,
          request.read.caseSensitive,
          request.read.carry
        );
        break;
      case 'pointPosition':
        value = resolveYrsPointPosition(
          displayPositionIndex(session),
          request.read.hit,
          request.read.expectVersion
        );
        break;
      case 'findParagraphs':
        value = findParagraphs(session.geometryReader, request.read.query, request.read);
        break;
      case 'selectionUnit':
        value = resolveSelectionUnit(
          displayPositionIndex(session),
          request.read.story,
          request.read.position,
          request.read.unit,
          request.read.expectVersion
        );
        break;
      case 'selectionText':
        value = resolveSelectionText(
          displayPositionIndex(session),
          request.read.story,
          request.read.anchor,
          request.read.head,
          request.read.expectVersion
        );
        break;
      case 'bookmarkPosition':
        value = resolveBookmarkPosition(
          displayPositionIndex(session),
          request.read.story,
          request.read.name,
          request.read.expectVersion
        );
        break;
      case 'selectionInfo':
        value = resolveSelectionInfo(
          displayPositionIndex(session),
          request.read.story,
          request.read.anchor,
          request.read.head,
          request.read.expectVersion
        );
        break;
      case 'paragraphTarget':
        value = resolveParagraphTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.paraId,
          request.read.expectVersion
        );
        break;
      case 'commentTarget':
        value = resolveCommentTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.commentId,
          request.read.expectVersion
        );
        break;
      case 'revisionTarget':
        value = resolveRevisionTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.revisionId,
          request.read.expectVersion
        );
        break;
      case 'sidebar':
        value = readSidebar(session.geometryReader, request.read.commentIds, request.read.expectVersion);
        break;
      case 'headings':
        value = readOutlineHeadings(session.geometryReader, request.read.expectVersion);
        break;
      case 'stickyAnchors': {
        const currentSession = session;
        value = request.read.locs.map((loc) => {
          try {
            return currentSession.encodeStickyPosition(loc);
          } catch {
            return null;
          }
        });
        break;
      }
    }
    reply({ id: request.id, ok: true, read: { version: engine.version(), value } });
    return;
  }
  if (request.type === 'layoutJson') {
    if (!retainedRegions || request.layoutRevision !== layoutRevision ||
      retainedLayoutVersion !== session.proposalEngine.version()) {
      reply({ id: request.id, ok: true, layoutJsonStatus: 'stale' });
    } else {
      reply({ id: request.id, ok: true, layoutJsonStatus: 'ok',
        layoutRevision, layoutJson: session.retainedLayoutJson() });
    }
    return;
  }
  if (request.type === 'sync') {
    unsubscribe?.();
    unsubscribe = null;
    previewFinalPages = null;
    clearProvisionalFinalPages();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, layoutMeta, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      true,
      request.layoutReply,
      request.headersFootersEpoch
    );
    if (previewFinalPages !== null || provisionalFinalPages !== null) {
      setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    }
    if (provisional) {
      incompleteLayout = {
        layoutInput: request.snapshot.layoutInput,
        workerAuthoritative: request.snapshot.workerAuthoritative,
        extras: request.extras,
        layoutExtras: request.layoutExtras,
      };
    }
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(
      frameExtras(request.extras, request.layoutExtras, layoutMeta ? headersFootersPayload : undefined),
      request.expectedFrameEpoch
    );
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret,
      request.layoutExtras === undefined ? undefined : (layoutMeta ?? layoutJson ?? undefined),
      provisional
    );
    return;
  }
  if (request.type === 'buildPages') {
    const limit = previewFinalPages === null
      ? provisionalFinalPages
      : Math.min(previewFinalPages, provisionalFinalPages ?? previewFinalPages);
    const pages = limit === null ? request.pages : request.pages.filter((index) => index < limit);
    if (request.background && pages.length > BACKGROUND_SLICE_PAGES) {
      const build: BackgroundPageBuild = {
        request: pages === request.pages ? request : { ...request, pages },
        owner: session, frameEpoch: request.expectedFrameEpoch,
        frames: [], offset: 0, started: performance.now(), engineMs: 0,
        pagesPerSlice: BACKGROUND_SLICE_PAGES,
      };
      backgroundPageBuild = build;
      scheduler.schedule({
        kind: 'pageBuild', version: scheduler.version, generation: scheduler.generation,
        onStale: () => 'cancel',
        cancel: () => {
          if (backgroundPageBuild === build) supersedeBackgroundPageBuild();
        },
        fail: (error) => {
          if (backgroundPageBuild === build) backgroundPageBuild = null;
          replyFailure(build.request.id, error);
        },
        run: async (budgetMs) => {
          if (backgroundPageBuild !== build) return 'done';
          if (trap) throw trap;
          handlingId = build.request.id;
          return (await backgroundPageSlice(build, budgetMs)) ? 'done' : 'yield';
        },
      });
      return;
    }
    setFrameDisplayWindow(session);
    // Pages of the provisional frame build between steps, as before a completion.
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayPagesFrame(pages, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'releasePages') {
    if (
      !retainedFrame ||
      retainedFrame.frameEpoch !== request.expectedFrameEpoch ||
      request.pages.some(
        ({ index, pageId }) => retainedFrame!.pages[index]?.pageId.toString() !== pageId
      )
    ) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    const started = performance.now();
    const frame = session.releaseDisplayPagesFrame(
      request.pages.map(({ index }) => index),
      request.expectedFrameEpoch
    );
    if (frame === null) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      [],
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'completeLayout') {
    setFrameDisplayWindow(session);
    if (incompleteLayout && request.sliceBlocks) {
      supersedeSlicedCompletion();
      slicedCompletion = {
        id: request.id,
        expectedFrameEpoch: request.expectedFrameEpoch,
        paintCaret: request.paintCaret,
        blocks: request.sliceBlocks,
        begun: false,
        restarts: 0,
      };
      scheduleCompletion(slicedCompletion);
      return;
    }
    await completeProvisionalLayout();
    await replyCompletedLayout(request.id, request.expectedFrameEpoch, request.paintCaret);
    return;
  }
  if (request.type === 'buildFrame') {
    if (!incompleteLayout?.workerAuthoritative) await completeProvisionalLayout();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'applyUpdate' || request.type === 'syncUpdate') {
    noteDocumentChange();
    fontRequirements = null;
    if (request.update.length > 0 || request.type === 'applyUpdate') session.applyUpdate(request.update);
    if (request.type === 'applyUpdate') {
      if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
    } else {
      const repair = exactBuffer(session.encodeStateAsUpdate(request.stateVector));
      const stateVector = exactBuffer(session.encodeStateVector());
      reply({
        id: request.id,
        ok: true,
        version: session.proposalEngine.version(),
        stateVector,
        repair,
      }, [stateVector, repair]);
    }
    return;
  }
  if (request.type === 'attachCanvases') {
    const environmentChanged =
      offscreenCanvases.size > 0 &&
      (offscreenDpr !== request.devicePixelRatio || offscreenZoom !== request.zoom);
    // Pages needing pixels: freshly transferred canvases plus retained
    // canvases re-entering the active window (their buffers were zeroed when
    // they left it).
    const forcedPageIds = new Set(request.pages.map((page) => page.pageId));
    for (const pageId of request.activePageIds) {
      if (!activeOffscreenPageIds.has(pageId)) forcedPageIds.add(pageId);
    }
    for (const { pageId, canvas } of request.pages) offscreenCanvases.set(pageId, canvas);
    activeOffscreenPageIds = new Set(request.activePageIds);
    offscreenDpr = request.devicePixelRatio;
    offscreenZoom = request.zoom;
    caretStyle = request.caretStyle;
    for (const [pageId, canvas] of offscreenCanvases) {
      if (!activeOffscreenPageIds.has(pageId)) {
        // out of the page window: release the bitmap but KEEP the canvas —
        // a transferred surface can never be re-transferred, so the element
        // must stay usable for re-entry
        releaseOffscreenPageCanvas(canvas);
        offscreenBackBuffers.delete(pageId);
        forgetOffscreenPagePixels(pageId);
      }
    }
    // A dpr/zoom change repaints every active page; otherwise only the forced
    // set needs pixels — surviving active pages keep their bitmaps.
    await replayOffscreen(environmentChanged ? true : forcedPageIds);
    reply({ id: request.id, ok: true });
    return;
  }
  await completeProvisionalLayout();
  noteDocumentChange();
  fontRequirements = null;
  // The edit replaces the pagination a cached completion's frame would paint.
  completedLayout = null;
  setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
  session.setSelection(request.selection.anchor, request.selection.head);
  pendingUpdates = [];
  const started = performance.now();
  try {
    const applied =
      request.type === 'applyDelete'
        ? request.profile
          ? session.applyDeleteProfiled(
              request.direction,
              request.expectedFrameEpoch,
              request.count
            )
          : {
              frame: session.applyDelete(
                request.direction,
                request.expectedFrameEpoch,
                request.count
              ),
              profile: undefined,
            }
        : request.profile
          ? session.applyInputProfiled(request.text, request.expectedFrameEpoch)
          : {
              frame: session.applyInput(request.text, request.expectedFrameEpoch),
              profile: undefined,
            };
    // Typing inside a surrogate pair lands before it and settles the caret after the text; the
    // host's sticky caret stays inside the pair, so it takes this one.
    let settled: CollaborationCursor | null = null;
    if (request.type === 'applyInput') {
      const sent = request.selection.head;
      const head = session.selection()?.head;
      const expected = sent.offset + request.text.length;
      if (head && (head.paraId !== sent.paraId || head.offset !== expected)) {
        settled = session.encodeSelection();
      }
    }
    await replyFrame(
      request.id,
      applied.frame,
      performance.now() - started,
      pendingUpdates,
      applied.profile,
      started,
      request.selection.head.story === 'body',
      request.paintCaret,
      undefined,
      false,
      request.type === 'applyDelete' ? session.residentDeletedUnits() : undefined,
      [],
      settled
    );
  } catch (error) {
    if (trap) throw trap;
    if (error instanceof WebAssembly.RuntimeError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    reply({
      id: request.id,
      ok: false,
      error: message,
      residentUnavailable: message.includes('resident input state is not ready'),
      // The edit committed here but never reached the host: this replica is
      // no longer the host's, so the host must replace it.
      ...(pendingUpdates.length > 0 ? { terminal: true } : {}),
    });
  } finally {
    pendingUpdates = [];
  }
}

function validateFontsBaseRevision(snapshot: YrsResidentWorkerSnapshot): void {
  if (snapshot.fontsBaseRevision !== undefined && snapshot.fontsBaseRevision !== fontsRevision) {
    throw new Error('Resident engine worker font base revision mismatch');
  }
}

/**
 * Loads a snapshot and runs its layout, over the first `provisionalPages`
 * pages only when given; returns the region layout reply, which a full pass
 * serializes only when `reply` asks for it.
 */
function hydrate(
  snapshot: YrsResidentWorkerSnapshot,
  provisionalPages?: number,
  reply = true,
  loadState = true,
  replyMode?: 'meta',
  knownHeadersFootersEpoch?: number
): { layoutJson: string | null; layoutMeta?: LayoutMetaV1; provisional: boolean } {
  if (!session) throw new Error('Resident engine worker is not initialized');
  validateFontsBaseRevision(snapshot);
  supersedeSlicedCompletion();
  incompleteLayout = null;
  clearProvisionalFinalPages();
  completedLayout = null;
  if (loadState && !snapshot.workerAuthoritative) {
    noteDocumentChange();
    fontRequirements = null;
    session.loadState(snapshot.state);
  }
  session.setPartialDocument(snapshot.partialDocument === true);
  previewFinalPages = snapshot.partialDocument === true ? 0 : null;
  if (!snapshot.workerAuthoritative) {
    session.loadMediaSources(snapshot.mediaSources ?? '');
    session.loadNoteSeparators(snapshot.noteSeparators ?? new Uint8Array(0));
  }
  if (snapshot.fontsRevision !== fontsRevision) {
    if (snapshot.fontsBaseRevision === undefined) {
      session.clearFonts();
      glyphCache = null;
      for (const pageId of activeOffscreenPageIds) pendingOffscreenPageIds.add(pageId);
      intactBackBuffers.clear();
    }
    fontsRevision = -1;
    for (const font of snapshot.fonts) {
      if (font instanceof Uint8Array) session.registerFont(font);
      else session.registerSubstituteFont(font.substituteOf, font.family);
    }
    fontsRevision = snapshot.fontsRevision;
  }
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  let layoutJson: string | null = null;
  let metadata: RetainedLayoutMeta | undefined;
  let provisional = false;
  let pageCount: number | null = null;
  if (snapshot.layoutWithRegions && provisionalPages !== undefined) {
    layoutJson = session.layoutDocumentWithRegionsPrefixRetainedJson(
      snapshot.layoutInput,
      provisionalPages
    );
    const layout = JSON.parse(layoutJson) as {
      provisional?: boolean;
      layout: { pages: unknown[] };
    };
    provisional = layout.provisional === true;
    pageCount = layout.layout.pages.length;
  } else if (snapshot.layoutWithRegions && !reply) {
    session.layoutDocumentWithRegionsRetained(snapshot.layoutInput);
  } else if (snapshot.layoutWithRegions && replyMode === 'meta') {
    metadata = session.layoutDocumentWithRegionsRetainedMeta(snapshot.layoutInput);
    provisional = metadata.provisional;
    pageCount = metadata.pageCount;
  } else if (snapshot.layoutWithRegions) {
    // the retained reply leaves out the tens-of-MB measured arena
    layoutJson = session.layoutDocumentWithRegionsRetainedJson(snapshot.layoutInput);
    if (snapshot.partialDocument === true) {
      pageCount = (JSON.parse(layoutJson) as { layout: { pages: unknown[] } }).layout.pages.length;
    }
  } else {
    session.layoutDocumentJson(snapshot.layoutInput);
  }
  if (!snapshot.workerAuthoritative && snapshot.selection) {
    session.setSelection(snapshot.selection.anchor, snapshot.selection.head);
  }
  layoutRevision = snapshot.layoutRevision;
  retainedRegions = snapshot.layoutWithRegions;
  retainedLayoutVersion = session.proposalEngine.version();
  pendingUpdates = [];
  previewFinalPages = finalPreviewPageCount(
    snapshot.partialDocument,
    provisional,
    provisionalPages,
    pageCount
  );
  provisionalFinalPages = provisional && snapshot.partialDocument !== true
    ? provisionalPages ?? null
    : null;
  return {
    layoutJson,
    ...(metadata ? { layoutMeta: layoutMetaReply(metadata, knownHeadersFootersEpoch) } : {}),
    provisional,
  };
}

function retainedHeadersFootersPayload(): string {
  const payload = session?.retainedHeadersFootersJson() ?? 'null';
  if (payload !== headersFootersPayload) {
    headersFootersPayload = payload;
    headersFootersEpoch += 1;
  }
  return payload;
}

function layoutMetaReply(meta: RetainedLayoutMeta, knownEpoch = sentHeadersFootersEpoch): LayoutMetaV1 {
  const payload = retainedHeadersFootersPayload();
  const changed = knownEpoch !== headersFootersEpoch;
  return {
    v: LAYOUT_META_VERSION,
    layoutRevision,
    ...meta,
    headersFootersEpoch,
    ...(changed ? { headersFooters: payload } : {}),
  };
}

function forgetRequirementsCache(): void {
  requirementsCache?.release();
  requirementsCache = null;
}

function layoutFontRequirements(engine: ResidentEngineSession, layoutInput: string): string {
  const version = engine.proposalEngine.version();
  if (requirementsCache?.owner !== engine || requirementsCache.version !== version) {
    forgetRequirementsCache();
    requirementsCache = {
      owner: engine,
      version,
      byInput: new Map(),
      release: engine.onUpdate(forgetRequirementsCache),
    };
  }
  const cached = requirementsCache.byInput.get(layoutInput);
  if (cached !== undefined) return cached;
  const requirementsJson = engine.layoutFontRequirementsJson(layoutInput);
  if (requirementsCache.byInput.size >= REQUIREMENTS_CACHE_INPUTS) {
    requirementsCache.byInput.delete(requirementsCache.byInput.keys().next().value as string);
  }
  requirementsCache.byInput.set(layoutInput, requirementsJson);
  return requirementsJson;
}

/**
 * The font requirements of `layoutInput` previewing `preview`, keyed by the input the host builds
 * for it (its render environment with `revisionPreview` replaced), or null when they cannot be read.
 */
function previewFontRequirements(
  engine: ResidentEngineSession,
  layoutInput: string,
  preview: ReturnType<typeof proposalRevisionPreview>,
  version: string
): NonNullable<typeof fontRequirements> | null {
  try {
    if (engine.proposalEngine.version() !== version) return null;
    const request = JSON.parse(layoutInput) as { renderEnv?: Record<string, unknown> | null };
    if (!request.renderEnv || typeof request.renderEnv !== 'object') return null;
    if (preview === undefined) delete request.renderEnv.revisionPreview;
    else request.renderEnv.revisionPreview = preview;
    const next = JSON.stringify(request);
    return { version, layoutInput: next, requirementsJson: layoutFontRequirements(engine, next) };
  } catch (error) {
    // The host reads them itself and meets the failure there.
    if (trap) throw trap;
    if (error instanceof WebAssembly.RuntimeError) throw error;
    return null;
  }
}

function setFrameDisplayWindow(
  engine: ResidentEngineSession,
  window?: [number, number],
  retainBuiltPages?: boolean
): void {
  if (previewFinalPages !== null) {
    engine.setDisplayWindow(...finalPreviewDisplayWindow(window, previewFinalPages));
    engine.setDisplayRetainBuiltPages(window !== undefined && retainBuiltPages === true);
    engine.setWindowedIncrementalBuilds(true);
    return;
  }
  if (window) {
    if (provisionalFinalPages !== null) {
      provisionalDisplayWindow = { window, retainBuiltPages: retainBuiltPages === true };
      engine.setDisplayWindow(...finalPreviewDisplayWindow(window, provisionalFinalPages));
      engine.setDisplayRetainBuiltPages(false);
    } else {
      engine.setDisplayWindow(...window);
      engine.setDisplayRetainBuiltPages(retainBuiltPages === true);
    }
  }
  engine.setWindowedIncrementalBuilds(window !== undefined);
}

function clearProvisionalFinalPages(): void {
  provisionalFinalPages = null;
  if (session && provisionalDisplayWindow) {
    session.setDisplayWindow(...provisionalDisplayWindow.window);
    session.setDisplayRetainBuiltPages(provisionalDisplayWindow.retainBuiltPages);
  }
  provisionalDisplayWindow = null;
}

/**
 * Replaces a provisional layout with the full one before anything reads it,
 * finishing a sliced completion at once and answering its request first.
 */
async function completeProvisionalLayout(): Promise<void> {
  if (!session || !incompleteLayout) return;
  const waiting = slicedCompletion;
  slicedCompletion = null;
  const { layoutInput, ...request } = incompleteLayout;
  incompleteLayout = null;
  try {
    let layoutJson: string | undefined;
    if (waiting?.begun) {
      try {
        layoutJson = session.resumeRegionLayout(ALL_BLOCKS).layoutJson;
      } catch (error) {
        if (error instanceof WebAssembly.RuntimeError) throw error;
      }
    }
    completedLayout = {
      ...request,
      layoutJson: layoutJson ?? session.layoutDocumentWithRegionsRetainedJson(layoutInput),
      headersFootersJson: session.retainedHeadersFootersJson(),
    };
    clearProvisionalFinalPages();
    if (waiting) {
      await replyCompletedLayout(waiting.id, waiting.expectedFrameEpoch, waiting.paintCaret);
    }
  } catch (error) {
    // The waiting completion is answered too, and the request that finished it fails.
    if (waiting) replyFailure(waiting.id, error);
    throw error;
  }
}

/** Answers a `completeLayout` with the completed layout and its frame, or with none. */
async function replyCompletedLayout(
  id: number,
  expectedFrameEpoch: number,
  paintCaret: boolean
): Promise<void> {
  const completed = completedLayout;
  completedLayout = null;
  if (!session || !completed) {
    reply({ id, ok: true });
    return;
  }
  setFrameDisplayWindow(session);
  pendingUpdates = [];
  const started = performance.now();
  const frame = session.buildDisplayListFrame(
    frameExtras(completed.extras, completed.layoutExtras, completed.headersFootersJson),
    expectedFrameEpoch
  );
  await replyFrame(
    id,
    frame,
    performance.now() - started,
    pendingUpdates,
    undefined,
    started,
    false,
    paintCaret,
    completed.layoutJson
  );
}

// A zero-delay turn of the event loop, so requests that arrived meanwhile
// queue ahead of the next step.
const completionTurns = typeof MessageChannel === 'function' ? new MessageChannel() : null;
const turnCallbacks: Array<() => void> = [];
if (completionTurns) completionTurns.port1.onmessage = () => turnCallbacks.shift()?.();

function nextTurn(callback: () => void): void {
  if (!completionTurns) {
    setTimeout(callback, 0);
    return;
  }
  turnCallbacks.push(callback);
  completionTurns.port2.postMessage(null);
}

const scheduler = createResidentScheduler({
  now: () => performance.now(),
  turn: (callback) => nextTurn(callback),
  timer: (callback, ms) => {
    const handle = setTimeout(callback, ms);
    return () => clearTimeout(handle);
  },
  failed: () => {},
});

function supersedeBackgroundPageBuild(): void {
  const build = backgroundPageBuild;
  if (!build) return;
  backgroundPageBuild = null;
  replyDropped({ id: build.request.id, ok: true, pageBuildSuperseded: true });
}

async function backgroundPageSlice(build: BackgroundPageBuild, budgetMs: number): Promise<boolean> {
  if (session !== build.owner || (build.offset > 0 && retainedFrame?.frameEpoch !== build.frameEpoch)) {
    supersedeBackgroundPageBuild();
    return true;
  }
  setFrameDisplayWindow(build.owner);
  const started = performance.now();
  const pages = build.request.pages.slice(build.offset, build.offset + build.pagesPerSlice);
  const bytes = build.owner.buildDisplayPagesFrame(pages, build.frameEpoch);
  const elapsed = performance.now() - started;
  build.engineMs += elapsed;
  build.pagesPerSlice = Math.min(
    32,
    Math.max(BACKGROUND_SLICE_PAGES, Math.round((pages.length * budgetMs) / Math.max(1, elapsed)))
  );
  build.offset += pages.length;
  if (build.offset === build.request.pages.length) {
    backgroundPageBuild = null;
    await replyFrame(
      build.request.id, bytes, build.engineMs, [], undefined, build.started,
      false, build.request.paintCaret, undefined, false, undefined, build.frames
    );
    return true;
  }
  applyWorkerFrame(bytes);
  build.frameEpoch = retainedFrame!.frameEpoch;
  build.frames.push(bytes);
  return false;
}

function scheduleCompletion(completion: SlicedCompletion): void {
  scheduler.schedule({
    kind: 'completion', version: scheduler.version, generation: scheduler.generation,
    idleAfterInputMs: COMPLETION_IDLE_MS,
    onStale: () => 'continue',
    fail: (error) => {
      if (slicedCompletion === completion) slicedCompletion = null;
      replyFailure(completion.id, error);
    },
    run: async (budgetMs) => {
      if (slicedCompletion !== completion) return 'done';
      if (trap) throw trap;
      handlingId = completion.id;
      await completionSlice(completion, budgetMs);
      return slicedCompletion === completion ? 'yield' : 'done';
    },
  });
}

async function completionSlice(completion: SlicedCompletion, budgetMs: number): Promise<void> {
  if (!session || slicedCompletion !== completion || !incompleteLayout) return;
  let progress;
  if (!completion.begun) {
    progress = session.beginRegionLayout(incompleteLayout.layoutInput);
    completion.begun = true;
  } else {
    const started = performance.now();
    try {
      progress = session.resumeRegionLayout(completion.blocks);
    } catch (error) {
      if (error instanceof WebAssembly.RuntimeError) throw error;
      const message = error instanceof Error ? error.message : String(error);
      if (
        message !== 'no region layout to resume' &&
        message !== 'the document or its fonts changed since the region layout began'
      ) throw error;
      // A change in between abandoned the pass: begin again on the new state. Host
      // proposals the worker holds keep yielding to user requests; other changes
      // finish in one step once they keep coming.
      completion.begun = false;
      completion.restarts += 1;
      if (
        completion.restarts > COMPLETION_RESTARTS &&
        !incompleteLayout.workerAuthoritative &&
        !proposals
      ) {
        try {
          await completeProvisionalLayout();
        } catch {
          // Answered as the completion's failure.
        }
        return;
      }
      return;
    }
    const elapsed = Math.max(1, performance.now() - started);
    completion.blocks = Math.min(
      8192,
      Math.max(8, Math.round((completion.blocks * budgetMs) / elapsed))
    );
  }
  if (progress.layoutJson === undefined) {
    return;
  }
  const { layoutInput: _input, ...request } = incompleteLayout;
  incompleteLayout = null;
  clearProvisionalFinalPages();
  slicedCompletion = null;
  completedLayout = {
    ...request,
    layoutJson: progress.layoutJson,
    headersFootersJson: session.retainedHeadersFootersJson(),
  };
  await replyCompletedLayout(completion.id, completion.expectedFrameEpoch, completion.paintCaret);
}

/** A snapshot or a new session replaces the layout a sliced completion was finishing. */
function supersedeSlicedCompletion(): void {
  const completion = slicedCompletion;
  slicedCompletion = null;
  if (completion) replyDropped({ id: completion.id, ok: true });
}

/**
 * The extras a frame is built with. For a layout this worker owns, the host
 * sends them without the header/footer payload, which only this layout has:
 * the session retains it after a region layout,
 * or a completed layout captured it.
 */
function frameExtras(
  extras: string,
  layoutExtras: string | undefined,
  headersFootersJson?: string
): string {
  if (layoutExtras === undefined) return extras;
  const retained =
    headersFootersJson ??
    (retainedRegions ? retainedHeadersFootersPayload() : undefined);
  const headersFooters =
    retained === undefined
      ? undefined
      : (JSON.parse(retained) as DisplayListBuildInputs['headersFooters']);
  return encodeDisplayListFrameExtras({
    ...(JSON.parse(layoutExtras) as DisplayListBuildInputs),
    ...(headersFooters ? { headersFooters } : {}),
  });
}

function subscribe(): void {
  if (!session) return;
  unsubscribe = session.onUpdate((update) => pendingUpdates.push(update.slice()));
}

function asciiProposalFontsUnchanged(
  engine: ResidentEngineSession,
  request: DocxProposalRequest,
  result: DocxProposalResult
): boolean {
  if (!result.ok || !fontRequirements) return false;
  const requirements = JSON.parse(fontRequirements.requirementsJson) as { scripts?: string[] }[];
  const input = JSON.parse(fontRequirements.layoutInput) as {
    renderEnv?: { revisionPreview?: Record<string, unknown> };
  };
  if (requirements.some(({ scripts }) => scripts && scripts.length > 0) ||
    Object.keys(input.renderEnv?.revisionPreview ?? {}).length > 0) return false;
  const stories = new Map<string, Set<string>>();
  for (const proposal of request.proposals) {
    const text = proposal.op === 'replaceText' ? proposal.replaceWith : proposal.text;
    if (!/^[\x20-\x7e]*$/.test(text)) return false;
    const record = result.snapshot.proposals.find(({ id }) => id === proposal.id);
    if (!record?.changed) continue;
    const { story, paraId } = record.paragraph;
    const paragraphs = stories.get(story) ?? new Set<string>();
    paragraphs.add(paraId);
    stories.set(story, paragraphs);
  }
  for (const [story, targets] of stories) {
    const found = new Map<string, number>();
    const spans = engine.geometryReader.paragraphSpans(story);
    for (let index = 0; index < spans.length; index += 1) {
      const { paraId } = spans[index]!;
      if (!targets.has(paraId)) continue;
      if (found.has(paraId)) return false;
      found.set(paraId, index);
    }
    if (found.size !== targets.size) return false;
    const indices = [...found.values()];
    for (const unit of engine.paragraphSegments(story, indices)) {
      let ascii = true;
      let existingText = false;
      for (const segment of unit) {
        if (segment.kind === 'text') {
          ascii &&= /^[\x20-\x7e]*$/.test(segment.text);
          existingText ||= segment.text.length > 0 && segment.attributes.ins == null;
        } else if (segment.kind !== 'pilcrow') {
          ascii = false;
        }
      }
      if (!ascii || !existingText) return false;
    }
  }
  return true;
}

/**
 * Drops the document. `keepSurfaces` keeps the attached page canvases, still
 * showing the old pages, for a document that replaces it page for page.
 */
function destroySession(keepSurfaces = false): void {
  supersedeBackgroundPageBuild();
  fontRequirements = null;
  forgetRequirementsCache();
  requestedRequirements = null;
  unsubscribe?.();
  unsubscribe = null;
  proposals?.destroy();
  proposals = null;
  session?.destroy();
  lastProposalMirrorVersion = null;
  session = null;
  positionIndex = null;
  openedDocument = null;
  openedSource = null;
  editorSaves = { full: false };
  openedVersion = undefined;
  previewing = false;
  previewFinalPages = null;
  clearProvisionalFinalPages();
  pendingUpdates = [];
  layoutRevision = 0;
  retainedRegions = false;
  retainedLayoutVersion = null;
  sentHeadersFootersEpoch = 0;
  fontsRevision = -1;
  supersedeSlicedCompletion();
  incompleteLayout = null;
  completedLayout = null;
  retainedFrame = null;
  glyphCache = null;
  offscreenBackBuffers.clear();
  pendingOffscreenPageIds.clear();
  if (!keepSurfaces) {
    offscreenCanvases.clear();
    activeOffscreenPageIds.clear();
  }
  caretPaintRect = null;
  paintedCaretPageId = null;
  paintedCaretKey = null;
  caretStage = null;
  intactBackBuffers.clear();
}

function forgetOffscreenPagePixels(pageId: string): void {
  intactBackBuffers.delete(pageId);
  if (paintedCaretPageId === pageId) {
    paintedCaretPageId = null;
    paintedCaretKey = null;
  }
}

function retainedPageByKey(frame: RetainedFrame, pageId: string) {
  return /^\d+$/.test(pageId) ? retainedFramePageById(frame, BigInt(pageId)) : undefined;
}

function applyWorkerFrame(bytes: Uint8Array): void {
  retainedFrame = applyFrameDeltaOwned(retainedFrame, decodeFrameDelta(bytes));
  for (const pageId of retainedFrame.damagedPageIds) pendingOffscreenPageIds.add(pageId.toString());
}

async function replyFrame(
  id: number,
  bytes: Uint8Array,
  engineMs: number,
  updates = pendingUpdates,
  engineProfile?: import('./index').YrsEngineApplyProfile,
  requestStarted = performance.now(),
  requireCaret = false,
  paintCaret = false,
  layoutReply?: string | LayoutMetaV1,
  layoutProvisional = false,
  deletedUnits?: number,
  precedingPageFrames: Uint8Array[] = [],
  selectionCursor: CollaborationCursor | null = null
): Promise<void> {
  const layoutJson = typeof layoutReply === 'string' ? layoutReply : undefined;
  const layoutMeta = typeof layoutReply === 'object' ? layoutReply : undefined;
  const documentVersion = session?.proposalEngine.version();
  if (layoutReply !== undefined) retainedLayoutVersion = documentVersion ?? null;
  const documentPreview = previewing;
  if (!previewing && openedVersion === undefined && documentVersion !== undefined) {
    openedVersion = documentVersion;
  }
  const documentAsOpened =
    !previewing && documentVersion !== undefined && documentVersion === openedVersion;
  applyWorkerFrame(bytes);
  const limit = provisionalFinalPages;
  if (session && limit !== null && retainedFrame) {
    const pages = retainedFrame.displayList.pages
      .filter((page) => page.pageIndex >= limit && !page.unbuilt)
      .map((page) => page.pageIndex);
    if (pages.length > 0) {
      const released = session.releaseDisplayPagesFrame(pages, retainedFrame.frameEpoch);
      if (released === null) throw new Error('Provisional display pages could not be released');
      bytes = session.buildDisplayPagesFrame([], 0);
      applyWorkerFrame(bytes);
    }
  }
  const caret = session?.residentCaretSnapshot();
  if (!caret || !retainedFrame || !residentCaretSnapshotForFrame(caret, retainedFrame)) {
    throw new Error('Resident caret snapshot does not match the produced frame');
  }
  if (requireCaret && !caret.caretRect) {
    throw new Error('Resident input frame omitted collapsed caret geometry');
  }
  const selection = session?.selection() ?? null;
  caretPaintRect = paintCaret ? (caret.caretRect ?? null) : null;
  // Pages no longer in the document drop their surfaces (their elements
  // unmounted main-side). An unbuilt page keeps its transferred canvas, which
  // can never be transferred again, and only loses its pixels.
  for (const pageId of pendingOffscreenPageIds) {
    const page = retainedPageByKey(retainedFrame, pageId);
    if (!page || page.page.unbuilt === true) pendingOffscreenPageIds.delete(pageId);
  }
  for (const pageId of new Set([...offscreenCanvases.keys(), ...offscreenBackBuffers.keys()])) {
    const page = retainedPageByKey(retainedFrame, pageId);
    const unbuilt = page ? page.page.unbuilt === true : undefined;
    if (unbuilt === false) continue;
    const canvas = offscreenCanvases.get(pageId);
    if (unbuilt === undefined) offscreenCanvases.delete(pageId);
    else if (canvas) releaseOffscreenPageCanvas(canvas);
    offscreenBackBuffers.delete(pageId);
    forgetOffscreenPagePixels(pageId);
  }
  const replayStarted = performance.now();
  const { replayedPages, caretPainted } = await replayOffscreen(false);
  const replayMs = performance.now() - replayStarted;
  const frame = exactBuffer(bytes);
  const pageFrames = precedingPageFrames.map(exactBuffer);
  const updateBuffers = updates.map(exactBuffer);
  const stateVector = session ? exactBuffer(session.encodeStateVector()) : undefined;
  if (layoutMeta) sentHeadersFootersEpoch = layoutMeta.headersFootersEpoch;
  reply(
    {
      id,
      ok: true,
      frame,
      ...(pageFrames.length > 0 ? { pageFrames: [...pageFrames, frame] } : {}),
      updates: updateBuffers,
      engineMs,
      workerTotalMs: performance.now() - requestStarted,
      engineProfile,
      caret,
      selection,
      ...(selectionCursor ? { selectionCursor } : {}),
      caretPainted,
      replayMs,
      replayedPages,
      layoutRevision,
      ...(documentVersion === undefined ? {} : { documentVersion }),
      ...(documentPreview ? { documentPreview: true } : {}),
      ...(documentAsOpened ? { documentAsOpened: true } : {}),
      ...(deletedUnits === undefined ? {} : { deletedUnits }),
      ...(stateVector ? { stateVector } : {}),
      ...(layoutJson !== undefined ? { layoutJson } : {}),
      ...(layoutMeta ? { layoutMeta } : {}),
      ...(layoutProvisional ? { layoutProvisional } : {}),
    },
    [frame, ...pageFrames, ...updateBuffers, ...(stateVector ? [stateVector] : []),
      ...(layoutMeta ? [layoutMeta.pageSizes.buffer as ArrayBuffer] : [])]
  );
}

async function replayOffscreen(
  force: boolean | Set<string>
): Promise<{ replayedPages: number; caretPainted: boolean }> {
  const forcedPageIds = force === true ? activeOffscreenPageIds : force;
  if (forcedPageIds) {
    for (const { pageId, page } of retainedFrame?.pages ?? []) {
      const key = pageId.toString();
      if (!page.unbuilt && forcedPageIds.has(key)) pendingOffscreenPageIds.add(key);
    }
  }
  if (!retainedFrame || offscreenCanvases.size === 0) {
    return { replayedPages: 0, caretPainted: false };
  }
  if (!glyphCache && session) {
    glyphCache = new GlyphCache({
      provider: (fontId, glyphId) => {
        try {
          return session!.outlineGlyphJson(fontId, glyphId);
        } catch (error) {
          // The raster paints on with browser text, so running out of memory is answered here.
          if (error instanceof WebAssembly.RuntimeError && editFailedAllocationBytes() > 0) {
            trapped(handlingId, error);
          }
          throw error;
        }
      },
    });
  }
  const caretTarget =
    caretPaintRect && activeOffscreenPageIds.has(caretPaintRect.pageId) ? caretPaintRect : null;
  const caretDevice = caretTarget
    ? residentCaretDeviceRect(caretTarget, caretStyle, offscreenDpr, offscreenZoom)
    : null;
  const caretKey =
    caretTarget && caretDevice
      ? `${caretTarget.pageId}|${caretDevice.x}|${caretDevice.y}|${caretDevice.width}|${caretDevice.height}|${caretStyle.color}`
      : null;
  const preparations: Array<
    Promise<{
      canvas: OffscreenCanvas;
      buffer: OffscreenCanvas;
      pageId: string;
    }>
  > = [];
  for (let index = 0; index < retainedFrame.pages.length; index += 1) {
    const retainedPage = retainedFrame.pages[index];
    const pageIdString = retainedPage.pageId.toString();
    // Off-window pages hold no pixels; they re-raster through the forced set
    // when they re-enter the window.
    if (!activeOffscreenPageIds.has(pageIdString)) continue;
    const damaged = pendingOffscreenPageIds.has(pageIdString);
    // Beyond damage, a page presents only for caret compositing: the page
    // gaining the painted line and the page losing it.
    const gainsCaret =
      pageIdString === caretTarget?.pageId &&
      !(paintedCaretPageId === pageIdString && paintedCaretKey === caretKey);
    const losesCaret =
      pageIdString === paintedCaretPageId && pageIdString !== caretTarget?.pageId;
    if (!damaged && !gainsCaret && !losesCaret) continue;
    const canvas = offscreenCanvases.get(pageIdString);
    const page = retainedFrame.displayList.pages[index];
    if (!canvas || !page || page.unbuilt) continue;
    const pageId = pageIdString;
    let buffer = offscreenBackBuffers.get(pageId);
    if (!buffer) {
      buffer = new OffscreenCanvas(1, 1);
      offscreenBackBuffers.set(pageId, buffer);
    }
    const resolvedBuffer = buffer;
    if (!damaged && intactBackBuffers.has(pageId)) {
      preparations.push(Promise.resolve({ canvas, buffer: resolvedBuffer, pageId }));
      continue;
    }
    preparations.push(
      rasterizeDisplayPageToBackBuffer(
        resolvedBuffer,
        page,
        { glyphCache: glyphCache ?? undefined },
        offscreenDpr,
        offscreenZoom
      ).then(() => ({ canvas, buffer: resolvedBuffer, pageId }))
    );
  }
  const prepared = await Promise.all(preparations).catch(async (error) => {
    await Promise.allSettled(preparations);
    throw error;
  });
  let caretPainted =
    caretTarget !== null && paintedCaretPageId === caretTarget.pageId && paintedCaretKey === caretKey;
  // Present only after the entire damaged frame is ready. This loop is
  // synchronous, so the compositor can never observe the renderer's clears.
  for (const { canvas, buffer, pageId } of prepared) {
    if (caretTarget && caretDevice && pageId === caretTarget.pageId) {
      if (!caretStage) caretStage = new OffscreenCanvas(1, 1);
      presentOffscreenPageBackBufferWithCaret(canvas, buffer, caretStage, {
        ...caretDevice,
        color: caretStyle.color,
      });
      intactBackBuffers.add(pageId);
      paintedCaretPageId = pageId;
      paintedCaretKey = caretKey;
      caretPainted = true;
    } else {
      presentOffscreenPageBackBuffer(canvas, buffer);
      intactBackBuffers.delete(pageId);
      if (paintedCaretPageId === pageId) {
        paintedCaretPageId = null;
        paintedCaretKey = null;
      }
    }
    pendingOffscreenPageIds.delete(pageId);
  }
  return { replayedPages: prepared.length, caretPainted };
}

function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer as ArrayBuffer;
  }
  return bytes.slice().buffer;
}

function editFailedAllocationBytes(): number {
  return wasmModuleMemories().find((module) => module.label === 'docx-edit')?.failedAllocationBytes ?? 0;
}

function reply(response: ResidentEngineWorkerResponse, transfer: Transferable[] = []): void {
  // A trap the raster painted past fails the request that would succeed, and
  // every request waiting on it, through the paths that answer failures.
  if (trap && response.ok) throw trap;
  scope.postMessage({ ...response, memory: wasmModuleMemories() }, transfer);
}
