/// <reference lib="webworker" />

import type { YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import { finalPreviewDisplayWindow, finalPreviewPageCount } from './previewDisplayWindow';
import { preloadEditWasm } from './wasm/index';
import {
  createProposalRegistry,
  proposalRevisionPreview,
  type DocxProposalRegistry,
  type DocxProposalRequest,
  type DocxProposalResult,
} from './proposals';
import { computeProposalGeometryMirror, resolveNavigationTarget } from './proposalGeometry';
import { readResidentSearch } from './residentSearch';
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
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { GlyphCache } from '../layout/render/glyphCache';
import { wasmModuleMemories } from '../wasm/loadWasmAsset';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';
import {
  residentCaretDeviceRect,
  residentCaretSnapshotForFrame,
  type ResidentCaretPaintStyle,
} from './residentCaret';

const scope = self as unknown as DedicatedWorkerGlobalScope;
let session: ResidentEngineSession | null = null;
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
// The opened document is a display-only preview that an `open` of the whole package replaces.
let previewing = false;
/** Pages of a cut preview's layout that match the whole document's; null for a whole document. */
let previewFinalPages: number | null = null;
let unsubscribe: (() => void) | null = null;
let pendingUpdates: Uint8Array[] = [];
let layoutRevision = 0;
// -1 = no fonts applied yet (fresh session); hydrate skips re-registration
// when the snapshot's revision matches what this session already holds.
let fontsRevision = -1;
let operations = Promise.resolve();
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
  /** Body blocks a step measures, tuned toward `COMPLETION_SLICE_MS`. */
  blocks: number;
  begun: boolean;
  /** Times a change in between abandoned the pass. */
  restarts: number;
}
let slicedCompletion: SlicedCompletion | null = null;
const COMPLETION_SLICE_MS = 24;
const COMPLETION_RESTARTS = 3;
const ALL_BLOCKS = 2 ** 32 - 1;

interface BackgroundPageBuild {
  request: Extract<ResidentEngineWorkerRequest, { type: 'buildPages' }>;
  owner: ResidentEngineSession;
  frameEpoch: number;
  frames: Uint8Array[];
  offset: number;
  started: number;
  engineMs: number;
}
let backgroundPageBuild: BackgroundPageBuild | null = null;
const BACKGROUND_SLICE_PAGES = 4;

// The request being handled, and the requests answered with a trap.
let handlingId = 0;
const trappedIds = new Set<number>();

scope.onmessage = (event: MessageEvent<ResidentEngineWorkerRequest>) => {
  enqueue(() => handle(event.data), event.data.id);
};

/** `current` drops an operation whose request was answered while it waited. */
function enqueue(
  operation: () => Promise<void> | void,
  id: number,
  current: () => boolean = () => true
): void {
  operations = operations
    .then(() => {
      if (!current()) return;
      if (trap) throw trap;
      handlingId = id;
      return operation();
    })
    .catch((error) => replyFailure(id, error));
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
  supersedeBackgroundPageBuild();
  if (request.type === 'warm') {
    try {
      await preloadEditWasm();
      reply({ id: request.id, ok: true });
    } catch (error) {
      // No session exists yet, so a failed load is retried by the next request.
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
    const opening = await createResidentEngineSession(request.heapLimitBytes);
    let hostJson: string | null;
    try {
      hostJson =
        request.previewBlocks === undefined
          ? opening.openDocx(new Uint8Array(request.bytes), request.digest, request.generation)
          : opening.openDocxPreview(new Uint8Array(request.bytes), request.previewBlocks);
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
    openedDocument = { heapLimitBytes: request.heapLimitBytes };
    previewing = request.previewBlocks !== undefined;
    const stateVector = exactBuffer(session.encodeStateVector());
    reply({ id: request.id, ok: true, hostJson, stateVector }, [stateVector]);
    return;
  }
  if (request.type === 'bootstrap') {
    if (!request.opened) {
      destroySession(request.keepSurfaces === true);
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
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      request.opened !== true
    );
    if (previewFinalPages !== null) {
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
      frameExtras(request.extras, request.layoutExtras, layoutJson),
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
      request.layoutExtras === undefined ? undefined : (layoutJson ?? undefined),
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
      const since = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      let result: DocxProposalResult | undefined;
      switch (request.operation.kind) {
        case 'propose':
          result = registry.propose(request.operation.request);
          break;
        case 'setStates':
          result = registry.setStates(request.operation.request);
          break;
        case 'withdraw':
          result = registry.withdraw(request.operation.request);
          break;
      }
      committed = request.operation.kind !== 'snapshot';
      const changedStories = session.storiesChangedSince(since).stories;
      if (changedStories.length > 0) completedLayout = null;
      const updates = pendingUpdates.map(exactBuffer);
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
            updates,
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
        [...updates, stateVector]
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
    let value: unknown;
    switch (request.read.kind) {
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
  if (request.type === 'sync') {
    unsubscribe?.();
    unsubscribe = null;
    previewFinalPages = null;
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined
    );
    if (previewFinalPages !== null) {
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
      frameExtras(request.extras, request.layoutExtras, layoutJson),
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
      request.layoutExtras === undefined ? undefined : (layoutJson ?? undefined),
      provisional
    );
    return;
  }
  if (request.type === 'buildPages') {
    const limit = previewFinalPages;
    const pages = limit === null ? request.pages : request.pages.filter((index) => index < limit);
    if (request.background && pages.length > BACKGROUND_SLICE_PAGES) {
      const build: BackgroundPageBuild = {
        request: pages === request.pages ? request : { ...request, pages },
        owner: session, frameEpoch: request.expectedFrameEpoch,
        frames: [], offset: 0, started: performance.now(), engineMs: 0,
      };
      backgroundPageBuild = build;
      scheduleBackgroundPageSlice(build);
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
      scheduleCompletionSlice(slicedCompletion);
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
  if (request.type === 'applyUpdate') {
    fontRequirements = null;
    session.applyUpdate(request.update);
    if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
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
      request.type === 'applyDelete' ? session.residentDeletedUnits() : undefined
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

/**
 * Loads a snapshot and runs its layout, over the first `provisionalPages`
 * pages only when given; returns the region layout reply, which a full pass
 * serializes only when `reply` asks for it.
 */
function hydrate(
  snapshot: YrsResidentWorkerSnapshot,
  provisionalPages?: number,
  reply = true,
  loadState = true
): { layoutJson: string | null; provisional: boolean } {
  if (!session) throw new Error('Resident engine worker is not initialized');
  supersedeSlicedCompletion();
  incompleteLayout = null;
  completedLayout = null;
  if (loadState && !snapshot.workerAuthoritative) {
    fontRequirements = null;
    session.loadState(snapshot.state);
  }
  session.setPartialDocument(snapshot.partialDocument === true);
  previewFinalPages = snapshot.partialDocument === true ? 0 : null;
  if (!snapshot.workerAuthoritative) session.loadMediaSources(snapshot.mediaSources ?? '');
  if (snapshot.fontsRevision !== fontsRevision) {
    // A mismatched revision always carries the full font set (the client only
    // omits fonts when it knows this session's applied revision matches).
    session.clearFonts();
    for (const font of snapshot.fonts) {
      if (font instanceof Uint8Array) session.registerFont(font);
      else session.registerSubstituteFont(font.substituteOf, font.family);
    }
    fontsRevision = snapshot.fontsRevision;
  }
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  let layoutJson: string | null = null;
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
  pendingUpdates = [];
  previewFinalPages = finalPreviewPageCount(
    snapshot.partialDocument,
    provisional,
    provisionalPages,
    pageCount
  );
  return { layoutJson, provisional };
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
    engine.setDisplayWindow(...window);
    engine.setDisplayRetainBuiltPages(retainBuiltPages === true);
  }
  engine.setWindowedIncrementalBuilds(window !== undefined);
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
    frameExtras(completed.extras, completed.layoutExtras, null, completed.headersFootersJson),
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

function supersedeBackgroundPageBuild(): void {
  const build = backgroundPageBuild;
  if (!build) return;
  backgroundPageBuild = null;
  reply({ id: build.request.id, ok: true, pageBuildSuperseded: true });
}

function scheduleBackgroundPageSlice(build: BackgroundPageBuild): void {
  nextTurn(() => {
    if (backgroundPageBuild !== build) return;
    enqueue(async () => {
      try {
        await backgroundPageSlice(build);
      } catch (error) {
        if (backgroundPageBuild === build) backgroundPageBuild = null;
        throw error;
      }
    }, build.request.id, () => backgroundPageBuild === build);
  });
}

async function backgroundPageSlice(build: BackgroundPageBuild): Promise<void> {
  if (session !== build.owner || (build.offset > 0 && retainedFrame?.frameEpoch !== build.frameEpoch)) {
    supersedeBackgroundPageBuild();
    return;
  }
  setFrameDisplayWindow(build.owner);
  const started = performance.now();
  const pages = build.request.pages.slice(build.offset, build.offset + BACKGROUND_SLICE_PAGES);
  const bytes = build.owner.buildDisplayPagesFrame(pages, build.frameEpoch);
  build.engineMs += performance.now() - started;
  build.offset += pages.length;
  if (build.offset === build.request.pages.length) {
    backgroundPageBuild = null;
    await replyFrame(
      build.request.id, bytes, build.engineMs, [], undefined, build.started,
      false, build.request.paintCaret, undefined, false, undefined, build.frames
    );
    return;
  }
  applyWorkerFrame(bytes);
  build.frameEpoch = retainedFrame!.frameEpoch;
  build.frames.push(bytes);
  scheduleBackgroundPageSlice(build);
}

/** Queues the next step of `completion`, which a later completion supersedes. */
function scheduleCompletionSlice(completion: SlicedCompletion): void {
  nextTurn(() => {
    if (slicedCompletion !== completion) return;
    enqueue(
      async () => {
        try {
          await completionSlice(completion);
        } catch (error) {
          if (slicedCompletion === completion) slicedCompletion = null;
          throw error;
        }
      },
      completion.id,
      () => slicedCompletion === completion
    );
  });
}

/** One bounded step of a sliced completion, which then queues the next. */
async function completionSlice(completion: SlicedCompletion): Promise<void> {
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
      scheduleCompletionSlice(completion);
      return;
    }
    const elapsed = Math.max(1, performance.now() - started);
    completion.blocks = Math.min(
      8192,
      Math.max(8, Math.round((completion.blocks * COMPLETION_SLICE_MS) / elapsed))
    );
  }
  if (progress.layoutJson === undefined) {
    scheduleCompletionSlice(completion);
    return;
  }
  const { layoutInput: _input, ...request } = incompleteLayout;
  incompleteLayout = null;
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
  if (completion) reply({ id: completion.id, ok: true });
}

/**
 * The extras a frame is built with. For a layout this worker owns, the host
 * sends them without the header/footer payload, which only this layout has:
 * the session retains it after a region layout (`layoutJson` is its reply),
 * or a completed layout captured it.
 */
function frameExtras(
  extras: string,
  layoutExtras: string | undefined,
  layoutJson: string | null,
  headersFootersJson?: string
): string {
  if (layoutExtras === undefined) return extras;
  const retained =
    headersFootersJson ??
    (layoutJson === null ? undefined : session?.retainedHeadersFootersJson());
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
    let ascii = true;
    let existingText = false;
    const found = new Set<string>();
    for (const segment of engine.geometryReader.storySegments(story)) {
      if (segment.kind === 'pilcrow') {
        if (targets.has(segment.paraId)) {
          if (!ascii || !existingText || found.has(segment.paraId)) return false;
          found.add(segment.paraId);
        }
        ascii = true;
        existingText = false;
      } else if (segment.kind === 'text') {
        ascii &&= /^[\x20-\x7e]*$/.test(segment.text);
        existingText ||= segment.text.length > 0 && segment.attributes.ins == null;
      } else {
        ascii = false;
      }
    }
    if (found.size !== targets.size) return false;
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
  openedDocument = null;
  previewing = false;
  previewFinalPages = null;
  pendingUpdates = [];
  layoutRevision = 0;
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
  layoutJson?: string,
  layoutProvisional = false,
  deletedUnits?: number,
  precedingPageFrames: Uint8Array[] = []
): Promise<void> {
  applyWorkerFrame(bytes);
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
  const unbuiltByPageId = new Map(
    retainedFrame.pages.map(({ pageId, page }) => [pageId.toString(), page.unbuilt === true])
  );
  for (const pageId of pendingOffscreenPageIds) {
    if (unbuiltByPageId.get(pageId) !== false) pendingOffscreenPageIds.delete(pageId);
  }
  for (const pageId of new Set([...offscreenCanvases.keys(), ...offscreenBackBuffers.keys()])) {
    const unbuilt = unbuiltByPageId.get(pageId);
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
      caretPainted,
      replayMs,
      replayedPages,
      layoutRevision,
      ...(deletedUnits === undefined ? {} : { deletedUnits }),
      ...(stateVector ? { stateVector } : {}),
      ...(layoutJson !== undefined ? { layoutJson } : {}),
      ...(layoutProvisional ? { layoutProvisional } : {}),
    },
    [frame, ...pageFrames, ...updateBuffers, ...(stateVector ? [stateVector] : [])]
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
