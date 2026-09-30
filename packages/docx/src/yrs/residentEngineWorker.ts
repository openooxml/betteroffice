/// <reference lib="webworker" />

import type { YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import { preloadEditWasm } from './wasm/index';
import {
  presentOffscreenPageBackBuffer,
  presentOffscreenPageBackBufferWithCaret,
  rasterizeDisplayPageToBackBuffer,
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
/** Set while the session holds the document `open` seeded, with the heap limit it used. */
let openedDocument: { heapLimitBytes?: number } | null = null;
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
let incompleteLayout: (LayoutRequest & { layoutInput: string }) | null = null;
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
    // One document per worker, so every queued request addresses the one it was sent for.
    if (session) {
      throw new Error('Resident engine worker already holds a document');
    }
    const opening = await createResidentEngineSession(request.heapLimitBytes);
    let hostJson: string;
    try {
      hostJson = opening.openDocx(new Uint8Array(request.bytes), request.digest, request.generation);
    } catch (error) {
      if (!(error instanceof WebAssembly.RuntimeError)) opening.destroy();
      throw error;
    }
    session = opening;
    openedDocument = { heapLimitBytes: request.heapLimitBytes };
    const stateVector = exactBuffer(session.encodeStateVector());
    reply({ id: request.id, ok: true, hostJson, stateVector }, [stateVector]);
    return;
  }
  if (request.type === 'bootstrap') {
    if (!request.opened) {
      destroySession();
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
    if (request.displayWindow) session.setDisplayWindow(...request.displayWindow);
    const { layoutJson, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      request.opened !== true
    );
    if (provisional) {
      incompleteLayout = {
        layoutInput: request.snapshot.layoutInput,
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
    reply({
      id: request.id,
      ok: true,
      requirementsJson: session.layoutFontRequirementsJson(request.layoutInput),
    });
    return;
  }
  if (request.type === 'encodeState') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    const state = exactBuffer(session.encodeState());
    reply({ id: request.id, ok: true, state }, [state]);
    return;
  }
  if (request.type === 'eraseCaret') {
    caretPaintRect = null;
    if (paintedCaretPageId !== null) await replayOffscreen(false);
    reply({ id: request.id, ok: true });
    return;
  }
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (request.type === 'sync') {
    unsubscribe?.();
    unsubscribe = null;
    if (request.displayWindow) session.setDisplayWindow(...request.displayWindow);
    const { layoutJson } = hydrate(request.snapshot, undefined, request.layoutExtras !== undefined);
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
      request.layoutExtras === undefined ? undefined : (layoutJson ?? undefined)
    );
    return;
  }
  if (request.type === 'buildPages') {
    // Pages of the provisional frame build between steps, as before a completion.
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayPagesFrame(request.pages, request.expectedFrameEpoch);
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
  if (request.type === 'completeLayout') {
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
    await completeProvisionalLayout();
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
        canvas.width = 0;
        canvas.height = 0;
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
  // The edit replaces the pagination a cached completion's frame would paint.
  completedLayout = null;
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
  if (loadState) session.loadState(snapshot.state);
  session.setPartialDocument(snapshot.partialDocument === true);
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
  if (snapshot.layoutWithRegions && provisionalPages !== undefined) {
    layoutJson = session.layoutDocumentWithRegionsPrefixRetainedJson(
      snapshot.layoutInput,
      provisionalPages
    );
    provisional = (JSON.parse(layoutJson) as { provisional?: boolean }).provisional === true;
  } else if (snapshot.layoutWithRegions && !reply) {
    session.layoutDocumentWithRegionsRetained(snapshot.layoutInput);
  } else if (snapshot.layoutWithRegions) {
    // the retained reply leaves out the tens-of-MB measured arena
    layoutJson = session.layoutDocumentWithRegionsRetainedJson(snapshot.layoutInput);
  } else {
    session.layoutDocumentJson(snapshot.layoutInput);
  }
  if (snapshot.selection) session.setSelection(snapshot.selection.anchor, snapshot.selection.head);
  layoutRevision = snapshot.layoutRevision;
  pendingUpdates = [];
  return { layoutJson, provisional };
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
      // A change in between abandoned the pass: begin again on the new state,
      // or finish in one step once changes keep coming.
      completion.begun = false;
      completion.restarts += 1;
      if (completion.restarts > COMPLETION_RESTARTS) {
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

function destroySession(): void {
  unsubscribe?.();
  unsubscribe = null;
  session?.destroy();
  session = null;
  openedDocument = null;
  pendingUpdates = [];
  layoutRevision = 0;
  fontsRevision = -1;
  supersedeSlicedCompletion();
  incompleteLayout = null;
  completedLayout = null;
  retainedFrame = null;
  glyphCache = null;
  offscreenCanvases.clear();
  offscreenBackBuffers.clear();
  pendingOffscreenPageIds.clear();
  activeOffscreenPageIds.clear();
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
  deletedUnits?: number
): Promise<void> {
  retainedFrame = applyFrameDeltaOwned(retainedFrame, decodeFrameDelta(bytes));
  for (const pageId of retainedFrame.damagedPageIds) pendingOffscreenPageIds.add(pageId.toString());
  const caret = session?.residentCaretSnapshot();
  if (!caret || !residentCaretSnapshotForFrame(caret, retainedFrame)) {
    throw new Error('Resident caret snapshot does not match the produced frame');
  }
  if (requireCaret && !caret.caretRect) {
    throw new Error('Resident input frame omitted collapsed caret geometry');
  }
  const selection = session?.selection() ?? null;
  caretPaintRect = paintCaret ? (caret.caretRect ?? null) : null;
  // Pages no longer in the document release their surfaces entirely (their
  // elements unmounted main-side); off-window pages are only zeroed, so this
  // is the sole place a live document's canvas reference is dropped.
  const livePageIds = new Set(retainedFrame.pages.map((page) => page.pageId.toString()));
  for (const pageId of pendingOffscreenPageIds) {
    if (!livePageIds.has(pageId)) pendingOffscreenPageIds.delete(pageId);
  }
  for (const pageId of offscreenCanvases.keys()) {
    if (!livePageIds.has(pageId)) {
      offscreenCanvases.delete(pageId);
      offscreenBackBuffers.delete(pageId);
      forgetOffscreenPagePixels(pageId);
    }
  }
  const replayStarted = performance.now();
  const { replayedPages, caretPainted } = await replayOffscreen(false);
  const replayMs = performance.now() - replayStarted;
  const frame = exactBuffer(bytes);
  const updateBuffers = updates.map(exactBuffer);
  const stateVector = session ? exactBuffer(session.encodeStateVector()) : undefined;
  reply(
    {
      id,
      ok: true,
      frame,
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
    [frame, ...updateBuffers, ...(stateVector ? [stateVector] : [])]
  );
}

async function replayOffscreen(
  force: boolean | Set<string>
): Promise<{ replayedPages: number; caretPainted: boolean }> {
  const forcedPageIds = force === true ? activeOffscreenPageIds : force;
  if (forcedPageIds) {
    for (const pageId of forcedPageIds) pendingOffscreenPageIds.add(pageId);
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
    if (!canvas || !page) continue;
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
