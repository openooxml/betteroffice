/// <reference lib="webworker" />

import type { YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
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

scope.onmessage = (event: MessageEvent<ResidentEngineWorkerRequest>) => {
  operations = operations
    .then(() => handle(event.data))
    .catch((error) => {
      reply({
        id: event.data.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    });
};

async function handle(request: ResidentEngineWorkerRequest): Promise<void> {
  if (request.type === 'destroy') {
    destroySession();
    return;
  }
  if (request.type === 'bootstrap') {
    destroySession();
    // The worker is a genuine yrs peer. Reusing the main replica's client id
    // makes a fast structural input race overlap one client's clock range and
    // corrupt the update; a fresh id lets yrs merge queued/local operations
    // safely while the main replica applies worker updates with local origin.
    session = await createResidentEngineSession();
    hydrate(request.snapshot);
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started
    );
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
    hydrate(request.snapshot);
    subscribe();
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
  if (request.type === 'buildFrame') {
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
  session.setSelection(request.selection.anchor, request.selection.head);
  pendingUpdates = [];
  const started = performance.now();
  try {
    const applied =
      request.type === 'applyDelete'
        ? request.profile
          ? session.applyDeleteProfiled(request.direction, request.expectedFrameEpoch)
          : {
              frame: session.applyDelete(request.direction, request.expectedFrameEpoch),
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
      true,
      request.paintCaret
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    reply({
      id: request.id,
      ok: false,
      error: message,
      residentUnavailable: message.includes('resident input state is not ready'),
    });
  } finally {
    pendingUpdates = [];
  }
}

function hydrate(snapshot: YrsResidentWorkerSnapshot) {
  if (!session) throw new Error('Resident engine worker is not initialized');
  session.loadState(snapshot.state);
  if (snapshot.fontsRevision !== fontsRevision) {
    // A mismatched revision always carries the full font set (the client only
    // omits fonts when it knows this session's applied revision matches).
    session.clearFonts();
    for (const font of snapshot.fonts) session.registerFont(font);
    fontsRevision = snapshot.fontsRevision;
  }
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  if (snapshot.layoutWithRegions) {
    session.layoutDocumentWithRegionsJson(snapshot.layoutInput);
  } else {
    session.layoutDocumentJson(snapshot.layoutInput);
  }
  if (snapshot.selection) session.setSelection(snapshot.selection.anchor, snapshot.selection.head);
  layoutRevision = snapshot.layoutRevision;
  pendingUpdates = [];
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
  pendingUpdates = [];
  layoutRevision = 0;
  fontsRevision = -1;
  retainedFrame = null;
  glyphCache = null;
  offscreenCanvases.clear();
  offscreenBackBuffers.clear();
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
  paintCaret = false
): Promise<void> {
  retainedFrame = applyFrameDeltaOwned(retainedFrame, decodeFrameDelta(bytes));
  // The decoder's primitive-id arrays are zero-copy views into `bytes`. The
  // FrameDelta buffer is transferred to the main thread below, so retain only
  // these compact identity arrays in worker-owned memory before detaching it.
  retainedFrame = {
    ...retainedFrame,
    pages: retainedFrame.pages.map((page) => ({
      ...page,
      primitiveIds: page.primitiveIds.slice(),
    })),
  };
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
      ...(stateVector ? { stateVector } : {}),
    },
    [frame, ...updateBuffers, ...(stateVector ? [stateVector] : [])]
  );
}

async function replayOffscreen(
  force: boolean | Set<string>
): Promise<{ replayedPages: number; caretPainted: boolean }> {
  if (!retainedFrame || offscreenCanvases.size === 0) {
    return { replayedPages: 0, caretPainted: false };
  }
  if (!glyphCache && session) {
    glyphCache = new GlyphCache({
      provider: (fontId, glyphId) => session!.outlineGlyphJson(fontId, glyphId),
    });
  }
  const forceAll = force === true;
  const forcedPageIds = force instanceof Set ? force : null;
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
    const damaged =
      forceAll ||
      forcedPageIds?.has(pageIdString) === true ||
      retainedFrame.damagedPageIds.has(retainedPage.pageId);
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
  const prepared = await Promise.all(preparations);
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
  }
  return { replayedPages: prepared.length, caretPainted };
}

function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer as ArrayBuffer;
  }
  return bytes.slice().buffer;
}

function reply(response: ResidentEngineWorkerResponse, transfer: Transferable[] = []): void {
  scope.postMessage(response, transfer);
}
