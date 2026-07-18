/// <reference lib="webworker" />

import type { YrsResidentWorkerSnapshot } from './index';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import {
  presentOffscreenPageBackBuffer,
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

const scope = self as unknown as DedicatedWorkerGlobalScope;
let session: ResidentEngineSession | null = null;
let unsubscribe: (() => void) | null = null;
let pendingUpdates: Uint8Array[] = [];
let layoutRevision = 0;
let operations = Promise.resolve();
let retainedFrame: RetainedFrame | null = null;
let glyphCache: GlyphCache | null = null;
const offscreenCanvases = new Map<string, OffscreenCanvas>();
const offscreenBackBuffers = new Map<string, OffscreenCanvas>();
let activeOffscreenPageIds = new Set<string>();
let offscreenDpr = 1;
let offscreenZoom = 1;

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
      started
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
      started
    );
    return;
  }
  if (request.type === 'applyUpdate') {
    session.applyUpdate(request.update);
    if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
    return;
  }
  if (request.type === 'attachCanvases') {
    for (const { pageId, canvas } of request.pages) offscreenCanvases.set(pageId, canvas);
    activeOffscreenPageIds = new Set(request.activePageIds);
    offscreenDpr = request.devicePixelRatio;
    offscreenZoom = request.zoom;
    for (const pageId of offscreenCanvases.keys()) {
      if (!activeOffscreenPageIds.has(pageId)) {
        offscreenCanvases.delete(pageId);
        offscreenBackBuffers.delete(pageId);
      }
    }
    await replayOffscreen(true);
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
      started
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
  session.clearFonts();
  for (const font of snapshot.fonts) session.registerFont(font);
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  session.layoutDocumentJson(snapshot.layoutInput);
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
  retainedFrame = null;
  glyphCache = null;
  offscreenCanvases.clear();
  offscreenBackBuffers.clear();
  activeOffscreenPageIds.clear();
}

async function replyFrame(
  id: number,
  bytes: Uint8Array,
  engineMs: number,
  updates = pendingUpdates,
  engineProfile?: import('./index').YrsEngineApplyProfile,
  requestStarted = performance.now()
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
  const replayStarted = performance.now();
  const replayedPages = await replayOffscreen(false);
  const replayMs = performance.now() - replayStarted;
  const frame = exactBuffer(bytes);
  const updateBuffers = updates.map(exactBuffer);
  reply(
    {
      id,
      ok: true,
      frame,
      updates: updateBuffers,
      engineMs,
      workerTotalMs: performance.now() - requestStarted,
      engineProfile,
      replayMs,
      replayedPages,
      layoutRevision,
    },
    [frame, ...updateBuffers]
  );
}

async function replayOffscreen(forceAll: boolean): Promise<number> {
  if (!retainedFrame || offscreenCanvases.size === 0) return 0;
  if (!glyphCache && session) {
    glyphCache = new GlyphCache({
      provider: (fontId, glyphId) => session!.outlineGlyphJson(fontId, glyphId),
    });
  }
  const preparations: Array<
    Promise<{
      canvas: OffscreenCanvas;
      buffer: OffscreenCanvas;
      page: RetainedFrame['displayList']['pages'][number];
    }>
  > = [];
  for (let index = 0; index < retainedFrame.pages.length; index += 1) {
    const retainedPage = retainedFrame.pages[index];
    if (!forceAll && !retainedFrame.damagedPageIds.has(retainedPage.pageId)) continue;
    const canvas = offscreenCanvases.get(retainedPage.pageId.toString());
    const page = retainedFrame.displayList.pages[index];
    if (!canvas || !page) continue;
    const pageId = retainedPage.pageId.toString();
    let buffer = offscreenBackBuffers.get(pageId);
    if (!buffer) {
      buffer = new OffscreenCanvas(1, 1);
      offscreenBackBuffers.set(pageId, buffer);
    }
    preparations.push(
      rasterizeDisplayPageToBackBuffer(
        buffer,
        page,
        { glyphCache: glyphCache ?? undefined },
        offscreenDpr,
        offscreenZoom
      ).then(() => ({ canvas, buffer, page }))
    );
  }
  const prepared = await Promise.all(preparations);
  // Present only after the entire damaged frame is ready. This loop is
  // synchronous, so the compositor can never observe the renderer's clears.
  for (const { canvas, buffer } of prepared) {
    presentOffscreenPageBackBuffer(canvas, buffer);
  }
  return prepared.length;
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
