import type { SessionTransport } from '../../../../shared/office-session/transport';
import { PPTX_RENDER_PORT, type PptxPaintFont, type PptxRenderReply, type PptxRenderRequest } from '../../../../shared/office-session/render/pptxWorkerRender';
import { clampToSurface } from '../../../../shared/office-session/render/rects';
import { readFontTraits } from '../../../../shared/fontTraits';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';
import type { SlideDisplayList } from '../types';
import { browserPresentationImageBlob, decodeBrowserImageBlob, needsElementDecode } from '../render/browserImage';
import { buildRaster, createSlideRasterImages, slideRasterCache, zero, type RasterEntry } from '../../../../shared/office-session/render/pptxSlideRaster';
import { paintSlide, sizeCanvasForSlide } from '../render/canvas';
import { createBrowserFrameImages } from '../render/browserFrameImages';
import { frameAssetIds } from './frame';

export interface WorkerSlideFrame {
  slideId: string;
  frameEpoch: number;
  displayList: SlideDisplayList;
  mediaBytes(assetId: string): Uint8Array;
}

interface WorkerSlideMedia {
  assetId: string;
  bytes: ArrayBuffer;
  mainOnly: boolean;
  error?: { cause: unknown };
}

export function createPptxWorkerRenderer(port: MessagePort) {
  const fontSet = (globalThis as typeof globalThis & { fonts?: FontFaceSet }).fonts;
  const installed: FontFace[] = [];
  const mediaCache = new Map<string, WorkerSlideMedia>();
  let mediaBlobs = new WeakMap<ArrayBufferLike, Blob>();
  const createImages = () => createSlideRasterImages<OffscreenCanvas, SlideDisplayList>({
    images: () => createBrowserFrameImages((bytes, message) => {
      const blob = mediaBlobs.get(bytes.buffer);
      return blob ? decodeBrowserImageBlob(blob, message) : Promise.reject(new Error('Slide image is unavailable'));
    }), frameAssetIds,
  });
  let images = createImages();
  let latest: WorkerSlideFrame | undefined;
  let view: { scale: number; dpr: number } | undefined;
  let revision = 0;
  let generation = 0;
  let fontWork = Promise.resolve();
  let fontError: unknown;
  let disposed = false;
  let running = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let previous: { slideId: string; epoch: number; content: number; factor: number; width: number; height: number;
    list: SlideDisplayList; } | undefined;
  let contents = 0;
  let held: RasterEntry<OffscreenCanvas> | undefined;
  let full = true;
  let awaitingReply = false;
  let supported = false;
  let inline: { resolve(reply?: PptxRenderReply): void } | undefined;
  let parked: (readonly PptxPaintFont[])[] | undefined = [];
  const post = (reply: PptxRenderReply, transfer: Transferable[] = []) => {
    if (inline) { inline.resolve(reply); return; }
    try { port.postMessage(reply, transfer); }
    catch { if (reply.kind === 'raster') reply.bitmap.close(); dispose(); }
  };
  const crop = (source: OffscreenCanvas, rect: { x: number; y: number; width: number; height: number }) => {
    if (typeof source.transferToImageBitmap === 'function') {
      const target = new OffscreenCanvas(rect.width, rect.height);
      try {
        const ctx = target.getContext('2d');
        if (ctx) {
          ctx.drawImage(source, rect.x, rect.y, rect.width, rect.height, 0, 0, rect.width, rect.height);
          return target.transferToImageBitmap();
        }
      } finally { zero(target); }
    }
    return createImageBitmap(source, rect.x, rect.y, rect.width, rect.height);
  };
  try {
    if (typeof OffscreenCanvas === 'function' && typeof createImageBitmap === 'function' &&
      typeof FontFace === 'function' && fontSet && typeof fontSet.add === 'function' && typeof fontSet.delete === 'function') {
      const probe = new OffscreenCanvas(1, 1);
      try { supported = !!probe.getContext('2d'); } finally { zero(probe); }
    }
  } catch {}
  if (!supported) post({ kind: 'unsupported' });

  function schedule() {
    if (disposed || !supported || running || timer !== undefined || !latest || !view || awaitingReply) return;
    timer = setTimeout(() => { timer = undefined; void paint(); }, 0);
  }

  async function paint() {
    if (disposed || !latest || !view || awaitingReply) return;
    running = true;
    const frame = latest, viewport = view, id = revision, fonts = generation;
    const valid = () => !disposed && revision === id && generation === fonts;
    let entry: RasterEntry<OffscreenCanvas> | undefined;
    try {
      await fontWork;
      if (!valid()) return;
      if (fontError) throw fontError;
      if (!full && previous && previous.slideId === frame.slideId && previous.list === frame.displayList &&
        previous.factor === viewport.scale * viewport.dpr) {
        post({ kind: 'unchanged', slideId: frame.slideId, frameEpoch: frame.frameEpoch, baseEpoch: previous.epoch,
          baseContent: previous.content });
        return;
      }
      if (inline && (full || !previous || previous.slideId !== frame.slideId ||
        [...frameAssetIds(frame.displayList)].some((assetId) => !mediaCache.has(assetId)))) {
        inline.resolve();
        await new Promise((resolve) => setTimeout(resolve, 0));
        if (!valid()) return;
      }
      const media = [...frameAssetIds(frame.displayList)].flatMap((assetId) => {
        if (!mediaCache.has(assetId)) {
          let view: Uint8Array;
          try { view = frame.mediaBytes(assetId); } catch { return []; }
          const asset: WorkerSlideMedia = { assetId, bytes: new ArrayBuffer(1), mainOnly: false };
          try {
            const blob = browserPresentationImageBlob(view);
            asset.mainOnly = needsElementDecode(blob);
            if (!asset.mainOnly) mediaBlobs.set(asset.bytes, blob);
          } catch (cause) { asset.error = { cause }; }
          mediaCache.set(assetId, asset);
        }
        const asset = mediaCache.get(assetId);
        if (asset?.error) throw asset.error.cause;
        return asset ? [asset] : [];
      });
      const mainOnly = media.filter((asset) => asset.mainOnly).map(({ assetId }) => assetId);
      if (mainOnly.length) {
        full = true;
        post({ kind: 'main-only', slideId: frame.slideId, frameEpoch: frame.frameEpoch, assetIds: mainOnly });
        return;
      }
      const factor = viewport.scale * viewport.dpr;
      const width = Math.ceil(frame.displayList.width * viewport.scale * viewport.dpr);
      const height = Math.ceil(frame.displayList.height * viewport.scale * viewport.dpr);
      const replace = full || !previous || previous.slideId !== frame.slideId || previous.factor !== factor ||
        previous.width !== width || previous.height !== height;
      const baseEpoch = replace ? null : previous!.epoch;
      const lease = images.pin({ displayList: frame.displayList, frame: { media } });
      const jobs = [lease.ready];
      const cache = slideRasterCache<OffscreenCanvas>(images);
      try {
        entry = buildRaster((w, h) => new OffscreenCanvas(w, h), frame.displayList, viewport.dpr, viewport.scale,
          lease, cache.byKey, cache.pool, jobs, () => {}, { paintSlide, sizeCanvasForSlide });
      } catch (error) { lease.release(); throw error; }
      held?.release(); held = entry;
      await entry.ready;
      if (!valid()) { full = true; return; }
      const damage = replace ? { x: 0, y: 0, width, height } : entry.raster.damage;
      const rect = clampToSurface(damage ?? { x: 0, y: 0, width: 1, height: 1 }, width, height);
      const bitmap = await crop(entry.raster.canvas, rect);
      if (!valid()) { bitmap.close(); full = true; return; }
      const content = ++contents;
      post({ kind: 'raster', slideId: frame.slideId, frameEpoch: frame.frameEpoch, factor,
        baseEpoch, content, baseContent: replace ? null : previous!.content, rect, bitmap, width, height }, [bitmap]);
      previous = { slideId: frame.slideId, epoch: frame.frameEpoch, content, factor, width, height, list: frame.displayList };
      full = false;
    } catch (error) {
      full = true;
      if (valid()) post({ kind: 'error', frameEpoch: frame.frameEpoch, message: error instanceof Error ? error.message : String(error) });
    } finally {
      if (entry) {
        if (!disposed && entry.raster.ready) entry.recycle?.();
        else { entry.release(); if (held === entry) held = undefined; }
      }
      running = false;
      inline?.resolve();
      if (!disposed && (revision !== id || generation !== fonts)) schedule();
    }
  }

  function fonts(faces: readonly PptxPaintFont[]) {
    if (disposed || !supported || !faces.length) return;
    if (parked) { parked.push(faces); return; }
    generation += 1;
    fontWork = fontWork.then(async () => {
      for (const face of faces) {
        if (disposed) return;
        const loaded = await new FontFace(face.family, face.source, face.descriptors).load();
        if (disposed) return;
        fontSet!.add(loaded); installed.push(loaded);
      }
    }).catch((error: unknown) => { fontError = error; }).then(() => {
      if (disposed) return;
      full = true;
      held?.release(); held = undefined;
      const cache = slideRasterCache<OffscreenCanvas>(images);
      for (const raster of cache.pool.canvases) zero(raster.canvas);
      cache.pool.canvases.length = 0;
      revision += 1; schedule();
    });
    revision += 1; schedule();
  }

  function resetMedia() {
    if (disposed) return;
    revision += 1;
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
    latest = undefined;
    awaitingReply = false;
    full = true;
    previous = undefined;
    held?.release(); held = undefined;
    images.dispose(); images = createImages();
    mediaCache.clear(); mediaBlobs = new WeakMap();
  }

  function dispose() {
    if (disposed) return;
    disposed = true;
    if (timer !== undefined) clearTimeout(timer);
    held?.release(); held = undefined;
    images.dispose(); mediaCache.clear(); mediaBlobs = new WeakMap();
    for (const face of installed) fontSet?.delete(face);
    installed.length = 0; latest = undefined;
    port.close();
  }

  port.onmessage = (event: MessageEvent<PptxRenderRequest>) => {
    const message = event.data;
    if (message.kind === 'dispose') { dispose(); return; }
    if (disposed || !supported) return;
    if (message.kind === 'fonts') { fonts(message.faces); return; }
    if (message.kind === 'view') {
      if (!(Number.isFinite(message.scale) && Number.isFinite(message.dpr) && message.scale > 0 && message.dpr > 0)) return;
      if (view?.scale === message.scale && view.dpr === message.dpr) return;
      view = { scale: message.scale, dpr: message.dpr };
      if (parked) { const queued = parked; parked = undefined; for (const faces of queued) fonts(faces); }
    }
    full = true; revision += 1; schedule();
  };
  port.onmessageerror = dispose;
  port.start();
  return {
    frame(frame: WorkerSlideFrame) {
      if (disposed || !supported) return;
      latest = frame; awaitingReply = true; revision += 1;
    },
    afterReply() { awaitingReply = false; schedule(); },
    pending: () => !disposed && supported && awaitingReply && !!latest && !!view,
    paintNow(): Promise<PptxRenderReply | undefined> {
      if (disposed || !supported || !awaitingReply || !latest || !view || running || inline) return Promise.resolve(undefined);
      awaitingReply = false;
      if (timer !== undefined) { clearTimeout(timer); timer = undefined; }
      return new Promise((resolve) => {
        const slot = { resolve(reply?: PptxRenderReply) {
          if (inline !== slot) return;
          inline = undefined; clearTimeout(deadline); resolve(reply);
        } };
        const deadline = setTimeout(() => slot.resolve(), 0);
        inline = slot;
        void paint();
      });
    },
    fonts,
    resetMedia,
    defaultFonts(faces: readonly DefaultFontFace[]) {
      fonts(faces.map((face) => {
        const traits = face.ownWeight ? undefined : readFontTraits(new Uint8Array(face.bytes));
        return { family: face.family, source: face.bytes.slice(0), descriptors: {
          weight: traits ? String(traits.weight) : face.bold ? '700' : '400',
          style: traits ? traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal' : face.italic ? 'italic' : 'normal',
        } };
      }));
    },
    dispose,
  };
}

function renderPortMessage(message: unknown): message is { kind: typeof PPTX_RENDER_PORT; port: MessagePort } {
  return !!message && typeof message === 'object' && 'kind' in message && message.kind === PPTX_RENDER_PORT && 'port' in message;
}

export function listenPptxWorkerRender(transport: SessionTransport) {
  let renderer: ReturnType<typeof createPptxWorkerRenderer> | undefined;
  const off = transport.listen((message) => {
    if (!renderPortMessage(message)) return;
    renderer?.dispose();
    renderer = createPptxWorkerRenderer(message.port);
  });
  return {
    transport: { ...transport, listen: (listener) => transport.listen((message) => {
      if (!renderPortMessage(message)) listener(message);
    }) } satisfies SessionTransport,
    frame: (frame: WorkerSlideFrame) => renderer?.frame(frame),
    active: () => !!renderer,
    afterReply: () => renderer?.afterReply(),
    pending: () => renderer?.pending() ?? false,
    paintNow: () => renderer?.paintNow() ?? Promise.resolve(undefined),
    resetMedia: () => renderer?.resetMedia(),
    defaultFonts: (faces: readonly DefaultFontFace[]) => renderer?.defaultFonts(faces),
    dispose() { off(); renderer?.dispose(); renderer = undefined; },
  };
}
