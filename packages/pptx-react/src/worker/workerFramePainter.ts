import { sizeCanvasForSlide } from '@betteroffice/pptx';
import type { PresentationCommandFrame } from '@betteroffice/pptx/worker';
import { applyPatch } from '../../../../shared/office-session/render/applyPatch';
import { pptxRasterOwnership, type PptxRenderReply, type PptxWorkerRenderClient } from '../../../../shared/office-session/render/pptxWorkerRender';
import { createCommandFrameImages, createCommandFramePainter, type CommandFramePainterOptions } from './commandFramePainter';
import { commandFrameCanvasVersion, hardReferences, invalidateCommandFrameCanvas, rasters, lease, retainCommandFrameRaster, type CommandFrameLease, type CommandFrameRaster } from './commandFrameRasters';
import { slideRasterCache, type RasterEntry } from '../../../../shared/office-session/render/pptxSlideRaster';
import type { PeerlessConfirmedFrame, PeerlessPaintCheckpoint, PeerlessViewSnapshot } from './peerlessViewStore';

type ContentReply = Exclude<PptxRenderReply, { kind: 'unsupported' | 'prerender' | 'font-generation' }>;
type PrerenderReply = Extract<PptxRenderReply, { kind: 'prerender' }>;

/** An ahead-of-time raster of one slide, painted on the stage only while it still shows the deck and the stage geometry. */
export interface PeerlessPrerender {
  request(slideId: string): void;
  drop(): void;
  /** Draws the stashed raster of `slideId` if it was painted at `version` and fits the stage; it is used at most once. */
  present(slideId: string, version: string): boolean;
  /** Whether a presented prerender is on the stage and no frame has replaced it yet. */
  presented(): boolean;
  /** Shows the displayed frame again if a presented prerender is still on the stage. */
  restore(): boolean;
}

/** The prerender of each worker frame painter. */
export const peerlessPrerenders = new WeakMap<object, PeerlessPrerender>();
const contentOf = (reply: ContentReply) =>
  reply.kind === 'raster' ? reply.content : reply.kind === 'unchanged' ? reply.baseContent : undefined;
/** Whether `reply` shows older pixels than `kept`; at equal content a raster outranks an `unchanged`. */
function olderContent(reply: ContentReply, kept: ContentReply): boolean {
  const incoming = contentOf(reply), held = contentOf(kept);
  if (incoming === undefined || held === undefined) return false;
  return incoming < held || incoming === held && kept.kind === 'raster' && reply.kind !== 'raster';
}

interface Paint {
  frame: PeerlessConfirmedFrame;
  view: PeerlessViewSnapshot;
  media: PresentationCommandFrame['media'];
  id: number;
  fallback: boolean;
  painted: boolean;
  settle(painted: boolean): void;
  reject(error: unknown): void;
}

export function createWorkerFramePainter(
  canvas: HTMLCanvasElement, client: PptxWorkerRenderClient,
  options: CommandFramePainterOptions & { mainFirstFrame?: boolean } = {},
): ReturnType<typeof createCommandFramePainter> {
  const scale = options.scale ?? 1;
  const dpr = options.dpr ?? canvas.ownerDocument.defaultView?.devicePixelRatio ?? 1;
  const factor = scale * dpr;
  const images = options.images ?? createCommandFrameImages();
  let backing = canvas.ownerDocument.createElement('canvas');
  let backingFrame: { slideId: string; epoch: number; content: number } | undefined;
  let displayed: CommandFrameLease | undefined;
  let displayedFrame: PeerlessConfirmedFrame | undefined;
  const owned = new Map<HTMLCanvasElement, Set<RasterEntry<HTMLCanvasElement>>>();
  let visibleBacking: { canvas: HTMLCanvasElement; slideId: string; content: number; version: number } | undefined;
  let pending: Paint | undefined;
  let completion = Promise.resolve(false);
  let disposed = false;
  let mainOnly = !client.supported;
  let paintId = 0;
  let fallback: ReturnType<typeof createCommandFramePainter> | undefined;
  let arrival = 0;
  let early: { reply: ContentReply; holder: object } | undefined;
  let requested = 0;
  let fontGeneration = 0;
  let prestaged = false;
  let stash: { reply: PrerenderReply; holder: object } | undefined;
  const dropStash = () => {
    if (stash) pptxRasterOwnership.release(stash.reply.bitmap, stash.holder);
    stash = undefined;
  };
  let seenEpoch = -Infinity;
  let warm = !options.mainFirstFrame;
  let viewed = warm;
  const dropEarly = () => {
    if (early?.reply.kind === 'raster') pptxRasterOwnership.release(early.reply.bitmap, early.holder);
    early = undefined;
  };
  const shadow = (reply: Extract<PptxRenderReply, { kind: 'raster' }>) => {
    if (reply.baseEpoch !== null && (backingFrame?.slideId !== reply.slideId || backingFrame.content !== reply.baseContent ||
      backing.width !== reply.width || backing.height !== reply.height)) return false;
    if (backingFrame?.slideId === reply.slideId && (reply.frameEpoch < backingFrame.epoch || reply.content < backingFrame.content)) return false;
    try {
      if ([...(owned.get(backing) ?? [])].some((entry) => hardReferences(entry) > 0)) backing = copy(backing);
      if (reply.baseEpoch === null) { backing.width = reply.width; backing.height = reply.height; }
      const ctx = backing.getContext('2d');
      if (!ctx) return false;
      applyPatch(ctx, { surface: { kind: 'slide', id: reply.slideId }, rect: reply.rect, bitmap: reply.bitmap,
        full: reply.baseEpoch === null });
      backingFrame = { slideId: reply.slideId, epoch: reply.frameEpoch, content: reply.content };
      visibleBacking = undefined;
      return true;
    } catch { backingFrame = undefined; return false; }
  };
  const checkpoints = new Set<() => void>();
  const zero = (value: HTMLCanvasElement) => { value.width = value.height = 0; };
  const current = (frame?: PeerlessConfirmedFrame) => !disposed && (options.current?.(frame) ?? true);
  const valid = (job: Paint) => current(job.frame) && pending === job && paintId === job.id;
  const copy = (source: HTMLCanvasElement) => {
    const target = canvas.ownerDocument.createElement('canvas');
    target.width = source.width; target.height = source.height;
    try {
      const ctx = target.getContext('2d');
      if (!ctx) throw new Error('Slide canvas is unavailable');
      ctx.drawImage(source, 0, 0);
      return target;
    } catch (error) { zero(target); throw error; }
  };
  const detach = () => {
    const entries = owned.get(backing);
    const references = [...(entries ?? [])].reduce((count, entry) => count + hardReferences(entry), 0);
    const displayedBacking = displayed?.raster.canvas === backing;
    if (references > (displayedBacking ? 1 : 0)) backing = copy(backing);
    else if (displayedBacking) {
      displayed!.release(); displayed = undefined; displayedFrame = undefined;
    }
  };
  const shadowed = (frame: PeerlessConfirmedFrame) => backingFrame?.slideId === frame.frame.slideId &&
    backingFrame.epoch === frame.frame.epoch;
  const adopt = (job: Paint) => {
    if (disposed || pending !== job || !job.painted || !shadowed(job.frame)) return;
    const retired = fallback, previous = displayed;
    fallback = undefined; displayed = undefined;
    retired?.dispose(); previous?.release();
    displayed = register(job.frame); displayedFrame = job.frame;
    try { options.onAdopt?.(job.frame); } finally { images.trim(); }
  };
  const cancel = () => {
    paintId += 1; arrival += 1;
    if (pending && !pending.painted) pending.settle(false);
    pending = undefined;
    fallback?.cancel();
  };
  function main(job: Paint) {
    if (!valid(job) || job.fallback) return;
    job.fallback = true;
    fallback ??= createCommandFramePainter(canvas, { ...options, images,
      current: (frame) => current(frame) && !!pending?.fallback && (!frame || pending.frame === frame),
      onCommit(frame, view) {
        if (!pending || pending.frame !== frame || !valid(pending)) return;
        prestaged = false;
        const held = retainCommandFrameRaster(frame, factor);
        try {
          visibleBacking = undefined;
          displayed?.release(); displayed = held?.retain(); displayedFrame = frame;
          pending.painted = true;
          options.onCommit?.(frame, view);
          if (!viewed && !disposed) { viewed = true; client.view(scale, dpr); client.resync(); }
          const job = pending;
          if (warm && job && shadowed(frame)) queueMicrotask(() => adopt(job));
        } finally { held?.release(); }
      },
    });
    try {
      fallback.paintFrame(job.frame, job.view, job.media);
      void fallback.whenPainted().then(job.settle, job.reject);
    } catch (error) { job.reject(error); }
  }

  function register(frame: PeerlessConfirmedFrame): CommandFrameLease {
    let byFactor = rasters.get(frame);
    if (!byFactor) { byFactor = new Map(); rasters.set(frame, byFactor); }
    const clean = backing;
    let entries = owned.get(clean);
    if (!entries) { entries = new Set(); owned.set(clean, entries); }
    const raster: CommandFrameRaster = { canvas: clean, factor, ready: true,
      subscribe: () => () => {},
      plate() { throw new Error('Text echo requires the main slide painter'); },
    };
    const entry = { raster, ready: Promise.resolve(), references: 0, cancelled: false,
      release() {
        if (entry.cancelled) return;
        entry.cancelled = true; entries!.delete(entry);
        if (!entries!.size) owned.delete(clean);
        if (byFactor!.get(factor) === entry) {
          byFactor!.delete(factor);
          if (!byFactor!.size) rasters.delete(frame);
        }
        if (!owned.has(clean) && (clean !== backing || disposed)) zero(clean);
      },
    };
    entries.add(entry);
    byFactor.set(factor, entry);
    return lease(entry);
  }

  async function receive(reply: PptxRenderReply) {
    if (reply.kind === 'font-generation') {
      fontGeneration = Math.max(fontGeneration, reply.fonts);
      if (stash && stash.reply.fonts !== fontGeneration) dropStash();
      return;
    }
    if (reply.kind === 'prerender') {
      if (disposed || reply.token !== requested || reply.fonts < fontGeneration) return;
      fontGeneration = reply.fonts;
      dropStash();
      const holder = {};
      pptxRasterOwnership.retain(reply.bitmap, holder);
      stash = { reply, holder };
      return;
    }
    if (reply.kind === 'unsupported') { mainOnly = true; dropStash(); if (pending) main(pending); return; }
    const job = pending;
    if (!disposed && reply.frameEpoch > seenEpoch && (!job || reply.frameEpoch > job.frame.frame.epoch)) {
      if (early && (early.reply.frameEpoch > reply.frameEpoch ||
        early.reply.frameEpoch === reply.frameEpoch && olderContent(reply, early.reply))) return;
      dropEarly();
      const holder = {};
      if (reply.kind === 'raster') pptxRasterOwnership.retain(reply.bitmap, holder);
      early = { reply, holder };
      return;
    }
    if (reply.kind !== 'raster' && reply.kind !== 'unchanged') {
      if (job && valid(job) && reply.frameEpoch === job.frame.frame.epoch &&
        (reply.kind !== 'main-only' || reply.slideId === job.frame.frame.slideId)) main(job);
      return;
    }
    const holder = {};
    if (reply.kind === 'raster') pptxRasterOwnership.retain(reply.bitmap, holder);
    try {
      const width = job && Math.ceil(job.frame.displayList.width * scale * dpr);
      const height = job && Math.ceil(job.frame.displayList.height * scale * dpr);
      if (reply.kind === 'unchanged') {
        if (!job || !valid(job) || job.painted && !job.fallback || reply.slideId !== job.frame.frame.slideId ||
          reply.frameEpoch !== job.frame.frame.epoch) return;
        if (backingFrame?.slideId !== reply.slideId || backingFrame.content !== reply.baseContent ||
          backing.width !== width || backing.height !== height) { client.resync(); return; }
        if (job.fallback) {
          backingFrame = { ...backingFrame, epoch: reply.frameEpoch }; warm = true;
          if (job.painted) adopt(job);
          return;
        }
      } else if (job && job.fallback && valid(job) && reply.slideId === job.frame.frame.slideId && reply.factor === factor &&
        reply.width === width && reply.height === height) {
        if (shadow(reply)) { warm = true; if (reply.frameEpoch === job.frame.frame.epoch) adopt(job); }
        else if (reply.baseEpoch !== null && reply.frameEpoch === job.frame.frame.epoch) client.resync();
        return;
      } else {
        if (!job || !valid(job) || job.fallback || reply.slideId !== job.frame.frame.slideId ||
          reply.frameEpoch !== job.frame.frame.epoch || reply.factor !== factor) return;
        if (backingFrame?.slideId === reply.slideId && reply.content < backingFrame.content) return;
        if (reply.width !== width || reply.height !== height) { client.resync(); return; }
        if (reply.baseEpoch !== null && (backingFrame?.content !== reply.baseContent || backingFrame.slideId !== reply.slideId ||
          backing.width !== reply.width || backing.height !== reply.height)) { client.resync(); return; }
      }
      const turn = ++arrival;
      if (options.beforeCommit) await options.beforeCommit(() => valid(job) && arrival === turn);
      if (!valid(job) || arrival !== turn) return;
      const patchVisible = reply.baseEpoch !== null && visibleBacking?.canvas === backing &&
        visibleBacking.slideId === reply.slideId && visibleBacking.content === reply.baseContent &&
        visibleBacking.version === commandFrameCanvasVersion(canvas) &&
        canvas.width === width && canvas.height === height;
      visibleBacking = undefined;
      if (reply.kind === 'raster') {
        detach();
        if (reply.baseEpoch === null) { backing.width = reply.width; backing.height = reply.height; }
        const ctx = backing.getContext('2d');
        if (!ctx) throw new Error('Slide canvas is unavailable');
        applyPatch(ctx, { surface: { kind: 'slide', id: reply.slideId }, rect: reply.rect, bitmap: reply.bitmap,
          full: reply.baseEpoch === null });
        backingFrame = { slideId: reply.slideId, epoch: reply.frameEpoch, content: reply.content };
      }
      const content = backingFrame!.content;
      const candidate = register(job.frame);
      let composed: ReturnType<NonNullable<CommandFramePainterOptions['compose']>> = undefined;
      try {
        composed = options.compose?.(candidate.raster, job.frame, job.view);
        const partial = !composed && patchVisible;
        if (!partial) sizeCanvasForSlide(canvas, job.frame.displayList, dpr, scale);
        const visible = canvas.getContext('2d');
        if (!visible) throw new Error('Slide canvas is unavailable');
        visible.save();
        try {
          visible.setTransform(1, 0, 0, 1, 0, 0);
          invalidateCommandFrameCanvas(canvas);
          if (partial) {
            if (reply.kind === 'raster') {
              const { x, y, width, height } = reply.rect;
              visible.clearRect(x, y, width, height);
              visible.drawImage(backing, x, y, width, height, x, y, width, height);
            }
          } else {
            visible.clearRect(0, 0, canvas.width, canvas.height);
            visible.drawImage(composed?.canvas ?? backing, 0, 0);
          }
        } finally { visible.restore(); }
        if (!composed) visibleBacking = { canvas: backing, slideId: reply.slideId, content,
          version: commandFrameCanvasVersion(canvas) };
        const previous = displayed; displayed = candidate; displayedFrame = job.frame;
        job.painted = true;
        prestaged = false;
        try { options.onCommit?.(job.frame, job.view); } finally { previous?.release(); }
        if (fallback) {
          const retired = fallback; fallback = undefined; retired.dispose();
          images.trim();
        } else {
          for (const pooled of slideRasterCache<HTMLCanvasElement>(images).pool?.canvases.splice(0) ?? []) zero(pooled.canvas);
        }
        job.settle(true);
      } catch (error) { if (displayed !== candidate) candidate.release(); throw error; }
      finally { composed?.release(); }
    } catch { if (job && valid(job)) main(job); }
    finally { if (reply.kind === 'raster') pptxRasterOwnership.release(reply.bitmap, holder); }
  }
  function takeEarly(job: Paint): boolean {
    const stashed = early;
    if (!stashed || stashed.reply.frameEpoch > job.frame.frame.epoch) return false;
    early = undefined;
    const answers = stashed.reply.frameEpoch === job.frame.frame.epoch;
    if (answers) client.idle();
    void receive(stashed.reply).finally(() => {
      if (stashed.reply.kind === 'raster') pptxRasterOwnership.release(stashed.reply.bitmap, stashed.holder);
    });
    return answers;
  }
  const off = client.subscribe((reply) => { void receive(reply); });
  if (viewed) { client.view(scale, dpr); client.resync(); }
  const painter: ReturnType<typeof createCommandFramePainter> = {
    paintFrame(frame, view, media = []) {
      if (!current(frame) || scale <= 0 || dpr <= 0) return;
      cancel();
      images.ingest([...media, ...frame.frame.media]);
      let settle!: (painted: boolean) => void, reject!: (error: unknown) => void;
      completion = new Promise<boolean>((resolve, fail) => { settle = resolve; reject = fail; });
      const job: Paint = { frame, view, media, id: paintId, fallback: false, painted: false, settle, reject };
      pending = job;
      seenEpoch = Math.max(seenEpoch, frame.frame.epoch);
      if (mainOnly) { main(job); return; }
      if (!warm) { main(job); takeEarly(job); return; }
      if (!takeEarly(job)) client.view(scale, dpr);
    },
    whenPainted: () => completion,
    resync() {
      if (pending && takeEarly(pending)) return;
      if (!mainOnly && pending && !pending.painted && !pending.fallback) client.resync();
    },
    checkpoint(includePixels = true): PeerlessPaintCheckpoint {
      const id = paintId;
      const saved = pending;
      const savedFrame = displayedFrame;
      const savedLease = displayed?.retain();
      const savedCompletion = completion;
      let pixels: HTMLCanvasElement | undefined;
      try { pixels = includePixels && displayed ? copy(canvas) : undefined; }
      catch (error) { savedLease?.release(); throw error; }
      const width = canvas.style.width, height = canvas.style.height;
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true; checkpoints.delete(finish);
        savedLease?.release(); if (pixels) zero(pixels);
      };
      checkpoints.add(finish);
      return Object.assign(() => {
        if (settled) return;
        if (!current() || paintId > id + 1) { finish(); return; }
        cancel();
        try {
          const restored = savedLease ? copy(savedLease.raster.canvas) : canvas.ownerDocument.createElement('canvas');
          const previous = backing; backing = restored;
          displayed?.release();
          displayed = savedFrame && savedLease ? register(savedFrame) : undefined;
          displayedFrame = savedFrame;
          if (!owned.has(previous)) zero(previous);
          backingFrame = undefined; visibleBacking = undefined; completion = savedCompletion;
          if (pixels) {
            invalidateCommandFrameCanvas(canvas);
            canvas.width = pixels.width; canvas.height = pixels.height;
            canvas.style.width = width; canvas.style.height = height;
            canvas.getContext('2d')?.drawImage(pixels, 0, 0);
          }
          if (saved?.painted) pending = { ...saved, id: paintId };
          else if (saved) {
            painter.paintFrame(saved.frame, saved.view, saved.media);
            if (pending) main(pending);
          }
          client.resync();
        } finally { finish(); }
      }, { commit: finish });
    },
    cancel,
    freeze() {
      cancel(); fallback?.freeze();
      for (const finish of checkpoints) finish();
    },
    dispose() {
      if (disposed) return;
      disposed = true; cancel(); off(); dropEarly(); dropStash();
      for (const finish of checkpoints) finish();
      fallback?.dispose(); displayed?.release(); displayed = undefined;
      if (!owned.has(backing)) zero(backing);
      if (!options.images) images.dispose();
    },
  };
  peerlessPrerenders.set(painter, {
    request(slideId) {
      if (disposed || mainOnly || !client.supported) return;
      dropStash();
      requested = client.prerender(slideId);
    },
    drop() { requested = 0; dropStash(); },
    present(slideId, version) {
      const held = stash;
      requested = 0; stash = undefined;
      if (!held) return false;
      try {
        const { reply } = held;
        if (disposed || mainOnly || !warm || pending?.fallback || reply.slideId !== slideId || reply.version !== version ||
          reply.fonts !== fontGeneration || reply.factor !== factor || reply.width !== canvas.width ||
          reply.height !== canvas.height) return false;
        const visible = canvas.getContext('2d');
        if (!visible) return false;
        visible.save();
        try {
          visible.setTransform(1, 0, 0, 1, 0, 0);
          visible.clearRect(0, 0, canvas.width, canvas.height);
          visible.drawImage(reply.bitmap, 0, 0);
        } finally { visible.restore(); }
        invalidateCommandFrameCanvas(canvas);
        visibleBacking = undefined;
        prestaged = true;
        return true;
      } finally { pptxRasterOwnership.release(held.reply.bitmap, held.holder); }
    },
    presented: () => prestaged && !disposed,
    restore() {
      if (!prestaged || disposed) return false;
      prestaged = false;
      if (!backingFrame || backing.width !== canvas.width || backing.height !== canvas.height) { client.resync(); return true; }
      const visible = canvas.getContext('2d');
      if (!visible) return false;
      visible.save();
      try {
        visible.setTransform(1, 0, 0, 1, 0, 0);
        visible.clearRect(0, 0, canvas.width, canvas.height);
        visible.drawImage(backing, 0, 0);
      } finally { visible.restore(); }
      invalidateCommandFrameCanvas(canvas);
      visibleBacking = undefined;
      return true;
    },
  });
  return painter;
}
