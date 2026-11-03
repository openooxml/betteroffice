import { frameAssetIds, browserFrameImages } from '@betteroffice/pptx/worker';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type { SlideDisplayList } from '@betteroffice/pptx';
import type { PresentationCommandFrame } from '@betteroffice/pptx/worker';
import { buildRaster, createSlideRasterImages, slideRasterCache, zero, type MediaLease } from '../../../../shared/office-session/render/pptxSlideRaster';
import { invalidateCommandFrameCanvas, rasters, lease, type CommandFrameLease, type CommandFrameRaster } from './commandFrameRasters';
import type { PeerlessConfirmedFrame, PeerlessPaintCheckpoint, PeerlessViewPainters, PeerlessViewSnapshot } from './peerlessViewStore';

export { commandFrameRaster, retainCommandFrameRaster } from './commandFrameRasters';
export type { CommandFrameRaster, CommandFrameLease } from './commandFrameRasters';
export { applyPrimitiveTransform } from '../../../../shared/office-session/render/pptxSlideRaster';
export { decodeCommandDisplayList } from '@betteroffice/pptx/worker';

export function createCommandFrameImages(options: { maxDimension?: number; reuse?: object } = {}) {
  return createSlideRasterImages<HTMLCanvasElement, SlideDisplayList>({
    images: () => browserFrameImages({ maxDimension: options.maxDimension }), frameAssetIds, reuse: options.reuse });
}
export type CommandFrameImages = ReturnType<typeof createCommandFrameImages>;

export interface CommandFramePainterOptions {
  scale?: number;
  dpr?: number;
  images?: CommandFrameImages;
  current?(frame?: PeerlessConfirmedFrame): boolean;
  beforeCommit?(current: () => boolean): Promise<void>;
  compose?(raster: CommandFrameRaster, frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot): {
    canvas: HTMLCanvasElement; release(): void;
  } | undefined;
  onCommit?(frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot): void;
  onAdopt?(frame: PeerlessConfirmedFrame): void;
}

export function createCommandFramePainter(
  canvas: HTMLCanvasElement, options: CommandFramePainterOptions = {},
): Pick<PeerlessViewPainters, 'paintFrame'> & {
  checkpoint(pixels?: boolean): PeerlessPaintCheckpoint;
  whenPainted(): Promise<boolean>;
  resync?(): void;
  cancel(): void;
  freeze(): void;
  dispose(): void;
} {
  const scale = options.scale ?? 1;
  const dpr = options.dpr ?? canvas.ownerDocument.defaultView?.devicePixelRatio ?? 1;
  const images = options.images ?? createCommandFrameImages();
  let disposed = false;
  let paintId = 0;
  let displayed: CommandFrameLease | undefined;
  let activeMedia: MediaLease | undefined;
  let pending: { lease: CommandFrameLease; frame: PeerlessConfirmedFrame; view: PeerlessViewSnapshot;
    media: PresentationCommandFrame['media']; cancel(): void } | undefined;
  const checkpoints = new Set<() => void>();
  let completion = Promise.resolve(false);
  const current = (frame?: PeerlessConfirmedFrame) => !disposed && (options.current?.(frame) ?? true);
  const cancel = () => { paintId += 1; pending?.cancel(); pending = undefined; };
  const painter: ReturnType<typeof createCommandFramePainter> = {
    checkpoint(includePixels = true) {
      const savedPaintId = paintId;
      const savedPending = pending && { ...pending, lease: pending.lease.retain() };
      const savedLease = displayed?.retain();
      const savedCompletion = completion;
      const savedMedia = activeMedia?.retain();
      let pixels: HTMLCanvasElement | undefined;
      if (includePixels && displayed) {
        pixels = canvas.ownerDocument.createElement('canvas');
        pixels.width = canvas.width;
        pixels.height = canvas.height;
        try {
          const ctx = pixels.getContext('2d');
          if (!ctx) throw new Error('Slide canvas is unavailable');
          ctx.drawImage(canvas, 0, 0);
        } catch (error) {
          savedLease?.release(); savedPending?.lease.release(); savedMedia?.release();
          zero(pixels); throw error;
        }
      }
      const width = canvas.style.width;
      const height = canvas.style.height;
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        checkpoints.delete(finish);
        savedLease?.release(); savedPending?.lease.release(); savedMedia?.release();
        if (pixels) zero(pixels);
      };
      checkpoints.add(finish);
      return Object.assign(() => {
        if (settled) return;
        if (!current() || paintId > savedPaintId + 1) { finish(); return; }
        cancel();
        try {
          displayed?.release();
          displayed = savedLease?.retain();
          activeMedia?.release(); activeMedia = savedMedia?.retain();
          completion = savedCompletion;
          if (pixels) {
            invalidateCommandFrameCanvas(canvas);
            canvas.width = pixels.width;
            canvas.height = pixels.height;
            canvas.style.width = width;
            canvas.style.height = height;
            canvas.getContext('2d')?.drawImage(pixels, 0, 0);
          }
          if (savedPending) painter.paintFrame(savedPending.frame, savedPending.view, savedPending.media);
        } finally { finish(); }
      }, { commit: finish });
    },
    paintFrame(frame, view, suppliedMedia = []) {
      if (!current(frame) || scale <= 0 || dpr <= 0) return;
      cancel();
      const id = paintId;
      const valid = () => current(frame) && paintId === id;
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('Slide canvas is unavailable');
      const nextMedia = images.pin(frame, suppliedMedia);
      activeMedia?.release();
      activeMedia = nextMedia;
      let byFactor = rasters.get(frame);
      if (!byFactor) { byFactor = new Map(); rasters.set(frame, byFactor); }
      const factor = scale * dpr;
      let entry = byFactor.get(factor);
      if (!entry) {
        const heldMedia = nextMedia.retain();
        const jobs = [heldMedia.ready];
        try {
          const { byKey, pool } = slideRasterCache<HTMLCanvasElement>(images);
          entry = buildRaster(() => canvas.ownerDocument.createElement('canvas'), frame.displayList, dpr, scale, heldMedia, byKey,
            pool, jobs, () => {
            if (byFactor!.get(factor) !== entry) return;
            byFactor!.delete(factor);
            if (!byFactor!.size) rasters.delete(frame);
          }, { paintSlide, sizeCanvasForSlide });
          byFactor.set(factor, entry);
        } catch (error) {
          activeMedia?.release(); activeMedia = undefined;
          heldMedia.release();
          completion = Promise.allSettled(jobs).then(() => false);
          throw error;
        }
      }
      const candidate = lease(entry);
      let settle!: (painted: boolean) => void;
      let reject!: (error: unknown) => void;
      let finished = false;
      completion = new Promise<boolean>((done, fail) => { settle = done; reject = fail; });
      const job = { lease: candidate, frame, view, media: suppliedMedia, cancel() {
        if (finished) return;
        finished = true;
        candidate.release();
        if (pending === job) {
          pending = undefined;
          activeMedia?.release(); activeMedia = undefined;
        }
        void candidate.ready.then(() => settle(false), () => settle(false));
      } };
      pending = job;
      void candidate.ready.then(async () => {
        if (!valid() || finished) { job.cancel(); return; }
        if (options.beforeCommit) await options.beforeCommit(valid);
        if (!valid() || finished) { job.cancel(); return; }
        const composed = options.compose?.(candidate.raster, frame, view);
        try {
          if (!valid() || finished) { job.cancel(); return; }
          invalidateCommandFrameCanvas(canvas);
          sizeCanvasForSlide(canvas, frame.displayList, dpr, scale);
          ctx.save();
          ctx.setTransform(1, 0, 0, 1, 0, 0);
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(composed?.canvas ?? candidate.raster.canvas, 0, 0);
          ctx.restore();
          const previous = displayed;
          displayed = candidate;
          pending = undefined;
          finished = true;
          try { options.onCommit?.(frame, view); } finally { previous?.release(); }
        } finally { composed?.release(); }
        settle(true);
      }).catch((error: unknown) => {
        if (finished) { if (valid()) reject(error); else settle(false); return; }
        finished = true;
        candidate.release();
        if (pending === job) {
          pending = undefined;
          activeMedia?.release(); activeMedia = undefined;
        }
        if (valid()) reject(error); else settle(false);
      });
    },
    whenPainted: () => completion,
    resync() {},
    cancel,
    freeze() {
      cancel();
      for (const finish of checkpoints) finish();
      activeMedia?.release(); activeMedia = undefined;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      cancel();
      for (const finish of checkpoints) finish();
      displayed?.release(); displayed = undefined;
      activeMedia?.release(); activeMedia = undefined;
      if (!options.images) images.dispose();
    },
  };
  return painter;
}
