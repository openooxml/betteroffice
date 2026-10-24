import { frameAssetIds } from '@betteroffice/pptx/worker';
import { browserFrameImages, paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx/editor';
import type {
  CanvasImageResolver, PresentationCommandFrame, SlideDisplayList, SlidePrimitive, TextBoxPrimitive,
} from '@betteroffice/pptx';
import type {
  PeerlessActiveLine, PeerlessConfirmedFrame, PeerlessPaintCheckpoint, PeerlessRect, PeerlessViewPainters,
  PeerlessViewSnapshot,
} from './peerlessViewStore';

interface RasterTile {
  canvas: HTMLCanvasElement;
  references: number;
  cancelled: boolean;
  ready: Promise<void>;
  releaseMedia?(): void;
}

interface RasterLayer {
  tile: RasterTile;
  bounds: PeerlessRect;
  box?: TextBoxPrimitive;
  line?: PeerlessActiveLine['line'];
}

export interface CommandFrameRaster {
  canvas: HTMLCanvasElement;
  factor: number;
  readonly ready: boolean;
  plate(active: PeerlessActiveLine, bounds: PeerlessRect): {
    background: HTMLCanvasElement;
    foreground: HTMLCanvasElement;
  };
  subscribe(listener: () => void): () => void;
}

interface RasterEntry {
  raster: CommandFrameRaster;
  ready: Promise<void>;
  references: number;
  cancelled: boolean;
  release(): void;
}

export interface CommandFrameLease {
  readonly raster: CommandFrameRaster;
  readonly ready: Promise<void>;
  retain(): CommandFrameLease;
  release(): void;
}

interface MediaLease {
  resolve: CanvasImageResolver;
  ready: Promise<void>;
  version(assetId: string): number | undefined;
  retain(): MediaLease;
  retainAsset(assetId: string): Pick<MediaLease, 'resolve' | 'ready' | 'release'>;
  release(): void;
}

interface PinnedAsset {
  resolve: CanvasImageResolver & { release(): void };
  source: Promise<CanvasImageSource | null>;
  references: number;
}

export function createCommandFrameImages() {
  const images = browserFrameImages();
  const media = new Map<string, { bytes: Uint8Array; version: number }>();
  const assets = new Map<string, PinnedAsset>();
  const frames = new Map<PeerlessConfirmedFrame, { resolve: CanvasImageResolver;
    versions: Map<string, number>; retainAsset: MediaLease['retainAsset']; release(): void;
    ready: Promise<void>; references: number }>();
  let disposed = false;
  let version = 0;
  const finish = () => { if (disposed && !frames.size) media.clear(); };
  const releaseAsset = (key: string, asset: PinnedAsset) => {
    if (--asset.references === 0) { assets.delete(key); asset.resolve.release(); }
  };
  return {
    pin(frame: PeerlessConfirmedFrame, supplied: PresentationCommandFrame['media'] = []): MediaLease {
      if (disposed) throw new Error('Slide images are disposed');
      let entry = frames.get(frame);
      for (const asset of entry ? supplied : [...supplied, ...frame.frame.media]) {
        const previous = media.get(asset.assetId);
        if (previous?.bytes.buffer === asset.bytes) continue;
        const bytes = new Uint8Array(asset.bytes);
        if (!previous || previous.bytes.length !== bytes.length ||
          !previous.bytes.every((byte, index) => byte === bytes[index])) {
          media.set(asset.assetId, { bytes, version: ++version });
        } else previous.bytes = bytes;
      }
      if (!entry) {
        const heldAssets = new Map<string, { key: string; asset: PinnedAsset }>();
        const versions = new Map<string, number>();
        for (const id of frameAssetIds(frame.displayList)) {
          const value = media.get(id);
          if (!value) continue;
          const key = JSON.stringify([id, value.version]);
          let asset = assets.get(key);
          if (!asset) {
            const resolve = images.resolve({ media: new Map([[key, value.bytes]]) });
            asset = { resolve, source: Promise.resolve(resolve(key)), references: 0 };
            assets.set(key, asset);
          }
          asset.references += 1;
          heldAssets.set(id, { key, asset });
          versions.set(id, value.version);
        }
        entry = { resolve: (id) => heldAssets.get(id)?.asset.source ?? null, versions,
          ready: Promise.all([...heldAssets.values()].map(({ asset }) => asset.source)).then(() => {}),
          retainAsset(id) {
            const value = heldAssets.get(id);
            if (value) value.asset.references += 1;
            let released = false;
            return { resolve: (assetId) => !released && assetId === id ? value?.asset.source ?? null : null,
              ready: (value?.asset.source ?? Promise.resolve(null)).then(() => {}),
              release() {
                if (released) return;
                released = true;
                if (value) releaseAsset(value.key, value.asset);
              } };
          },
          release() {
            for (const { key, asset } of heldAssets.values()) releaseAsset(key, asset);
            heldAssets.clear();
          },
          references: 0 };
        frames.set(frame, entry);
      }
      const held = entry;
      const acquire = (): MediaLease => {
        held.references += 1;
        let released = false;
        return { resolve: held.resolve, ready: held.ready, version: (id) => held.versions.get(id),
          retain: acquire, retainAsset: held.retainAsset, release() {
            if (released) return;
            released = true;
            if (--held.references === 0) { frames.delete(frame); held.release(); finish(); }
          } };
      };
      return acquire();
    },
    dispose() { if (!disposed) { disposed = true; images.dispose(); finish(); } },
  };
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
}

const rasters = new WeakMap<PeerlessConfirmedFrame, Map<number, RasterEntry>>();
const tiles = new WeakMap<CommandFrameImages, Map<string, RasterTile>>();
export { decodeCommandDisplayList } from '@betteroffice/pptx/worker';

export function commandFrameRaster(frame: PeerlessConfirmedFrame, scale = 1,
  dpr = typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1): CommandFrameRaster | undefined {
  return rasters.get(frame)?.get(scale * dpr)?.raster;
}

function lease(entry: RasterEntry): CommandFrameLease {
  if (entry.cancelled) throw new Error('Slide raster is released');
  entry.references += 1;
  let released = false;
  return { raster: entry.raster, ready: entry.ready, retain: () => lease(entry), release() {
    if (released) return;
    released = true;
    if (--entry.references === 0) entry.release();
  } };
}

export function retainCommandFrameRaster(frame: PeerlessConfirmedFrame, factor: number): CommandFrameLease | undefined {
  const entry = rasters.get(frame)?.get(factor);
  return entry && !entry.cancelled ? lease(entry) : undefined;
}

export function createCommandFramePainter(
  canvas: HTMLCanvasElement, options: CommandFramePainterOptions = {},
): Pick<PeerlessViewPainters, 'paintFrame'> & {
  checkpoint(pixels?: boolean): PeerlessPaintCheckpoint;
  whenPainted(): Promise<boolean>;
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
          let byKey = tiles.get(images);
          if (!byKey) { byKey = new Map(); tiles.set(images, byKey); }
          entry = buildRaster(canvas.ownerDocument, frame.displayList, dpr, scale, heldMedia, byKey, jobs, () => {
            byFactor!.delete(factor);
            if (!byFactor!.size) rasters.delete(frame);
          });
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

function zero(canvas: HTMLCanvasElement): void { canvas.width = 0; canvas.height = 0; }

function guardedContext(ctx: CanvasRenderingContext2D, current: () => boolean,
  transform?: () => void): CanvasRenderingContext2D {
  return new Proxy(ctx, {
    get(target, property) {
      if (transform && property === 'clearRect') return () => {};
      if (transform && property === 'setTransform') return () => { if (current()) transform(); };
      const value: unknown = Reflect.get(target, property, target);
      return typeof value === 'function' ? (...args: unknown[]) => current() ? value.apply(target, args) : undefined : value;
    },
    set: (target, property, value: unknown) => !current() || Reflect.set(target, property, value, target),
  });
}

function buildRaster(
  owner: Document, list: SlideDisplayList, dpr: number, scale: number, media: MediaLease,
  byKey: Map<string, RasterTile>, jobs: Promise<void>[], remove: () => void,
): RasterEntry {
  const factor = dpr * scale;
  const authoritative = owner.createElement('canvas');
  sizeCanvasForSlide(authoritative, list, dpr, scale);
  const ctx = authoritative.getContext('2d');
  if (!ctx) { zero(authoritative); throw new Error('Slide canvas is unavailable'); }
  const ordered: RasterLayer[] = [];
  const held = new Map<RasterTile, string>();
  const listeners = new Set<() => void>();
  let cancelled = false;
  let ready = false;
  let releasedMedia = false;
  const releaseMedia = () => { if (!releasedMedia) { releasedMedia = true; media.release(); } };
  const current = () => !cancelled;
  const release = () => {
    if (cancelled) return;
    cancelled = true;
    remove();
    listeners.clear();
    zero(authoritative);
    for (const [tile, key] of held) if (--tile.references === 0) {
      tile.cancelled = true;
      tile.releaseMedia?.();
      if (byKey.get(key) === tile) byKey.delete(key);
      zero(tile.canvas);
    }
    held.clear();
    releaseMedia();
  };

  function add(primitive: SlidePrimitive, parents: readonly SlidePrimitive[], box?: TextBoxPrimitive): void {
    const bounds = layerBounds(primitive, parents, factor, list);
    if (!bounds) return;
    const content = primitive.kind === 'textBox' ? { ...primitive, paragraphs: [] } : primitive;
    const key = JSON.stringify([content, parents.map((parent) => ({
      x: parent.x, y: parent.y, w: parent.w, h: parent.h, transform: parent.transform,
    })), factor, bounds, primitive.kind === 'image' && primitive.assetId ? media.version(primitive.assetId) : null]);
    let tile = byKey.get(key);
    if (!tile) {
      const canvas = owner.createElement('canvas');
      canvas.width = Math.ceil(bounds.width * factor);
      canvas.height = Math.ceil(bounds.height * factor);
      const context = canvas.getContext('2d');
      if (!context) { zero(canvas); throw new Error('Slide canvas is unavailable'); }
      const imageMedia = primitive.kind === 'image' && primitive.assetId ? media.retainAsset(primitive.assetId) : undefined;
      let painting = false;
      const item: RasterTile = { canvas, references: 0, cancelled: false, ready: Promise.resolve(),
        releaseMedia() { if (!painting) imageMedia?.release(); } };
      const proxy = guardedContext(context, () => !item.cancelled, () => {
        context.setTransform(factor, 0, 0, factor, -bounds.x * factor, -bounds.y * factor);
        for (const parent of parents) {
          applyPrimitiveTransform(context, parent);
          context.beginPath(); context.rect(parent.x, parent.y, parent.w, parent.h); context.clip();
        }
      });
      const paint = () => {
        if (item.cancelled) return;
        painting = true;
        return paintSlide(proxy, { ...list, background: undefined, primitives: [primitive] }, dpr, scale,
          { resolveImage: (assetId) => (imageMedia ?? media).resolve(assetId) });
      };
      try {
        item.ready = imageMedia ? imageMedia.ready.then(paint).finally(() => imageMedia.release()) : Promise.resolve(paint());
      } catch (error) { imageMedia?.release(); item.ready = Promise.reject(error); }
      void item.ready.catch(() => {});
      tile = item;
      byKey.set(key, tile);
    }
    if (!held.has(tile)) { held.set(tile, key); tile.references += 1; }
    jobs.push(tile.ready);
    ordered.push({ tile, bounds, box, line: primitive.kind === 'textBox' ? primitive.lines[0] : undefined });
  }

  function visit(primitives: readonly SlidePrimitive[], parents: readonly SlidePrimitive[]): void {
    for (const primitive of primitives) {
      if (primitive.kind === 'chart' || primitive.kind === 'table') visit(primitive.primitives, [...parents, primitive]);
      else if (primitive.kind === 'textBox') {
        for (const line of primitive.lines) add({ ...primitive, lines: [line] }, parents, primitive);
      } else add(primitive, parents);
    }
  }

  const raster: CommandFrameRaster = {
    canvas: authoritative, factor, get ready() { return ready && current(); },
    subscribe(listener) { if (!cancelled) listeners.add(listener); return () => { listeners.delete(listener); }; },
    plate(active, bounds) {
      const background = owner.createElement('canvas');
      const foreground = owner.createElement('canvas');
      background.width = foreground.width = Math.ceil(bounds.width * factor);
      background.height = foreground.height = Math.ceil(bounds.height * factor);
      try {
        const back = background.getContext('2d');
        const front = foreground.getContext('2d');
        if (ready && current() && back && front) {
          const proxy = guardedContext(back, current, () => {
            back.setTransform(factor, 0, 0, factor, -bounds.x * factor, -bounds.y * factor);
          });
          void paintSlide(proxy, { ...list, primitives: [] }, dpr, scale).catch(() => {});
          let removed = false;
          for (const layer of ordered) {
            if (layer.box?.objectId === active.box.objectId && layer.line?.start === active.line.start &&
              layer.line.end === active.line.end) { removed = true; continue; }
            const target = removed ? front : back;
            target.drawImage(layer.tile.canvas, (layer.bounds.x - bounds.x) * factor, (layer.bounds.y - bounds.y) * factor);
          }
        }
        return { background, foreground };
      } catch (error) { zero(background); zero(foreground); throw error; }
    },
  };
  try { visit(list.primitives, []); } catch (error) { release(); throw error; }
  const completion = Promise.allSettled(jobs).then(async (results) => {
    for (const result of results) if (result.status === 'rejected') throw result.reason;
    if (!current()) return;
    await paintSlide(guardedContext(ctx, current), { ...list, primitives: [] }, dpr, scale);
    if (!current()) return;
    for (const layer of ordered) ctx.drawImage(layer.tile.canvas, layer.bounds.x * factor, layer.bounds.y * factor);
    ready = true;
    for (const listener of listeners) listener();
  }).finally(releaseMedia);
  return { raster, ready: completion, references: 0, cancelled: false,
    release() { this.cancelled = true; release(); } };
}

export function applyPrimitiveTransform(ctx: CanvasRenderingContext2D, primitive: SlidePrimitive): void {
  if (!primitive.transform) return;
  const cx = primitive.x + primitive.w / 2;
  const cy = primitive.y + primitive.h / 2;
  ctx.translate(cx, cy);
  ctx.rotate((primitive.transform.rotationDeg ?? 0) * Math.PI / 180);
  ctx.scale(primitive.transform.flipH ? -1 : 1, primitive.transform.flipV ? -1 : 1);
  ctx.translate(-cx, -cy);
}

function layerBounds(
  primitive: SlidePrimitive, parents: readonly SlidePrimitive[], factor: number, list: SlideDisplayList,
): PeerlessRect | null {
  const line = primitive.kind === 'textBox' ? primitive.lines[0] : undefined;
  const shadow = primitive.kind === 'textBox' ? primitive.textShadow
    : primitive.kind === 'shape' || primitive.kind === 'image' ? primitive.shadow : undefined;
  const stroke = primitive.kind === 'shape' || primitive.kind === 'image' ? primitive.stroke : undefined;
  const ends = [stroke?.headEnd, stroke?.tailEnd];
  const outline = (stroke?.width ?? 0) * ((stroke?.join ?? 'miter') === 'miter' ? 5 : 0.5) + Math.max(0, ...ends.map((end) =>
    end && end.kind !== 'none' ? Math.hypot(end.length, end.width / 2) : 0));
  const spread = outline + Math.abs(shadow?.dx ?? 0) + Math.abs(shadow?.dy ?? 0) + (shadow?.blur ?? 0) * 3 + 2;
  const rect = line ? { x: Math.min(primitive.x, line.x), y: line.y,
    w: Math.max(primitive.x + primitive.w, line.x + line.width) - Math.min(primitive.x, line.x),
    h: line.height } : primitive;
  const points = [{ x: rect.x - spread, y: rect.y - spread },
    { x: rect.x + rect.w + spread, y: rect.y - spread },
    { x: rect.x - spread, y: rect.y + rect.h + spread },
    { x: rect.x + rect.w + spread, y: rect.y + rect.h + spread }];
  const transforms = [...parents, primitive];
  for (let i = transforms.length - 1; i >= 0; i -= 1) {
    const item = transforms[i];
    const cx = item.x + item.w / 2; const cy = item.y + item.h / 2;
    const angle = (item.transform?.rotationDeg ?? 0) * Math.PI / 180;
    const cos = Math.cos(angle); const sin = Math.sin(angle);
    for (const point of points) {
      const dx = (point.x - cx) * (item.transform?.flipH ? -1 : 1);
      const dy = (point.y - cy) * (item.transform?.flipV ? -1 : 1);
      point.x = cx + dx * cos - dy * sin;
      point.y = cy + dx * sin + dy * cos;
    }
  }
  const x = Math.max(0, Math.floor(Math.min(...points.map((point) => point.x)) * factor) / factor);
  const y = Math.max(0, Math.floor(Math.min(...points.map((point) => point.y)) * factor) / factor);
  const right = Math.min(list.width, Math.ceil(Math.max(...points.map((point) => point.x)) * factor) / factor);
  const bottom = Math.min(list.height, Math.ceil(Math.max(...points.map((point) => point.y)) * factor) / factor);
  return right > x && bottom > y ? { x, y, width: right - x, height: bottom - y } : null;
}
