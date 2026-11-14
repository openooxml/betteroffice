interface RasterTransform { rotationDeg?: number; flipH?: boolean; flipV?: boolean }
interface RasterPrimitiveBase {
  objectId: number;
  x: number; y: number; w: number; h: number;
  transform?: RasterTransform;
}
interface RasterShadow { dx?: number; dy?: number; blur?: number }
interface RasterStroke {
  width: number;
  join?: string;
  headEnd?: { kind: string; length: number; width: number };
  tailEnd?: { kind: string; length: number; width: number };
}
interface RasterTextLine { x: number; y: number; width: number; height: number; start: number; end: number }
interface RasterTextBox extends RasterPrimitiveBase {
  kind: 'textBox';
  paragraphs: unknown[];
  lines: RasterTextLine[];
  textShadow?: RasterShadow;
}
type RasterPrimitive = RasterTextBox
  | (RasterPrimitiveBase & { kind: 'shape' | 'image'; assetId?: string; shadow?: RasterShadow; stroke?: RasterStroke })
  | (RasterPrimitiveBase & { kind: 'chart' | 'table'; primitives: RasterPrimitive[] })
  | (RasterPrimitiveBase & { kind: 'placeholder' });
interface RasterDisplayList { width: number; height: number; background?: unknown; primitives: RasterPrimitive[] }
type RasterImageResolver = ((assetId: string) => CanvasImageSource | Promise<CanvasImageSource | null> | null) & {
  acquire?(): RasterImageResolver;
  release?(): void;
};
interface RasterImageFactory {
  resolve(frame: { media: ReadonlyMap<string, Uint8Array> }): RasterImageResolver & { release(): void };
  trim?(): void;
  dispose(): void;
}
interface RasterPaintOptions { resolveImage?: RasterImageResolver }

export type SlideRasterCanvas = HTMLCanvasElement | OffscreenCanvas;
export interface RasterRect { x: number; y: number; width: number; height: number }
export interface RasterActiveLine { box: RasterTextBox; line: RasterTextBox['lines'][number] }
export interface RasterMedia { assetId: string; bytes: ArrayBuffer }
export interface RasterFrame<L extends RasterDisplayList> { displayList: L; frame: { media: RasterMedia[] } }

export interface RasterTile<C extends SlideRasterCanvas> {
  id: number;
  canvas: C;
  references: number;
  cancelled: boolean;
  ready: Promise<void>;
  releaseMedia?(): void;
}

interface RasterLayer<C extends SlideRasterCanvas> {
  tile: RasterTile<C>;
  bounds: RasterRect;
  box?: RasterTextBox;
  line?: RasterActiveLine['line'];
}

interface PooledLayer {
  id: number;
  bounds: RasterRect;
}

interface PooledRaster<C extends SlideRasterCanvas> {
  canvas: C;
  background: string;
  layers: readonly PooledLayer[];
}

export interface RasterPool<C extends SlideRasterCanvas> {
  canvases: PooledRaster<C>[];
  disposed: boolean;
}

export interface SlideRaster<C extends SlideRasterCanvas> {
  canvas: C;
  factor: number;
  readonly damage?: RasterRect;
  readonly ready: boolean;
  plate(active: RasterActiveLine, bounds: RasterRect): {
    background: C;
    foreground: C;
  };
  subscribe(listener: () => void): () => void;
}

export interface RasterEntry<C extends SlideRasterCanvas> {
  raster: SlideRaster<C>;
  ready: Promise<void>;
  references: number;
  cancelled: boolean;
  recycle?(): void;
  release(): void;
}

export interface MediaLease {
  resolve: RasterImageResolver;
  ready: Promise<void>;
  version(assetId: string): number | undefined;
  retain(): MediaLease;
  retainAsset(assetId: string): Pick<MediaLease, 'resolve' | 'ready' | 'release'>;
  release(): void;
}

interface PinnedAsset {
  resolve: RasterImageResolver & { release(): void };
  source: Promise<CanvasImageSource | null>;
  references: number;
}

interface SharedRasterImages {
  media: ReadonlyMap<string, { bytes: Uint8Array; version: number }>;
  assets: ReadonlyMap<string, PinnedAsset>;
  disposed(): boolean;
  releaseAsset(key: string, asset: PinnedAsset): void;
}

const sharedRasterImages = new WeakMap<object, SharedRasterImages>();

export function createSlideRasterImages<C extends SlideRasterCanvas, L extends RasterDisplayList>(options: {
  images: () => RasterImageFactory;
  frameAssetIds(list: L): Iterable<string>;
  reuse?: object;
}) {
  const images = options.images();
  const reused = options.reuse ? sharedRasterImages.get(options.reuse) : undefined;
  const pool: RasterPool<C> = { canvases: [], disposed: false };
  const media = new Map<string, { bytes: Uint8Array; version: number }>();
  const assets = new Map<string, PinnedAsset>();
  const frames = new Map<RasterFrame<L>, { resolve: RasterImageResolver;
    versions: Map<string, number>; retainAsset: MediaLease['retainAsset']; release(): void;
    ready: Promise<void>; references: number }>();
  let disposed = false;
  let version = 0;
  const finish = () => { if (disposed && !frames.size) media.clear(); };
  const releaseAsset = (key: string, asset: PinnedAsset) => {
    if (--asset.references === 0) { assets.delete(key); asset.resolve.release(); }
  };
  const pinnedElsewhere = (id: string, bytes: Uint8Array) => {
    if (!reused || reused.disposed()) return undefined;
    const value = reused.media.get(id);
    if (!value) return undefined;
    const key = JSON.stringify([id, value.version]);
    const asset = reused.assets.get(key);
    if (!asset || !(value.bytes.buffer === bytes.buffer || (value.bytes.length === bytes.length &&
      value.bytes.every((byte, index) => byte === bytes[index])))) return undefined;
    return { asset, release: () => reused.releaseAsset(key, asset) };
  };
  const ingest = (supplied: readonly RasterMedia[]) => {
    for (const asset of supplied) {
      const previous = media.get(asset.assetId);
      if (previous?.bytes.buffer === asset.bytes) continue;
      const bytes = new Uint8Array(asset.bytes);
      if (!previous || previous.bytes.length !== bytes.length ||
        !previous.bytes.every((byte, index) => byte === bytes[index])) {
        media.set(asset.assetId, { bytes, version: ++version });
      } else previous.bytes = bytes;
    }
  };
  const result = {
    ingest,
    pin(frame: RasterFrame<L>, supplied: RasterMedia[] = []): MediaLease {
      if (disposed) throw new Error('Slide images are disposed');
      let entry = frames.get(frame);
      ingest(entry ? supplied : [...supplied, ...frame.frame.media]);
      if (!entry) {
        const heldAssets = new Map<string, { asset: PinnedAsset; release(): void }>();
        const versions = new Map<string, number>();
        for (const id of options.frameAssetIds(frame.displayList)) {
          const value = media.get(id);
          if (!value) continue;
          const key = JSON.stringify([id, value.version]);
          let asset = assets.get(key);
          const shared = asset ? undefined : pinnedElsewhere(id, value.bytes);
          if (shared) asset = shared.asset;
          else if (!asset) {
            const resolve = images.resolve({ media: new Map([[key, value.bytes]]) });
            asset = { resolve, source: Promise.resolve(resolve(key)), references: 0 };
            assets.set(key, asset);
          }
          const pinned = asset;
          pinned.references += 1;
          heldAssets.set(id, { asset: pinned, release: shared?.release ?? (() => releaseAsset(key, pinned)) });
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
                value?.release();
              } };
          },
          release() {
            for (const { release } of heldAssets.values()) release();
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
    trim() {
      for (const entry of pool.canvases) zero(entry.canvas);
      pool.canvases.length = 0;
      images.trim?.();
    },
    dispose() {
      if (disposed) return;
      disposed = pool.disposed = true;
      for (const entry of pool.canvases) zero(entry.canvas);
      pool.canvases.length = 0;
      images.dispose(); finish();
    },
  };
  rasterPools.set(result, pool);
  sharedRasterImages.set(result, { media, assets, disposed: () => disposed, releaseAsset });
  return result;
}

const tiles = new WeakMap<object, Map<string, RasterTile<SlideRasterCanvas>>>();
const rasterPools = new WeakMap<object, RasterPool<SlideRasterCanvas>>();
let tileSequence = 0;

export function slideRasterCache<C extends SlideRasterCanvas>(images: object) {
  let byKey = tiles.get(images);
  if (!byKey) { byKey = new Map(); tiles.set(images, byKey); }
  return { byKey: byKey as Map<string, RasterTile<C>>, pool: rasterPools.get(images)! as RasterPool<C> };
}

export function zero(canvas: SlideRasterCanvas): void { canvas.width = 0; canvas.height = 0; }

function guardedContext(ctx: CanvasRenderingContext2D, current: () => boolean,
  transform?: () => void): CanvasRenderingContext2D {
  const methods = new Map<PropertyKey, (...args: unknown[]) => unknown>();
  return new Proxy(ctx, {
    get(target, property) {
      const cached = methods.get(property);
      if (cached) return cached;
      const value: unknown = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      const method = transform && property === 'clearRect' ? () => {}
        : transform && property === 'setTransform' ? () => { if (current()) transform(); }
        : (...args: unknown[]) => current() ? value.apply(target, args) : undefined;
      methods.set(property, method);
      return method;
    },
    set: (target, property, value: unknown) => !current() || Reflect.set(target, property, value, target),
  });
}

interface RasterLayerSource {
  primitive: RasterPrimitive;
  parents: readonly RasterPrimitive[];
  parentsKey: string;
  box?: RasterTextBox;
}

function rasterLayers(primitives: readonly RasterPrimitive[], parents: readonly RasterPrimitive[] = [],
  layers: RasterLayerSource[] = []): RasterLayerSource[] {
  const parentsKey = JSON.stringify(parents.map((parent) => ({
    x: parent.x, y: parent.y, w: parent.w, h: parent.h, transform: parent.transform,
  })));
  for (const primitive of primitives) {
    if (primitive.kind === 'chart' || primitive.kind === 'table') rasterLayers(primitive.primitives, [...parents, primitive], layers);
    else if (primitive.kind === 'textBox') {
      for (const line of primitive.lines) layers.push({ primitive: { ...primitive, lines: [line] }, parents, parentsKey, box: primitive });
    } else layers.push({ primitive, parents, parentsKey });
  }
  return layers;
}

/**
 * Paints a slide into `ctx` as {@link buildRaster} composites it, each layer clipped to its tile's device pixels,
 * without allocating tile canvases. The result is not retained, so later damage patches still need a built raster.
 */
export async function paintRasterDirect<L extends RasterDisplayList>(
  ctx: CanvasRenderingContext2D, list: L, dpr: number, scale: number, media: Pick<MediaLease, 'resolve' | 'ready'>,
  paint: (ctx: CanvasRenderingContext2D, list: L, dpr: number, scale: number, options?: RasterPaintOptions) => Promise<void>,
  current: () => boolean,
): Promise<void> {
  const factor = dpr * scale;
  await media.ready;
  if (!current()) return;
  ctx.save();
  try {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    await paint(guardedContext(ctx, current), { ...list, primitives: [] }, dpr, scale);
    for (const { primitive, parents } of rasterLayers(list.primitives)) {
      if (!current()) return;
      const bounds = layerBounds(primitive, parents, factor, list);
      if (!bounds) continue;
      const x = Math.round(bounds.x * factor); const y = Math.round(bounds.y * factor);
      const width = Math.ceil(bounds.width * factor); const height = Math.ceil(bounds.height * factor);
      ctx.save();
      try {
        const layer = guardedContext(ctx, current, () => {
          ctx.setTransform(1, 0, 0, 1, 0, 0);
          ctx.beginPath(); ctx.rect(x, y, width, height); ctx.clip();
          ctx.setTransform(factor, 0, 0, factor, 0, 0);
          for (const parent of parents) {
            applyPrimitiveTransform(ctx, parent);
            ctx.beginPath(); ctx.rect(parent.x, parent.y, parent.w, parent.h); ctx.clip();
          }
        });
        await paint(layer, { ...list, background: undefined, primitives: [primitive] }, dpr, scale,
          { resolveImage: media.resolve });
      } finally { ctx.restore(); }
    }
  } finally { ctx.restore(); }
}

export function buildRaster<C extends SlideRasterCanvas, L extends RasterDisplayList>(
  factory: (width: number, height: number) => C, list: L, dpr: number, scale: number, media: MediaLease,
  byKey: Map<string, RasterTile<C>>, pool: RasterPool<C>, jobs: Promise<void>[], remove: () => void,
  options: {
    paintSlide(ctx: CanvasRenderingContext2D, list: L, dpr: number, scale: number, options?: RasterPaintOptions): Promise<void>;
    sizeCanvasForSlide(canvas: HTMLCanvasElement, list: L, dpr: number, scale: number): void;
  },
): RasterEntry<C> {
  const paint = options.paintSlide;
  const factor = dpr * scale;
  const width = Math.ceil(list.width * scale * dpr); const height = Math.ceil(list.height * scale * dpr);
  const index = pool.canvases.findIndex(({ canvas }) => canvas.width === width && canvas.height === height);
  const previous = pool.canvases.splice(index < 0 ? pool.canvases.length - 1 : index, 1)[0];
  const authoritative = previous?.canvas ?? factory(width, height);
  const sized = authoritative.width === width && authoritative.height === height;
  if (!previous || !sized) {
    if ('style' in authoritative) options.sizeCanvasForSlide(authoritative, list, dpr, scale);
    else { authoritative.width = width; authoritative.height = height; }
  }
  const background = JSON.stringify([list.width, list.height, factor, list.background]);
  const ctx = authoritative.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) { zero(authoritative); throw new Error('Slide canvas is unavailable'); }
  const ordered: RasterLayer<C>[] = [];
  const held = new Map<RasterTile<C>, string>();
  const listeners = new Set<() => void>();
  let cancelled = false;
  let ready = false;
  let damageRect: RasterRect | undefined;
  let releasedMedia = false;
  const releaseMedia = () => { if (!releasedMedia) { releasedMedia = true; media.release(); } };
  const current = () => !cancelled;
  let recycled = false;
  const recycle = () => {
    if (recycled || !ready || cancelled) return;
    recycled = true;
    if (!pool.disposed && pool.canvases.length < 2) pool.canvases.push({ canvas: authoritative, background,
      layers: ordered.map(({ tile, bounds }) => ({ id: tile.id, bounds })) });
    else zero(authoritative);
  };
  const release = () => {
    if (cancelled) return;
    cancelled = true;
    remove();
    listeners.clear();
    if (!recycled && ready && !pool.disposed && pool.canvases.length < 2) pool.canvases.push({ canvas: authoritative, background,
      layers: ordered.map(({ tile, bounds }) => ({ id: tile.id, bounds })) });
    else if (!recycled) zero(authoritative);
    for (const [tile, key] of held) if (--tile.references === 0) {
      tile.cancelled = true;
      tile.releaseMedia?.();
      if (byKey.get(key) === tile) byKey.delete(key);
      zero(tile.canvas);
    }
    held.clear();
    releaseMedia();
  };

  function add(primitive: RasterPrimitive, parents: readonly RasterPrimitive[], parentsKey: string, box?: RasterTextBox): void {
    const bounds = layerBounds(primitive, parents, factor, list);
    if (!bounds) return;
    const content = primitive.kind === 'textBox' ? { ...primitive, paragraphs: [] } : primitive;
    const key = `[${JSON.stringify(content)},${parentsKey},${JSON.stringify(factor)},${JSON.stringify(bounds)},${
      JSON.stringify(primitive.kind === 'image' && primitive.assetId ? media.version(primitive.assetId) ?? null : null)}]`;
    let tile = byKey.get(key);
    if (!tile) {
      const canvas = factory(Math.ceil(bounds.width * factor), Math.ceil(bounds.height * factor));
      canvas.width = Math.ceil(bounds.width * factor);
      canvas.height = Math.ceil(bounds.height * factor);
      const context = canvas.getContext('2d') as CanvasRenderingContext2D | null;
      if (!context) { zero(canvas); throw new Error('Slide canvas is unavailable'); }
      const imageMedia = primitive.kind === 'image' && primitive.assetId ? media.retainAsset(primitive.assetId) : undefined;
      let painting = false;
      const item: RasterTile<C> = { id: ++tileSequence, canvas, references: 0, cancelled: false, ready: Promise.resolve(),
        releaseMedia() { if (!painting) imageMedia?.release(); } };
      const proxy = guardedContext(context, () => !item.cancelled, () => {
        context.setTransform(factor, 0, 0, factor, -bounds.x * factor, -bounds.y * factor);
        for (const parent of parents) {
          applyPrimitiveTransform(context, parent);
          context.beginPath(); context.rect(parent.x, parent.y, parent.w, parent.h); context.clip();
        }
      });
      const paintTile = () => {
        if (item.cancelled) return;
        painting = true;
        return paint(proxy, { ...list, background: undefined, primitives: [primitive] }, dpr, scale,
          { resolveImage: (assetId) => (imageMedia ?? media).resolve(assetId) });
      };
      try {
        item.ready = imageMedia ? imageMedia.ready.then(paintTile).finally(() => imageMedia.release()) : Promise.resolve(paintTile());
      } catch (error) { imageMedia?.release(); item.ready = Promise.reject(error); }
      void item.ready.catch(() => {});
      tile = item;
      byKey.set(key, tile);
    }
    if (!held.has(tile)) { held.set(tile, key); tile.references += 1; }
    jobs.push(tile.ready);
    ordered.push({ tile, bounds, box, line: primitive.kind === 'textBox' ? primitive.lines[0] : undefined });
  }


  const raster: SlideRaster<C> = {
    canvas: authoritative, factor, get damage() { return damageRect; }, get ready() { return ready && current(); },
    subscribe(listener) { if (!cancelled) listeners.add(listener); return () => { listeners.delete(listener); }; },
    plate(active, bounds) {
      const background = factory(Math.ceil(bounds.width * factor), Math.ceil(bounds.height * factor));
      const foreground = factory(Math.ceil(bounds.width * factor), Math.ceil(bounds.height * factor));
      background.width = foreground.width = Math.ceil(bounds.width * factor);
      background.height = foreground.height = Math.ceil(bounds.height * factor);
      try {
        const back = background.getContext('2d') as CanvasRenderingContext2D | null;
        const front = foreground.getContext('2d') as CanvasRenderingContext2D | null;
        if (ready && current() && back && front) {
          const proxy = guardedContext(back, current, () => {
            back.setTransform(factor, 0, 0, factor, -bounds.x * factor, -bounds.y * factor);
          });
          void paint(proxy, { ...list, primitives: [] }, dpr, scale).catch(() => {});
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
  try {
    for (const layer of rasterLayers(list.primitives)) add(layer.primitive, layer.parents, layer.parentsKey, layer.box);
  } catch (error) { release(); throw error; }
  const completion = Promise.allSettled(jobs).then(async (results) => {
    for (const result of results) if (result.status === 'rejected') throw result.reason;
    if (!current()) return;
    const incremental = previous && sized && previous.background === background;
    const damage = incremental ? rasterDamage(previous.layers, ordered, factor) : undefined;
    damageRect = incremental ? damage : { x: 0, y: 0, width, height };
    if (!incremental || damage) {
      ctx.save();
      try {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        if (damage) { ctx.beginPath(); ctx.rect(damage.x, damage.y, damage.width, damage.height); ctx.clip(); }
        if (previous) ctx.clearRect(0, 0, authoritative.width, authoritative.height);
        await paint(guardedContext(ctx, current), { ...list, primitives: [] }, dpr, scale);
        if (!current()) return;
        for (const layer of ordered) {
          if (!damage || intersectsDamage(layer.bounds, damage, factor)) {
            ctx.drawImage(layer.tile.canvas, layer.bounds.x * factor, layer.bounds.y * factor);
          }
        }
      } finally { ctx.restore(); }
    }
    ready = true;
    for (const listener of listeners) listener();
  }).finally(releaseMedia);
  return { raster, ready: completion, references: 0, cancelled: false, recycle,
    release() { this.cancelled = true; release(); } };
}

export function rasterDamage<C extends SlideRasterCanvas>(previous: readonly PooledLayer[], next: readonly RasterLayer<C>[], factor: number): RasterRect | undefined {
  let x = Infinity; let y = Infinity; let right = -Infinity; let bottom = -Infinity;
  const include = (layer: { bounds: RasterRect } | undefined) => {
    if (!layer) return;
    const bounds = layer.bounds;
    x = Math.min(x, bounds.x); y = Math.min(y, bounds.y);
    right = Math.max(right, bounds.x + bounds.width); bottom = Math.max(bottom, bounds.y + bounds.height);
  };
  for (let index = 0; index < Math.max(previous.length, next.length); index += 1) {
    const a = previous[index]; const b = next[index];
    if (a?.id === b?.tile.id && a?.bounds.x === b?.bounds.x && a?.bounds.y === b?.bounds.y &&
      a?.bounds.width === b?.bounds.width && a?.bounds.height === b?.bounds.height) continue;
    include(a); include(b);
  }
  if (x === Infinity) return undefined;
  x = Math.floor(x * factor); y = Math.floor(y * factor);
  return { x, y, width: Math.ceil(right * factor) - x, height: Math.ceil(bottom * factor) - y };
}

function intersectsDamage(bounds: RasterRect, damage: RasterRect, factor: number): boolean {
  return bounds.x * factor < damage.x + damage.width && (bounds.x + bounds.width) * factor > damage.x &&
    bounds.y * factor < damage.y + damage.height && (bounds.y + bounds.height) * factor > damage.y;
}

export function applyPrimitiveTransform(ctx: CanvasRenderingContext2D, primitive: RasterPrimitive): void {
  if (!primitive.transform) return;
  const cx = primitive.x + primitive.w / 2;
  const cy = primitive.y + primitive.h / 2;
  ctx.translate(cx, cy);
  ctx.rotate((primitive.transform.rotationDeg ?? 0) * Math.PI / 180);
  ctx.scale(primitive.transform.flipH ? -1 : 1, primitive.transform.flipV ? -1 : 1);
  ctx.translate(-cx, -cy);
}

function layerBounds(
  primitive: RasterPrimitive, parents: readonly RasterPrimitive[], factor: number, list: RasterDisplayList,
): RasterRect | null {
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
