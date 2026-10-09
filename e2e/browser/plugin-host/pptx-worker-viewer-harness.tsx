import { createRoot } from 'react-dom/client';
import * as React from 'react';
import JSZip from 'jszip';
import { PptxEditor, type PptxEditorApi, type PptxWorkerViewerApi } from '@betteroffice/pptx-react';
import demoUrl from '../../../apps/demo/public/betteroffice-demo.pptx?url';
import tiffUrl from '../../../packages/pptx/src/render/fixtures/tiff-image.pptx?url';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import type { PixelComparison, ViewerArm, WorkerViewerProbe } from './pptx-worker-viewer-probe';

interface ImageRegion { left: number; top: number; right: number; bottom: number; pixels: number }

interface CanvasPaint {
  serial: number;
  depth: number;
  finishedAt: number;
  firstFinishedAt: number;
  scale: number;
  regions: ImageRegion[];
}

export function installPaintProbe(captureImages = false) {
  const paints = new WeakMap<HTMLCanvasElement, CanvasPaint>();
  const imageCanvases = new WeakSet<object>();
  let serial = 0;
  const state = (canvas: HTMLCanvasElement) => {
    const existing = paints.get(canvas);
    if (existing) return existing;
    if (!canvas.isConnected || !canvas.closest('[data-arm]')) return;
    const paint = { serial: 0, depth: 0, finishedAt: 0, firstFinishedAt: 0, scale: 0, regions: [] as ImageRegion[] };
    paints.set(canvas, paint);
    return paint;
  };
  for (const dimension of ['width', 'height'] as const) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, dimension)!;
    Object.defineProperty(HTMLCanvasElement.prototype, dimension, {
      ...descriptor,
      set(this: HTMLCanvasElement, value: number) {
        descriptor.set!.call(this, value);
        const paint = state(this);
        if (paint) { paint.depth = 0; paint.finishedAt = 0; paint.scale = 0; }
      },
    });
  }
  const nativeSave = CanvasRenderingContext2D.prototype.save;
  const nativeRestore = CanvasRenderingContext2D.prototype.restore;
  CanvasRenderingContext2D.prototype.save = function (this: CanvasRenderingContext2D) {
    nativeSave.call(this);
    const paint = state(this.canvas);
    if (!paint) return;
    if (paint.depth++ === 0) {
      paint.serial = ++serial;
      paint.finishedAt = 0;
      paint.regions = [];
    }
  };
  const nativeSetTransform = CanvasRenderingContext2D.prototype.setTransform;
  CanvasRenderingContext2D.prototype.setTransform = function (
    this: CanvasRenderingContext2D, ...args: unknown[]
  ) {
    Reflect.apply(nativeSetTransform, this, args);
    const paint = state(this.canvas);
    if (paint?.depth === 1) paint.scale = this.getTransform().a;
  } as CanvasRenderingContext2D['setTransform'];
  CanvasRenderingContext2D.prototype.restore = function (this: CanvasRenderingContext2D) {
    nativeRestore.call(this);
    const paint = state(this.canvas);
    if (paint && paint.depth > 0 && --paint.depth === 0) {
      paint.finishedAt = performance.now();
      paint.firstFinishedAt ||= paint.finishedAt;
    }
  };
  if (captureImages) {
    const countedDraw = (nativeDraw: CanvasRenderingContext2D['drawImage']) => function (
      this: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, ...args: unknown[]
    ) {
      const source = args[0];
      const isImage = source instanceof ImageBitmap || source instanceof HTMLImageElement ||
        (typeof source === 'object' && source !== null && imageCanvases.has(source));
      const paint = this.canvas instanceof HTMLCanvasElement ? state(this.canvas) : undefined;
      if (!isImage || !paint || paint.depth === 0) {
        Reflect.apply(nativeDraw, this, args);
        if (isImage) imageCanvases.add(this.canvas);
        return;
      }
      const offset = args.length === 9 ? 5 : 1;
      const x = Number(args[offset]);
      const y = Number(args[offset + 1]);
      const size = source as { width: number; height: number };
      const width = args.length === 3 ? size.width : Number(args[offset + 2]);
      const height = args.length === 3 ? size.height : Number(args[offset + 3]);
      const transform = this.getTransform();
      const corners = [[x, y], [x + width, y], [x, y + height], [x + width, y + height]]
        .map(([cx, cy]) => new DOMPoint(cx, cy).matrixTransform(transform));
      const region = {
        left: Math.min(...corners.map((point) => point.x)),
        top: Math.min(...corners.map((point) => point.y)),
        right: Math.max(...corners.map((point) => point.x)),
        bottom: Math.max(...corners.map((point) => point.y)),
        pixels: 0,
      };
      const left = Math.max(0, Math.floor(region.left));
      const top = Math.max(0, Math.floor(region.top));
      const clippedWidth = Math.min(this.canvas.width, Math.ceil(region.right)) - left;
      const clippedHeight = Math.min(this.canvas.height, Math.ceil(region.bottom)) - top;
      const before = clippedWidth > 0 && clippedHeight > 0 ?
        this.getImageData(left, top, clippedWidth, clippedHeight).data : null;
      Reflect.apply(nativeDraw, this, args);
      imageCanvases.add(this.canvas);
      if (before) {
        const after = this.getImageData(left, top, clippedWidth, clippedHeight).data;
        for (let offset = 0; offset < after.length; offset += 4) {
          if (after[offset + 3] > 0 && [0, 1, 2, 3].some((channel) =>
            after[offset + channel] !== before[offset + channel])) region.pixels += 1;
        }
      }
      paint.regions.push(region);
    };
    CanvasRenderingContext2D.prototype.drawImage = countedDraw(
      CanvasRenderingContext2D.prototype.drawImage
    ) as CanvasRenderingContext2D['drawImage'];
    if (typeof OffscreenCanvasRenderingContext2D !== 'undefined') {
      OffscreenCanvasRenderingContext2D.prototype.drawImage = countedDraw(
        OffscreenCanvasRenderingContext2D.prototype.drawImage
      ) as OffscreenCanvasRenderingContext2D['drawImage'];
    }
  }
  return { get: (canvas: HTMLCanvasElement | null) => canvas ? paints.get(canvas) : undefined };
}

export async function until<T>(read: () => T | null | undefined, label: string): Promise<T> {
  const deadline = performance.now() + 120_000;
  for (;;) {
    const value = read();
    if (value != null) return value;
    if (performance.now() > deadline) throw new Error(`Timed out waiting for ${label}`);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  }
}

export function mainCanvas(arm: ViewerArm) {
  return Array.from(document.querySelectorAll<HTMLCanvasElement>(`[data-arm="${arm}"] canvas`))
    .find((canvas) => !canvas.closest('aside')) ?? null;
}

export async function waitForPaint(
  probe: ReturnType<typeof installPaintProbe>,
  canvas: () => HTMLCanvasElement | null,
  after = 0,
  size?: { width: number; height: number; scale?: number }
) {
  return until(() => {
    const target = canvas();
    const paint = probe.get(target);
    if (!target || !paint || paint.serial <= after || !paint.finishedAt || paint.depth !== 0) return;
    if (size && (target.width !== size.width || target.height !== size.height)) return;
    if (size?.scale !== undefined && paint.scale !== size.scale) return;
    return { canvas: target, paint };
  }, 'completed canvas paint');
}

export async function loadViewerInputs(fileUrl = demoUrl) {
  const [file, font] = await Promise.all([fileUrl, fontUrl].map(async (url) => {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Fetch failed (${response.status}): ${url}`);
    return new Uint8Array(await response.arrayBuffer());
  }));
  await document.fonts.ready;
  return { file: fileUrl === demoUrl ? await visibleImageFixture(file) : file,
    fonts: [{ family: 'Liberation Sans', bytes: font }] };
}

async function visibleImageFixture(file: Uint8Array) {
  const zip = await JSZip.loadAsync(file);
  const slidePath = 'ppt/slides/slide1.xml';
  const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
  const slide = parse(await zip.file(slidePath)!.async('string'));
  const picture = slide.getElementsByTagNameNS('*', 'pic')[0];
  const transform = picture.getElementsByTagNameNS('*', 'xfrm')[0];
  const offset = transform.getElementsByTagNameNS('*', 'off')[0];
  const extent = transform.getElementsByTagNameNS('*', 'ext')[0];
  offset.setAttribute('x', String(80 * 9525));
  offset.setAttribute('y', String(640 * 9525));
  extent.setAttribute('cx', String(240 * 9525));
  extent.setAttribute('cy', String(40 * 9525));
  const id = picture.getElementsByTagNameNS('*', 'blip')[0].getAttributeNS(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
  const rels = parse(await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'));
  const relationship = Array.from(rels.getElementsByTagNameNS('*', 'Relationship'))
    .find((rel) => rel.getAttribute('Id') === id)!;
  const mediaPath = relationship.getAttribute('Target')!.replace(/^\.\.\//, 'ppt/');
  const canvas = document.createElement('canvas');
  canvas.width = 64;
  canvas.height = 32;
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#f97316';
  ctx.fillRect(0, 0, 32, 32);
  ctx.fillStyle = '#16a34a';
  ctx.fillRect(32, 0, 32, 32);
  const image = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => {
    if (blob) resolve(blob);
    else reject(new Error('Image fixture encoding failed'));
  }, 'image/png'));
  zip.file(mediaPath, await image.arrayBuffer());
  zip.file(slidePath, new XMLSerializer().serializeToString(slide));
  return zip.generateAsync({ type: 'uint8array' });
}

function imagePixels(regions: ImageRegion[]) {
  return regions.reduce((count, region) => count + region.pixels, 0);
}

function comparePixels(
  local: Awaited<ReturnType<typeof waitForPaint>>,
  worker: Awaited<ReturnType<typeof waitForPaint>>,
  slide: number,
  zoom: number,
  thumbnail = false
): PixelComparison {
  const a = local.canvas;
  const b = worker.canvas;
  if (a.width !== b.width || a.height !== b.height) throw new Error('Canvas sizes differ');
  const left = a.getContext('2d')!.getImageData(0, 0, a.width, a.height).data;
  const right = b.getContext('2d')!.getImageData(0, 0, b.width, b.height).data;
  let differingPixels = 0;
  let maxChannelDelta = 0;
  for (let offset = 0; offset < left.length; offset += 4) {
    let differs = false;
    for (let channel = 0; channel < 4; channel += 1) {
      const delta = Math.abs(left[offset + channel] - right[offset + channel]);
      differs ||= delta !== 0;
      maxChannelDelta = Math.max(maxChannelDelta, delta);
    }
    if (differs) differingPixels += 1;
  }
  return {
    slide, zoom, dpr: devicePixelRatio, thumbnail, differingPixels, maxChannelDelta,
    localImageDraws: local.paint.regions.length, workerImageDraws: worker.paint.regions.length,
    localImagePixels: imagePixels(local.paint.regions),
    workerImagePixels: imagePixels(worker.paint.regions),
  };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: Error) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

async function mount(root: HTMLElement, probe: WorkerViewerProbe) {
  const params = new URLSearchParams(location.search);
  const workerOnly = params.get('arm') === 'worker';
  const initialSlide = Number(params.get('initialSlide') ?? 1);
  const inputs = await loadViewerInputs(params.get('fixture') === 'tiff' ? tiffUrl : demoUrl);
  const paints = installPaintProbe(true);
  const localReady = deferred<PptxEditorApi>();
  const workerReady = deferred<PptxWorkerViewerApi>();
  const onError = (error: Error) => {
    probe.errors.push(error.message);
    if (!workerOnly) localReady.reject(error);
    workerReady.reject(error);
  };
  const common = { ...inputs, readOnly: true as const, initialSlide, showToolbar: false, onError };
  createRoot(root).render(<React.Fragment>
    {!workerOnly && <div data-arm="in-thread">
      <PptxEditor {...common} onReady={localReady.resolve} />
    </div>}
    <div data-arm="worker">
      <PptxEditor {...common} experimentalWorkerOpen onReady={workerReady.resolve} />
    </div>
  </React.Fragment>);
  const [local, worker] = await Promise.all([
    workerOnly ? Promise.resolve(null) : localReady.promise, workerReady.promise,
  ]);
  if (!await worker.goToSlideAsync(initialSlide)) throw new Error('Initial worker navigation failed');
  await waitForPaint(paints, () => mainCanvas('worker'));
  if (local) await waitForPaint(paints, () => mainCanvas('in-thread'));
  const slideCount = document.querySelectorAll('[data-arm="worker"] aside button').length;
  if (local && local.handle.snapshot().slides.length !== slideCount) throw new Error('Slide counts differ');
  let localSlide = initialSlide;
  let zoom: number | null = null;

  const current = async (slide: number) => {
    if (!local || zoom === null) throw new Error('Parity requires two editors at an explicit zoom');
    await until(() => {
      const selected = document.querySelector('[data-arm="worker"] aside [aria-current="page"]');
      return selected?.getAttribute('data-slide-index') === String(slide - 1) ? true : null;
    }, `worker navigation to slide ${slide}`);
    if (!await worker.goToSlideAsync(slide)) throw new Error('Worker paint did not complete');
    const before = paints.get(mainCanvas('in-thread'))?.serial ?? 0;
    if (localSlide !== slide && !local.goToSlide(slide)) throw new Error('Local navigation failed');
    const frame = local.handle.layoutSlide(slide - 1);
    const size = {
      width: Math.ceil(frame.width * zoom * devicePixelRatio),
      height: Math.ceil(frame.height * zoom * devicePixelRatio),
    };
    const [a, b] = await Promise.all([
      waitForPaint(paints, () => mainCanvas('in-thread'), localSlide === slide ? 0 : before, size),
      waitForPaint(paints, () => mainCanvas('worker'), 0, size),
    ]);
    localSlide = slide;
    return comparePixels(a, b, slide, zoom);
  };
  probe.compareCurrent = current;
  probe.show = async (slide, nextZoom) => {
    if (!local) throw new Error('Missing local editor');
    if (zoom !== nextZoom) {
      const frame = local.handle.layoutSlide(localSlide - 1);
      const size = {
        width: Math.ceil(frame.width * nextZoom * devicePixelRatio),
        height: Math.ceil(frame.height * nextZoom * devicePixelRatio),
        scale: nextZoom * devicePixelRatio,
      };
      const after = (arm: ViewerArm) => {
        const paint = paints.get(mainCanvas(arm));
        return paint?.scale === size.scale ? 0 : paint?.serial ?? 0;
      };
      const beforeLocal = after('in-thread');
      const beforeWorker = after('worker');
      const results = await Promise.all([local, worker].map((api) =>
        api.commands.execute('zoom', { scale: nextZoom })));
      if (results.some((result) => !result.ok)) throw new Error('Zoom command failed');
      await Promise.all([
        waitForPaint(paints, () => mainCanvas('in-thread'), beforeLocal, size),
        waitForPaint(paints, () => mainCanvas('worker'), beforeWorker, size),
      ]);
      zoom = nextZoom;
    }
    if (!await worker.goToSlideAsync(slide)) throw new Error('Worker navigation failed');
    return current(slide);
  };
  probe.thumbnails = async () => {
    if (!local) throw new Error('Missing local editor');
    for (const rail of document.querySelectorAll('aside')) rail.scrollTop = 0;
    const results: PixelComparison[] = [];
    for (let index = 0; index < Math.min(3, slideCount); index += 1) {
      const canvas = (arm: ViewerArm) => {
        const rail = document.querySelector(`[data-arm="${arm}"] aside`)!;
        const row = rail.querySelectorAll('button')[index];
        const bounds = row.getBoundingClientRect();
        const visible = rail.getBoundingClientRect();
        if (bounds.top < visible.top || bounds.bottom > visible.bottom) return null;
        return row.querySelector('canvas');
      };
      const [a, b] = await Promise.all([
        waitForPaint(paints, () => canvas('in-thread')),
        waitForPaint(paints, () => canvas('worker')),
      ]);
      results.push(comparePixels(a, b, index + 1, 128 / local.handle.layoutSlide(index).width, true));
    }
    return results;
  };
  probe.workerImage = async () => {
    const { paint } = await waitForPaint(paints, () => mainCanvas('worker'));
    return { draws: paint.regions.length, pixels: imagePixels(paint.regions) };
  };
  return { slideCount, initialSlide };
}

const root = document.getElementById('worker-viewer-root');
if (root) {
  const probe: WorkerViewerProbe = {
    ready: Promise.resolve({ slideCount: 0, initialSlide: 1 }), errors: [],
    show: async () => { throw new Error('Not ready'); },
    compareCurrent: async () => { throw new Error('Not ready'); },
    thumbnails: async () => { throw new Error('Not ready'); },
    workerImage: async () => { throw new Error('Not ready'); },
  };
  window.__pptxWorkerViewer = probe;
  probe.ready = mount(root, probe);
}
