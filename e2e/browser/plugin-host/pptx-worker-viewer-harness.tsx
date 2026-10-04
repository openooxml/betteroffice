import { createRoot } from 'react-dom/client';
import * as React from 'react';
import { PptxEditor, type PptxEditorApi, type PptxWorkerViewerApi } from '@betteroffice/pptx-react';
import demoUrl from '../../../apps/demo/public/betteroffice-demo.pptx?url';
import tiffUrl from '../../../packages/pptx/src/render/fixtures/tiff-image.pptx?url';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';

export type ViewerArm = 'in-thread' | 'worker';

interface ImageRegion { left: number; top: number; right: number; bottom: number }

interface CanvasPaint {
  serial: number;
  depth: number;
  finishedAt: number;
  firstFinishedAt: number;
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
    const paint = { serial: 0, depth: 0, finishedAt: 0, firstFinishedAt: 0, regions: [] as ImageRegion[] };
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
        if (paint) { paint.depth = 0; paint.finishedAt = 0; }
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
  // paintSlide's outer restore runs after all awaited primitives and image decodes.
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
      Reflect.apply(nativeDraw, this, args);
      const source = args[0];
      const isImage = source instanceof ImageBitmap || source instanceof HTMLImageElement ||
        (typeof source === 'object' && source !== null && imageCanvases.has(source));
      if (!isImage) return;
      imageCanvases.add(this.canvas);
      const paint = this.canvas instanceof HTMLCanvasElement ? state(this.canvas) : undefined;
      if (!paint || paint.depth === 0) return;
      const offset = args.length === 9 ? 5 : 1;
      const x = Number(args[offset]);
      const y = Number(args[offset + 1]);
      const size = source as { width: number; height: number };
      const width = args.length === 3 ? size.width : Number(args[offset + 2]);
      const height = args.length === 3 ? size.height : Number(args[offset + 3]);
      const transform = this.getTransform();
      const corners = [[x, y], [x + width, y], [x, y + height], [x + width, y + height]]
        .map(([cx, cy]) => new DOMPoint(cx, cy).matrixTransform(transform));
      paint.regions.push({
        left: Math.min(...corners.map((point) => point.x)),
        top: Math.min(...corners.map((point) => point.y)),
        right: Math.max(...corners.map((point) => point.x)),
        bottom: Math.max(...corners.map((point) => point.y)),
      });
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
  size?: { width: number; height: number }
) {
  return until(() => {
    const target = canvas();
    const paint = probe.get(target);
    if (!target || !paint || paint.serial <= after || !paint.finishedAt || paint.depth !== 0) return;
    if (size && (target.width !== size.width || target.height !== size.height)) return;
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
  return { file, fonts: [{ family: 'Liberation Sans', bytes: font }] };
}

export interface PixelComparison {
  slide: number;
  zoom: number;
  dpr: number;
  thumbnail: boolean;
  differingPixels: number;
  maxChannelDelta: number;
  localImageDraws: number;
  workerImageDraws: number;
  localImagePixels: number;
  workerImagePixels: number;
}

function imagePixels(canvas: HTMLCanvasElement, regions: ImageRegion[]) {
  const ctx = canvas.getContext('2d')!;
  const background = ctx.getImageData(0, 0, 1, 1).data;
  let count = 0;
  for (const region of regions) {
    const x = Math.max(0, Math.ceil(region.left));
    const y = Math.max(0, Math.ceil(region.top));
    const width = Math.min(canvas.width, Math.floor(region.right)) - x;
    const height = Math.min(canvas.height, Math.floor(region.bottom)) - y;
    if (width <= 0 || height <= 0) continue;
    const pixels = ctx.getImageData(x, y, width, height).data;
    for (let offset = 0; offset < pixels.length; offset += 4) {
      if (pixels[offset + 3] > 0 && [0, 1, 2, 3].some((channel) =>
        pixels[offset + channel] !== background[channel])) count += 1;
    }
  }
  return count;
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
    localImagePixels: imagePixels(a, local.paint.regions),
    workerImagePixels: imagePixels(b, worker.paint.regions),
  };
}

export interface WorkerViewerProbe {
  ready: Promise<{ slideCount: number; initialSlide: number }>;
  errors: string[];
  show(slide: number, zoom: number): Promise<PixelComparison>;
  compareCurrent(slide: number): Promise<PixelComparison>;
  thumbnails(): Promise<PixelComparison[]>;
  workerImage(): Promise<{ draws: number; pixels: number }>;
}

declare global {
  interface Window { __pptxWorkerViewer: WorkerViewerProbe }
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
      const beforeLocal = paints.get(mainCanvas('in-thread'))?.serial ?? 0;
      const beforeWorker = paints.get(mainCanvas('worker'))?.serial ?? 0;
      const results = await Promise.all([local, worker].map((api) =>
        api.commands.execute('zoom', { scale: nextZoom })));
      if (results.some((result) => !result.ok)) throw new Error('Zoom command failed');
      await Promise.all([
        waitForPaint(paints, () => mainCanvas('in-thread'), beforeLocal),
        waitForPaint(paints, () => mainCanvas('worker'), beforeWorker),
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
    const { canvas, paint } = await waitForPaint(paints, () => mainCanvas('worker'));
    return { draws: paint.regions.length, pixels: imagePixels(canvas, paint.regions) };
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
