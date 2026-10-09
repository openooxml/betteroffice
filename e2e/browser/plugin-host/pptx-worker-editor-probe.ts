import type { PptxWorkerEditorFrame } from '../../../packages/pptx/src/index';

export interface PixelResult {
  thumbnail: boolean;
  dpr: number;
  differingPixels: number;
  maxChannelDelta: number;
  sequence: number;
}

export interface WorkerEditorProbe {
  mounted: Promise<void>;
  ready: Promise<void>;
  errors: { name: string; code: string; cause: string; typed: boolean }[];
  state(): { hydrated: boolean; sequence: number; acknowledged: number; held: number };
  refuse(): string;
  releaseHydration(): Promise<void>;
  hold(): void;
  release(): Promise<void>;
  text(): string;
  select(start: number, end: number): boolean;
  point(position: number): { x: number; y: number; position: number | null };
  baseline(): Promise<void>;
  unchanged(): boolean;
  parity(): Promise<PixelResult[]>;
  overlay(position: number): Promise<{ differingPixels: number; pixels: number }>;
  provenance(): { total: number; thumbnails: number; unknown: number };
  save(): Promise<{ bytes: number; text: string }>;
  toggleReadOnly(value: boolean): Promise<{ sessions: number; ready: number }>;
  edit(text: string): Promise<string>;
  proposal(): Promise<void>;
  fail(): void;
  rejected(index: number): Promise<string[]>;
  recover(index: number): Promise<{ bytes: number; text: string; recovery: boolean }>;
  replace(): Promise<void>;
  unmount(): Promise<void>;
}

declare global {
  interface Window { __pptxWorkerEditor: WorkerEditorProbe }
}

export async function until<T>(read: () => T | undefined, label: string): Promise<T> {
  const deadline = performance.now() + 120_000;
  for (;;) {
    const value = read();
    if (value !== undefined) return value;
    if (performance.now() > deadline) throw new Error(`Timed out waiting for ${label}`);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  }
}

export function pixels(canvas: HTMLCanvasElement) {
  return canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
}

export function difference(a: Uint8ClampedArray, b: Uint8ClampedArray) {
  if (a.length !== b.length) throw new Error('Canvas sizes differ');
  let differingPixels = 0;
  let maxChannelDelta = 0;
  for (let offset = 0; offset < a.length; offset += 4) {
    let differs = false;
    for (let channel = 0; channel < 4; channel += 1) {
      const delta = Math.abs(a[offset + channel] - b[offset + channel]);
      differs ||= delta !== 0;
      maxChannelDelta = Math.max(maxChannelDelta, delta);
    }
    if (differs) differingPixels += 1;
  }
  return { differingPixels, maxChannelDelta };
}

export function installPaintProbe() {
  const paints = new WeakMap<HTMLCanvasElement, {
    depth: number; finished: boolean; frame?: PptxWorkerEditorFrame;
    provenance?: { thumbnail: boolean; worker: boolean };
  }>();
  const history: { thumbnail: boolean; worker: boolean }[] = [];
  const tagged = new WeakSet<object>();
  let active: HTMLCanvasElement | undefined;
  for (const dimension of ['width', 'height'] as const) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, dimension)!;
    Object.defineProperty(HTMLCanvasElement.prototype, dimension, {
      ...descriptor,
      set(this: HTMLCanvasElement, value: number) {
        descriptor.set!.call(this, value);
        paints.delete(this);
      },
    });
  }
  const nativeSave = CanvasRenderingContext2D.prototype.save;
  const nativeRestore = CanvasRenderingContext2D.prototype.restore;
  const nativeClear = CanvasRenderingContext2D.prototype.clearRect;
  CanvasRenderingContext2D.prototype.save = function (this: CanvasRenderingContext2D) {
    nativeSave.call(this);
    if (!this.canvas.closest('[data-editor]')) return;
    const paint = paints.get(this.canvas) ?? { depth: 0, finished: false };
    paint.depth += 1;
    paints.set(this.canvas, paint);
  };
  CanvasRenderingContext2D.prototype.clearRect = function (
    this: CanvasRenderingContext2D, ...args: Parameters<CanvasRenderingContext2D['clearRect']>
  ) {
    nativeClear.apply(this, args);
    const paint = paints.get(this.canvas);
    if (!paint || paint.depth !== 1) return;
    paint.finished = false;
    paint.frame = undefined;
    if (this.canvas.matches('[data-testid="pptx-slide-canvas"], aside canvas')) {
      paint.provenance = { thumbnail: !!this.canvas.closest('aside'), worker: false };
      history.push(paint.provenance);
    }
    active = this.canvas;
    queueMicrotask(() => { active = undefined; });
  };
  CanvasRenderingContext2D.prototype.restore = function (this: CanvasRenderingContext2D) {
    nativeRestore.call(this);
    const paint = paints.get(this.canvas);
    if (!paint || --paint.depth !== 0) return;
    if (!this.canvas.matches('[data-testid="pptx-slide-canvas"], aside canvas')) return;
    paint.finished = true;
  };
  return {
    paints, history,
    tag(frame: PptxWorkerEditorFrame) {
      if (tagged.has(frame.displayList)) return;
      tagged.add(frame.displayList);
      const primitives = frame.displayList.primitives;
      Object.defineProperty(frame.displayList, 'primitives', { get() {
        const paint = active && paints.get(active);
        if (paint) {
          paint.frame = frame;
          if (paint.provenance) paint.provenance.worker = true;
        }
        return primitives;
      } });
    },
  };
}
