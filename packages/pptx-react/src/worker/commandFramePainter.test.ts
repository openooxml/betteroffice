import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import * as pptx from '@betteroffice/pptx';
import type { ShapePrimitive, SlideDisplayList, TextBoxPrimitive } from '@betteroffice/pptx';
import { commandFrameRaster, createCommandFrameImages, createCommandFramePainter,
  retainCommandFrameRaster } from './commandFramePainter';
import { createLocalEchoPainter } from './localEchoPainter';
import type { PeerlessConfirmedFrame, PeerlessRect, PeerlessViewSnapshot } from './peerlessViewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(() => { if (ownsDom) GlobalRegistrator.unregister(); });

interface Draw { source: HTMLCanvasElement; clip: PeerlessRect; x: number; y: number }
interface Surface { pixels: string[]; draws: Draw[]; clears: PeerlessRect[]; sizes: [number, number][] }
interface State { sx: number; sy: number; tx: number; ty: number; clip: PeerlessRect; fillStyle: string }
let surfaces: Map<HTMLCanvasElement, Surface>;
let restorers: (() => void)[];
let painters: ReturnType<typeof createCommandFramePainter>[];
let revision = 0;

function intersection(a: PeerlessRect, b: PeerlessRect): PeerlessRect {
  const x = Math.max(a.x, b.x); const y = Math.max(a.y, b.y);
  return { x, y, width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x),
    height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y) };
}

function surface(canvas: HTMLCanvasElement): Surface {
  let value = surfaces.get(canvas);
  if (!value) { value = { pixels: [], draws: [], clears: [], sizes: [] }; surfaces.set(canvas, value); }
  return value;
}

beforeEach(() => {
  surfaces = new Map(); restorers = []; painters = [];
  const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
  const resets = new Map<HTMLCanvasElement, () => void>();
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (
    this: HTMLCanvasElement, type: string,
  ) {
    if (type !== '2d') return null;
    let context = contexts.get(this);
    if (context) return context;
    const canvas = this; const target = surface(canvas);
    const initial = (): State => ({ sx: 1, sy: 1, tx: 0, ty: 0,
      clip: { x: 0, y: 0, width: canvas.width, height: canvas.height }, fillStyle: '#000000' });
    let state = initial(); let stack: State[] = []; let path: PeerlessRect | undefined;
    const rect = (x: number, y: number, width: number, height: number): PeerlessRect =>
      ({ x: x * state.sx + state.tx, y: y * state.sy + state.ty, width: width * state.sx, height: height * state.sy });
    const write = (bounds: PeerlessRect, sample: (x: number, y: number) => string | undefined) => {
      const clipped = intersection(intersection(bounds, state.clip), { x: 0, y: 0, width: canvas.width, height: canvas.height });
      for (let y = Math.ceil(clipped.y); y < clipped.y + clipped.height; y += 1) {
        for (let x = Math.ceil(clipped.x); x < clipped.x + clipped.width; x += 1) {
          const value = sample(x, y);
          if (value !== undefined) target.pixels[y * canvas.width + x] = value;
        }
      }
    };
    const methods: Record<string, unknown> = { canvas,
      save() { stack.push({ ...state, clip: { ...state.clip } }); },
      restore() { state = stack.pop() ?? state; },
      setTransform(sx: number, _b: number, _c: number, sy: number, tx: number, ty: number) {
        Object.assign(state, { sx, sy, tx, ty });
      },
      translate(x: number, y: number) { state.tx += x * state.sx; state.ty += y * state.sy; },
      scale(x: number, y: number) { state.sx *= x; state.sy *= y; },
      beginPath() { path = undefined; },
      rect(x: number, y: number, width: number, height: number) { path = rect(x, y, width, height); },
      moveTo(x: number, y: number) { path = rect(x, y, 0, 0); },
      lineTo(x: number, y: number) {
        const point = rect(x, y, 0, 0);
        if (path) {
          const left = Math.min(path.x, point.x); const top = Math.min(path.y, point.y);
          path = { x: left, y: top, width: Math.max(path.x + path.width, point.x) - left,
            height: Math.max(path.y + path.height, point.y) - top };
        }
      },
      closePath() {},
      clip() { if (path) state.clip = intersection(state.clip, path); },
      fill() { if (path) write(path, () => state.fillStyle); },
      fillRect(x: number, y: number, width: number, height: number) { write(rect(x, y, width, height), () => state.fillStyle); },
      clearRect(x: number, y: number, width: number, height: number) {
        const bounds = rect(x, y, width, height);
        target.clears.push(intersection(bounds, state.clip)); write(bounds, () => '');
      },
      fillText(text: string, x: number, y: number) { write(rect(x, y - 6, text.length * 4, 6), () => `${state.fillStyle}:${text}`); },
      measureText(text: string) { return { width: text.length * 4 }; },
      drawImage(source: HTMLCanvasElement, ...args: number[]) {
        const [sx, sy, sw, sh, dx, dy, dw, dh] = args.length === 8 ? args
          : [0, 0, source.width, source.height, args[0], args[1], args[2] ?? source.width, args[3] ?? source.height];
        const bounds = rect(dx, dy, dw, dh); const pixels = [...surface(source).pixels];
        target.draws.push({ source, clip: { ...state.clip }, x: bounds.x, y: bounds.y });
        write(bounds, (x, y) => pixels[Math.floor(sy + (y - bounds.y) * sh / bounds.height) * source.width +
          Math.floor(sx + (x - bounds.x) * sw / bounds.width)] || undefined);
      },
    };
    context = new Proxy(methods, {
      get: (object, key) => key === 'fillStyle' ? state.fillStyle : Reflect.get(object, key),
      set(object, key, value: unknown) {
        if (key === 'fillStyle') state.fillStyle = String(value);
        else Reflect.set(object, key, value);
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
    contexts.set(canvas, context);
    resets.set(canvas, () => { state = initial(); stack = []; path = undefined; });
    return context;
  } as typeof HTMLCanvasElement.prototype.getContext);
  restorers.push(() => getContext.mockRestore());
  for (const dimension of ['width', 'height'] as const) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, dimension)!;
    const set = descriptor.set!;
    Object.defineProperty(HTMLCanvasElement.prototype, dimension, { ...descriptor,
      set(this: HTMLCanvasElement, value: number) {
        set.call(this, value);
        const target = surface(this);
        target.sizes.push([this.width, this.height]); target.pixels = Array<string>(this.width * this.height).fill('');
        resets.get(this)?.();
      },
    });
    restorers.push(() => Object.defineProperty(HTMLCanvasElement.prototype, dimension, descriptor));
  }
});

afterEach(() => { for (const painter of painters) painter.dispose(); for (const restore of restorers.reverse()) restore(); });

function box(text = 'a', y = 12): TextBoxPrimitive {
  return { kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story', x: 10, y: 8, w: 80, h: 94,
    anchor: 'top', paragraphs: [], lines: [text, 'steady', 'last'].map((value, index) => ({
      x: 10, y: index ? 12 + index * 30 : y, width: value.length * 4, height: 10,
      baseline: (index ? 12 + index * 30 : y) + 8, start: index * 20, end: index * 20 + value.length,
      caretStops: [], runs: [{ text: value, start: index * 20, end: index * 20 + value.length,
        x: 10, width: value.length * 4, fontId: 0, fontFamily: 'Arial', fontSizePx: 8,
        bold: false, italic: false, underline: false, color: '#000000', glyphs: [] }],
    })) };
}

function shape(x = 110, y = 80, objectId = 2): ShapePrimitive {
  return { kind: 'shape', objectId, name: 'rectangle', geometry: 'rect', x, y, w: 20, h: 10,
    path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 0 }, { type: 'line', x: 1, y: 1 },
      { type: 'line', x: 0, y: 1 }, { type: 'close' }], fill: { kind: 'solid', color: '#ff0000' } };
}

function list(text = 'a', y = 12): SlideDisplayList {
  return { contractVersion: 1, width: 160, height: 120, background: { kind: 'solid', color: '#ffffff' },
    primitives: [box(text, y), shape()] };
}

function frame(displayList: SlideDisplayList): PeerlessConfirmedFrame {
  const id = ++revision;
  return { displayList, frame: { slideId: 'slide', slideIndex: 0, slideRevision: id, coversThroughSeq: id,
    documentEpoch: 1, incarnation: 'worker', epoch: 1, version: `v${id}`,
    displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer, media: [] } };
}

function view(confirmed: PeerlessConfirmedFrame): PeerlessViewSnapshot {
  return { confirmed, selection: null, focused: false, caret: null, selectionRects: [], composition: null,
    echo: null, activeLine: null, activeContextBytes: 0, pending: [], waitingInputCount: 0,
    appliedThroughSeq: 0, presentedThroughSeq: 0, thumbnails: new Map(), dirtyThumbnailIds: new Set(),
    editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: null, storedMarks: {},
      formatting: null, alignment: null, hyperlink: null, table: null, canEditText: true } } };
}

function painter(options: Parameters<typeof createCommandFramePainter>[1] = {}) {
  const canvas = document.createElement('canvas');
  const value = createCommandFramePainter(canvas, { dpr: 1, ...options });
  painters.push(value);
  return { canvas, value };
}

async function paint(target: ReturnType<typeof painter>, displayList: SlideDisplayList, factor = 1, slideId = 'slide') {
  const confirmed = frame(displayList);
  confirmed.frame.slideId = slideId;
  target.value.paintFrame(confirmed, view(confirmed));
  expect(await target.value.whenPainted()).toBe(true);
  return { confirmed, raster: commandFrameRaster(confirmed, factor)! };
}

async function expectFull(target: ReturnType<typeof painter>, displayList: SlideDisplayList, factor = 1) {
  const full = painter({ scale: factor });
  await paint(full, displayList, factor);
  expect(surface(target.canvas).pixels).toEqual(surface(full.canvas).pixels);
  full.value.dispose();
}

describe('command frame incremental raster', () => {
  test('typing alternates two full-slide canvases and draws only the changed line', async () => {
    const target = painter(); const canvases: HTMLCanvasElement[] = [];
    for (const text of ['a', 'ab', 'abc', 'abcd']) {
      const before = new Map([...surfaces].map(([canvas, value]) => [canvas, value.draws.length]));
      const { raster } = await paint(target, list(text));
      canvases.push(raster.canvas);
      const draws = surface(raster.canvas).draws.slice(before.get(raster.canvas) ?? 0);
      expect(draws).toHaveLength(canvases.length > 2 ? 1 : 4);
      if (canvases.length > 2) expect(draws[0].clip.height).toBeLessThan(120);
    }
    expect(canvases[2]).toBe(canvases[0]); expect(canvases[3]).toBe(canvases[1]);
    expect([...surfaces].filter(([canvas, value]) => canvas !== target.canvas &&
      value.sizes.some(([width, height]) => width === 160 && height === 120))).toHaveLength(2);
    expect(canvases.map((canvas) => surface(canvas).sizes.filter(([width, height]) => width === 160 && height === 120).length))
      .toEqual([1, 1, 1, 1]);
  });

  test('identical layers reuse pixels without clearing or drawing', async () => {
    const target = painter(); const previous = list();
    const { raster } = await paint(target, previous); await paint(target, previous);
    const draws = surface(raster.canvas).draws.length; const clears = surface(raster.canvas).clears.length;
    const next = await paint(target, previous);
    expect(next.raster.canvas).toBe(raster.canvas);
    expect(surface(raster.canvas).draws).toHaveLength(draws);
    expect(surface(raster.canvas).clears).toHaveLength(clears);
    await expectFull(target, previous);
  });

  test.each(['changed', 'moved', 'removed', 'added', 'overlap', 'reordered', 'transparent', 'fractional'])(
    '%s layers match a full repaint under the effective clip', async (change) => {
      const factor = change === 'fractional' ? 1.25 : 1;
      const target = painter({ scale: factor }); const previous = list();
      if (change === 'overlap' || change === 'reordered') previous.primitives.push(shape(12, 15, 3));
      if (change === 'transparent') previous.background = undefined;
      await paint(target, previous, factor); await paint(target, previous, factor);
      const next = { ...previous, primitives: [...previous.primitives] };
      if (change === 'removed') next.primitives = [previous.primitives[1]];
      else if (change === 'added') next.primitives.push(shape(50, 110, 4));
      else if (change === 'reordered') next.primitives.reverse();
      else next.primitives[0] = box('abcd', change === 'moved' ? 26 : change === 'fractional' ? 12.4 : 12);
      const first = [...surfaces].find(([canvas, value]) => canvas !== target.canvas &&
        value.sizes.some(([width, height]) => width === 160 * factor && height === 120 * factor))![0];
      const before = surface(first).draws.length;
      const { raster } = await paint(target, next, factor);
      expect(raster.canvas).toBe(first);
      if (change === 'overlap') expect(surface(first).draws.slice(before)).toHaveLength(2);
      await expectFull(target, next, factor);
    });

  test.each(['size', 'background', 'factor'])('%s changes fully repaint a pooled canvas', async (change) => {
    const images = createCommandFrameImages();
    try {
      const target = painter({ images });
      const first = await paint(target, list()); await paint(target, list('ab'));
      const next = list('abc');
      if (change === 'size') next.width = 180;
      if (change === 'background') next.background = { kind: 'solid', color: '#00ff00' };
      const consumer = change === 'factor' ? painter({ images, scale: 2 }) : target;
      const factor = change === 'factor' ? 2 : 1;
      const before = surface(first.raster.canvas).draws.length;
      const { raster } = await paint(consumer, next, factor, change === 'factor' ? 'slide' : 'other');
      expect(raster.canvas).toBe(first.raster.canvas);
      expect(surface(raster.canvas).draws.slice(before)).toHaveLength(4);
      const clears = surface(raster.canvas).clears;
      expect(clears[clears.length - 1]).toEqual({ x: 0, y: 0, width: next.width * factor, height: next.height * factor });
      await expectFull(consumer, next, factor);
    } finally { images.dispose(); }
  });

  test('owning painter disposal zeros displayed and pooled canvases', async () => {
    const target = painter(); const first = await paint(target, list()); const second = await paint(target, list('ab'));
    expect([first.raster.canvas.width, first.raster.canvas.height]).toEqual([160, 120]);
    target.value.dispose();
    for (const canvas of [first.raster.canvas, second.raster.canvas]) expect([canvas.width, canvas.height]).toEqual([0, 0]);
  });

  test('shared images keep at most two released canvases and dispose the pool', async () => {
    const images = createCommandFrameImages(); const canvases: HTMLCanvasElement[] = [];
    try {
      const targets = Array.from({ length: 4 }, () => painter({ images }));
      for (const target of targets) canvases.push((await paint(target, list())).raster.canvas);
      for (const target of targets) target.value.dispose();
      expect(canvases.filter((canvas) => canvas.width > 0)).toHaveLength(2);
      images.dispose();
      expect(canvases.map((canvas) => [canvas.width, canvas.height])).toEqual(canvases.map(() => [0, 0]));
    } finally { images.dispose(); }
  });

  test('cancelling during asynchronous background painting discards stale contents', async () => {
    const target = painter(); await paint(target, list()); await paint(target, list('ab'));
    let start!: () => void; let resume!: () => void;
    const started = new Promise<void>((resolve) => { start = resolve; });
    const blocked = new Promise<void>((resolve) => { resume = resolve; });
    const original = pptx.paintSlide; let hold = true;
    const spy = spyOn(pptx, 'paintSlide').mockImplementation(async (...args: Parameters<typeof pptx.paintSlide>) => {
      if (hold && !args[1].primitives.length) { hold = false; start(); await blocked; }
      await original(...args);
    });
    restorers.push(() => spy.mockRestore());
    try {
      const confirmed = frame(list('abc')); target.value.paintFrame(confirmed, view(confirmed));
      const completion = target.value.whenPainted();
      await started;
      const cancelled = commandFrameRaster(confirmed, 1)!.canvas;
      target.value.cancel();
      expect([cancelled.width, cancelled.height]).toEqual([0, 0]);
      const next = list('abcd'); const { raster } = await paint(target, next);
      expect(raster.canvas).not.toBe(cancelled);
      expect(surface(raster.canvas).draws).toHaveLength(4);
      const pixels = [...surface(target.canvas).pixels];
      resume(); expect(await completion).toBe(false);
      expect(surface(target.canvas).pixels).toEqual(pixels);
      await expectFull(target, next);
    } finally { resume(); }
  });

  test('checkpoints and held leases preserve pixels until their last release', async () => {
    const target = painter(); const first = await paint(target, list());
    const held = retainCommandFrameRaster(first.confirmed, 1)!;
    const pixels = [...surface(first.raster.canvas).pixels]; const checkpoint = target.value.checkpoint();
    try {
      const second = await paint(target, list('ab'));
      expect(second.raster.canvas).not.toBe(first.raster.canvas);
      checkpoint();
      expect(surface(target.canvas).pixels).toEqual(pixels);
      await paint(target, list('abc')); await paint(target, list('abcd'));
      expect(surface(first.raster.canvas).pixels).toEqual(pixels);
      expect(first.raster.ready).toBe(true);
    } finally { checkpoint.commit?.(); held.release(); }
  });

  test('local echo restores from its held raster after newer frames commit', async () => {
    const target = painter(); const first = await paint(target, list());
    const canvas = document.createElement('canvas'); canvas.width = 160; canvas.height = 120;
    canvas.getContext('2d')!.drawImage(first.raster.canvas, 0, 0);
    const pixels = [...surface(canvas).pixels];
    const echo = createLocalEchoPainter(canvas, { dpr: 1 });
    const snapshot = { ...view(first.confirmed), selection: { slideId: 'slide', shapeId: 'box', storyId: 'story',
      anchor: 0, focus: 1, affinity: 'downstream' as const }, selectionRects: [{ x: 10, y: 12, width: 4, height: 10 }] };
    try {
      echo.paintEcho(snapshot);
      expect(surface(canvas).pixels).not.toEqual(pixels);
      await paint(target, list('ab')); const next = await paint(target, list('abc'));
      expect(next.raster.canvas).not.toBe(first.raster.canvas);
      echo.restoreLine(snapshot);
      expect(surface(canvas).pixels).toEqual(pixels);
    } finally { echo.dispose(); }
  });

  test('pooled canvases do not keep earlier typing frames alive', async () => {
    const WeakReference = (globalThis as unknown as {
      WeakRef: new (value: object) => { deref(): object | undefined } }).WeakRef;
    const target = painter(); const lists: { deref(): object | undefined }[] = [];
    for (let index = 0; index < 40; index += 1) {
      const displayList = list(`a${index}`);
      lists.push(new WeakReference(displayList));
      await paint(target, displayList);
    }
    for (let round = 0; round < 3; round += 1) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      Bun.gc(true);
    }
    expect(lists.filter((value) => value.deref()).length).toBeLessThanOrEqual(4);
  });
});
