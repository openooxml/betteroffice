import { afterEach, beforeEach, expect, spyOn, test } from 'bun:test';
import * as browserImages from '../render/browserImage';
import type { SlideDisplayList, TextBoxPrimitive } from '../types';
import { createPptxWorkerRenderer, listenPptxWorkerRender, type WorkerSlideFrame } from './workerRender';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import type { PptxRenderReply, PptxRenderRequest } from '../../../../shared/office-session/render/pptxWorkerRender';

class Bitmap {
  closed = 0;
  constructor(readonly width: number, readonly height: number) {}
  close() { this.closed += 1; }
}
class Canvas {
  static all: Canvas[] = [];
  draws = 0;
  text: string[] = [];
  constructor(public width: number, public height: number) { Canvas.all.push(this); }
  getContext() {
    return new Proxy({ canvas: this, drawImage: () => { this.draws += 1; },
      fillText: (text: string) => { this.text.push(text); }, measureText: (text: string) => ({ width: text.length * 4 }) }, {
      get: (target, key) => Reflect.get(target, key) ?? (() => {}),
    });
  }
}
class Port {
  onmessage: ((event: { data: PptxRenderRequest }) => void) | null = null;
  onmessageerror: (() => void) | null = null;
  replies: PptxRenderReply[] = [];
  transfers: Transferable[][] = [];
  generations: number[] = [];
  postMessage(reply: PptxRenderReply, transfer: Transferable[]) {
    if (reply.kind === 'font-generation') { this.generations.push(reply.fonts); return; }
    this.replies.push(reply); this.transfers.push(transfer);
  }
  start() {}
  close() {}
  send(data: PptxRenderRequest) { this.onmessage?.({ data }); }
}
const originals = new Map<string, PropertyDescriptor | undefined>();
let renderers: ReturnType<typeof createPptxWorkerRenderer>[];
let fonts: string[];
let bitmaps: Bitmap[];
let restores: (() => void)[];
function set(key: string, value: unknown) {
  if (!originals.has(key)) originals.set(key, Object.getOwnPropertyDescriptor(globalThis, key));
  Object.defineProperty(globalThis, key, { configurable: true, value });
}
beforeEach(() => {
  Canvas.all = []; renderers = []; fonts = []; bitmaps = []; restores = [];
  set('OffscreenCanvas', Canvas);
  set('createImageBitmap', async (_source: unknown, _x: number, _y: number, width = 1, height = 1) => {
    const bitmap = new Bitmap(width, height); bitmaps.push(bitmap); return bitmap;
  });
  set('fonts', { add: (face: { family: string }) => fonts.push(face.family), delete() {} });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { return this; }
  });
});
afterEach(() => {
  for (const renderer of renderers) renderer.dispose();
  for (const restore of restores.reverse()) restore();
  for (const [key, descriptor] of originals) {
    if (descriptor) Object.defineProperty(globalThis, key, descriptor);
    else Reflect.deleteProperty(globalThis, key);
  }
  originals.clear();
});
function host() {
  const port = new Port();
  const renderer = createPptxWorkerRenderer(port as unknown as MessagePort);
  renderers.push(renderer);
  port.send({ kind: 'view', scale: 1, dpr: 1 });
  return { port, renderer };
}
function list(text = 'a'): SlideDisplayList {
  const box: TextBoxPrimitive = { kind: 'textBox', objectId: 1, x: 10, y: 8, w: 80, h: 94,
    anchor: 'top', paragraphs: [], lines: [text, 'steady'].map((value, index) => ({
      x: 10, y: 12 + index * 30, width: value.length * 4, height: 10, baseline: 20 + index * 30,
      start: index * 20, end: index * 20 + value.length, caretStops: [],
      runs: [{ text: value, start: index * 20, end: index * 20 + value.length, x: 10,
        width: value.length * 4, fontId: 0, fontFamily: 'Arial', fontSizePx: 8, bold: false,
        italic: false, underline: false, color: '#000000', glyphs: [] }],
    })) };
  return { contractVersion: 1, width: 160, height: 120, background: { kind: 'solid', color: '#ffffff' }, primitives: [box] };
}
function frame(epoch: number, displayList = list(), slideId = 'slide'): WorkerSlideFrame {
  return { slideId, frameEpoch: epoch, displayList, mediaBytes: () => new Uint8Array() };
}
const tick = () => new Promise((resolve) => setTimeout(resolve, 0));
async function replies(port: Port, count: number) {
  for (let index = 0; index < 50 && port.replies.length < count; index++) await tick();
  expect(port.replies).toHaveLength(count);
}
function raster(port: Port, index: number) {
  const reply = port.replies[index]!;
  if (reply.kind !== 'raster') throw new Error(JSON.stringify(reply));
  return reply;
}
function ack(renderer: ReturnType<typeof createPptxWorkerRenderer>, value: WorkerSlideFrame) {
  renderer.frame(value); renderer.afterReply();
}

test('the render port reaches the renderer and never the session host', async () => {
  const listeners = new Set<(message: unknown) => void>();
  const scope: SessionTransport = { post() {}, onError: () => () => {}, close() {},
    listen(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; } };
  const render = listenPptxWorkerRender(scope);
  const seen: unknown[] = [];
  render.transport.listen((message) => seen.push(message));
  const port = new Port();
  for (const listener of [...listeners]) listener({ kind: 'pptx-render-environment', port });
  for (const listener of [...listeners]) listener({ protocol: 1, kind: 'call', id: 1, method: 'open', args: [] });
  expect(seen).toEqual([{ protocol: 1, kind: 'call', id: 1, method: 'open', args: [] }]);
  port.send({ kind: 'view', scale: 1, dpr: 1 });
  render.frame(frame(1)); render.afterReply();
  await replies(port, 1);
  expect(raster(port, 0).frameEpoch).toBe(1);
  render.dispose();
});

test('ack defers paint until posted and coalesces superseded active frames', async () => {
  const { port, renderer } = host();
  renderer.frame(frame(1));
  await tick(); expect(port.replies).toEqual([]);
  renderer.afterReply(); ack(renderer, frame(2, list('latest')));
  await replies(port, 1);
  expect(raster(port, 0).frameEpoch).toBe(2);
  expect(Canvas.all.flatMap((canvas) => canvas.text)).not.toContain('a');
  expect(port.transfers[0]).toEqual([raster(port, 0).bitmap]);
});

test('one changed line transfers its device damage against the preceding epoch', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  expect(raster(port, 0).baseEpoch).toBeNull();
  ack(renderer, frame(2, list('ab'))); await replies(port, 2);
  expect(raster(port, 1).baseEpoch).toBe(1);
  expect(raster(port, 1).rect).toEqual({ x: 8, y: 10, width: 84, height: 14 });
  expect(Canvas.all.flatMap((canvas) => canvas.text).filter((text) => text === 'steady')).toHaveLength(1);
});

test('view, resync, slide and size changes transfer full replacements', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  port.send({ kind: 'view', scale: 0.5, dpr: 2.5 }); await replies(port, 2);
  expect(raster(port, 1)).toMatchObject({ factor: 1.25, baseEpoch: null,
    rect: { x: 0, y: 0, width: 200, height: 150 }, width: 200, height: 150 });
  port.send({ kind: 'resync' }); await replies(port, 3);
  expect(raster(port, 2).baseEpoch).toBeNull();
  ack(renderer, frame(2, list(), 'other')); await replies(port, 4);
  expect(raster(port, 3).baseEpoch).toBeNull();
  ack(renderer, frame(3, { ...list(), width: 180 }, 'other')); await replies(port, 5);
  expect(raster(port, 4).baseEpoch).toBeNull();
});

test('an in-flight bitmap is closed when superseded and only the latest is posted', async () => {
  const { port, renderer } = host();
  let finish!: (bitmap: Bitmap) => void;
  set('createImageBitmap', () => new Promise<Bitmap>((resolve) => { finish = resolve; }));
  ack(renderer, frame(1));
  for (let index = 0; index < 50 && !finish; index++) await tick();
  ack(renderer, frame(2, list('new')));
  const stale = new Bitmap(160, 120); finish(stale);
  set('createImageBitmap', async () => new Bitmap(160, 120));
  await replies(port, 1);
  expect(stale.closed).toBe(1);
  expect(raster(port, 0).frameEpoch).toBe(2);
  expect(raster(port, 0).baseEpoch).toBeNull();
});

test('SVG requests the main painter without allocating raster tiles', async () => {
  const { port, renderer } = host();
  const displayList: SlideDisplayList = { contractVersion: 1, width: 100, height: 100,
    primitives: [{ kind: 'image', name: 'vector', objectId: 1, x: 0, y: 0, w: 10, h: 10, assetId: 'vector.svg' }] };
  ack(renderer, { ...frame(1, displayList), mediaBytes: () => new TextEncoder().encode('<svg width="10" height="10"/>') });
  await replies(port, 1);
  expect(port.replies[0]).toEqual({ kind: 'main-only', slideId: 'slide', frameEpoch: 1, assetIds: ['vector.svg'] });
  expect(Canvas.all).toHaveLength(1);
});

test.each(['OffscreenCanvas', 'createImageBitmap', 'FontFace', 'fonts'])('%s absence reports unsupported once', async (key) => {
  set(key, undefined);
  const { port, renderer } = host();
  ack(renderer, frame(1)); await tick();
  expect(port.replies).toEqual([{ kind: 'unsupported' }]);
});

test('requested fonts gate painting and late fonts repaint the latest frame', async () => {
  let finish!: () => void;
  const loaded = new Promise<void>((resolve) => { finish = resolve; });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { await loaded; return this; }
  });
  const { port, renderer } = host();
  port.send({ kind: 'fonts', faces: [{ family: 'Host', source: new ArrayBuffer(1), descriptors: {} }] });
  ack(renderer, frame(1)); await tick();
  expect(port.replies).toEqual([]); expect(fonts).toEqual([]);
  finish(); await replies(port, 1); expect(fonts).toEqual(['Host']);
  renderer.defaultFonts([{ family: 'Default', bold: false, italic: false, bytes: new ArrayBuffer(1) }]);
  await replies(port, 2);
  expect(fonts).toEqual(['Host', 'Default']);
  expect(raster(port, 1).baseEpoch).toBeNull();
});

test('disposal zeros pooled canvases and retained tiles', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  ack(renderer, frame(2, list('ab'))); await replies(port, 2);
  renderer.dispose();
  expect(Canvas.all.every((canvas) => canvas.width === 0 && canvas.height === 0)).toBe(true);
});

test('a superseding ack still waits for its own reply after an earlier paint was scheduled', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); renderer.frame(frame(2, list('latest')));
  await tick(); expect(port.replies).toEqual([]);
  renderer.afterReply(); await replies(port, 1);
  expect(raster(port, 0).frameEpoch).toBe(2);
});

test('a missing offscreen 2D context reports unsupported', () => {
  set('OffscreenCanvas', class { width = 1; height = 1; getContext() { return null; } });
  const { port } = host();
  expect(port.replies).toEqual([{ kind: 'unsupported' }]);
});

test('a rejected worker font reports an error for the pending frame', async () => {
  set('FontFace', class { async load() { throw new Error('font unavailable'); } });
  const { port, renderer } = host();
  port.send({ kind: 'fonts', faces: [{ family: 'Host', source: new ArrayBuffer(1), descriptors: {} }] });
  ack(renderer, frame(1)); await replies(port, 1);
  expect(port.replies[0]).toEqual({ kind: 'error', frameEpoch: 1, message: 'font unavailable' });
});

test('pictures decode in the worker and a late image gates the latest raster', async () => {
  let finish!: () => void;
  const decoded = new Promise<void>((resolve) => { finish = resolve; });
  let decodes = 0;
  const picture = new Bitmap(10, 10);
  set('createImageBitmap', async (source: unknown, _x: number, _y: number, width = 1, height = 1) => {
    if (source instanceof Blob) { decodes += 1; await decoded; return picture; }
    return new Bitmap(width, height);
  });
  const { port, renderer } = host();
  const displayList: SlideDisplayList = { contractVersion: 1, width: 100, height: 100,
    primitives: [{ kind: 'image', name: 'picture', objectId: 1, x: 0, y: 0, w: 10, h: 10, assetId: 'picture' }] };
  const source = (epoch: number) => ({ ...frame(epoch, displayList), mediaBytes: () => Uint8Array.of(137, 80, 78, 71) });
  ack(renderer, source(1));
  for (let index = 0; index < 50 && !decodes; index++) await tick();
  expect(decodes).toBe(1); expect(port.replies).toEqual([]);
  ack(renderer, source(2)); finish(); await replies(port, 1);
  expect(raster(port, 0).frameEpoch).toBe(2); expect(decodes).toBe(1);
  renderer.dispose(); expect(picture.closed).toBe(1);
});


function pictureList(): SlideDisplayList {
  return { contractVersion: 1, width: 100, height: 100,
    primitives: [{ kind: 'image', name: 'picture', objectId: 1, x: 0, y: 0, w: 10, h: 10, assetId: 'picture' }] };
}

test('worker media is fetched and classified once across frames, views, resyncs and fonts without byte comparisons', async () => {
  const classify = spyOn(browserImages, 'browserPresentationImageBlob');
  restores.push(() => classify.mockRestore());
  const { port, renderer } = host();
  let fetches = 0;
  const source = (epoch: number, slideId = 'slide') => ({ ...frame(epoch, pictureList(), slideId),
    mediaBytes: () => { fetches += 1; return Uint8Array.of(137, 80, 78, 71); } });
  ack(renderer, source(1)); await replies(port, 1);
  const compares = spyOn(Uint8Array.prototype, 'every');
  restores.push(() => compares.mockRestore());
  ack(renderer, source(2)); await replies(port, 2);
  port.send({ kind: 'resync' }); await replies(port, 3);
  port.send({ kind: 'view', scale: 2, dpr: 1 }); await replies(port, 4);
  ack(renderer, source(3, 'other')); await replies(port, 5);
  renderer.defaultFonts([{ family: 'Default', bold: false, italic: false, bytes: new ArrayBuffer(1) }]);
  await replies(port, 6);
  expect(fetches).toBe(1); expect(classify).toHaveBeenCalledTimes(1);
  expect(compares).not.toHaveBeenCalled();
  expect(port.replies.every((reply) => reply.kind === 'raster')).toBe(true);
});

test('resetting sent media invalidates cached classification and decoded pixels for reused asset ids', async () => {
  const classify = spyOn(browserImages, 'browserPresentationImageBlob');
  restores.push(() => classify.mockRestore());
  const { port, renderer } = host();
  let fetches = 0;
  let bytes = Uint8Array.of(137, 80, 78, 71);
  const source = (epoch: number) => ({ ...frame(epoch, pictureList()), mediaBytes: () => { fetches += 1; return bytes; } });
  ack(renderer, source(1)); await replies(port, 1);
  const decoded = bitmaps[0]!;
  renderer.resetMedia();
  expect(decoded.closed).toBe(1);
  bytes = new TextEncoder().encode('<svg width="10" height="10"/>');
  ack(renderer, source(1)); await replies(port, 2);
  ack(renderer, source(2)); await replies(port, 3);
  expect(port.replies.slice(1)).toEqual([1, 2].map((frameEpoch) => ({
    kind: 'main-only', slideId: 'slide', frameEpoch, assetIds: ['picture'],
  })));
  expect(fetches).toBe(2); expect(classify).toHaveBeenCalledTimes(2);
  renderer.resetMedia(); bytes = Uint8Array.of(137, 80, 78, 71);
  ack(renderer, source(1)); await replies(port, 4);
  expect(raster(port, 3).baseEpoch).toBeNull();
  expect(fetches).toBe(3); expect(classify).toHaveBeenCalledTimes(3);
});

test('resetting media during a decode drops the retired paint and closes its image once', async () => {
  let finish!: () => void;
  const decoding = new Promise<void>((resolve) => { finish = resolve; });
  const decoded = new Bitmap(10, 10);
  let started = false;
  set('createImageBitmap', async (source: unknown, _x: number, _y: number, width = 1, height = 1) => {
    if (source instanceof Blob) { started = true; await decoding; return decoded; }
    return new Bitmap(width, height);
  });
  const { port, renderer } = host();
  ack(renderer, { ...frame(1, pictureList()), mediaBytes: () => Uint8Array.of(137, 80, 78, 71) });
  for (let index = 0; index < 50 && !started; index++) await tick();
  expect(started).toBe(true);
  renderer.resetMedia();
  ack(renderer, { ...frame(1, pictureList()), mediaBytes: () => new TextEncoder().encode('<svg width="10" height="10"/>') });
  finish(); await replies(port, 1);
  expect(port.replies[0]).toMatchObject({ kind: 'main-only', frameEpoch: 1, assetIds: ['picture'] });
  expect(decoded.closed).toBe(1);
});

test('an unavailable asset can arrive later and is then fetched only once', async () => {
  const { port, renderer } = host();
  let fetches = 0;
  const source = (epoch: number) => ({ ...frame(epoch, pictureList()), mediaBytes: () => {
    fetches += 1;
    if (fetches === 1) throw new Error('not yet available');
    return Uint8Array.of(137, 80, 78, 71);
  } });
  ack(renderer, source(1)); await replies(port, 1);
  ack(renderer, source(2)); await replies(port, 2);
  ack(renderer, source(3)); await replies(port, 3);
  expect(fetches).toBe(2);
  expect(raster(port, 1).rect).toEqual({ x: 0, y: 0, width: 12, height: 12 });
});

test('paintNow returns the raster for the acknowledged frame without posting it on the port', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  renderer.frame(frame(2, list('ab')));
  expect(renderer.pending()).toBe(true);
  const reply = await renderer.paintNow();
  expect(reply?.kind).toBe('raster'); expect(renderer.pending()).toBe(false);
  if (reply?.kind === 'raster') { expect(reply.frameEpoch).toBe(2); expect(reply.baseEpoch).toBe(1); reply.bitmap.close(); }
  await tick(); await tick();
  expect(port.replies).toHaveLength(1);
});

test('paintNow leaves first paints, slide changes and new media to the port so the reply is not held', async () => {
  const { port, renderer } = host();
  renderer.frame(frame(1));
  expect(await renderer.paintNow()).toBeUndefined();
  await replies(port, 1);
  renderer.frame(frame(2, list(), 'other'));
  expect(await renderer.paintNow()).toBeUndefined();
  await replies(port, 2);
  expect(raster(port, 1).baseEpoch).toBeNull();
});

test('paintNow crops synchronously when the canvas can transfer a bitmap', async () => {
  const transfers: Bitmap[] = [];
  const Transfer = class extends Canvas {
    transferToImageBitmap() { const bitmap = new Bitmap(this.width, this.height); transfers.push(bitmap); return bitmap; }
  };
  set('OffscreenCanvas', Transfer);
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  const before = transfers.length;
  renderer.frame(frame(2, list('ab')));
  const reply = await renderer.paintNow();
  expect(reply?.kind).toBe('raster');
  expect(transfers).toHaveLength(before + 1); expect(bitmaps).toHaveLength(0);
});

test('paintNow gives up after one task while fonts load and the raster follows on the port', async () => {
  let release!: () => void;
  set('FontFace', class {
    constructor(readonly family: string) {}
    load() { return new Promise<this>((resolve) => { release = () => resolve(this); }); }
  });
  const { port, renderer } = host();
  renderer.fonts([{ family: 'Slow', source: new ArrayBuffer(4), descriptors: {} }]);
  await tick();
  renderer.frame(frame(1));
  expect(await renderer.paintNow()).toBeUndefined();
  release();
  await replies(port, 1);
  expect(raster(port, 0).frameEpoch).toBe(1);
});

test('an unchanged display list replies with its content epoch and no raster', async () => {
  const { port, renderer } = host();
  const shared = list();
  ack(renderer, frame(1, shared)); await replies(port, 1);
  const tiles = Canvas.all.length;
  ack(renderer, frame(2, shared)); await replies(port, 2);
  expect(port.replies[1]).toEqual({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 1, baseContent: 1 });
  expect(Canvas.all).toHaveLength(tiles);
  ack(renderer, frame(3, list('ab'))); await replies(port, 3);
  expect(raster(port, 2).baseEpoch).toBe(1);
  port.send({ kind: 'resync' }); await replies(port, 4);
  expect(raster(port, 3).baseEpoch).toBeNull();
});

test('a repaint of the same epoch gets a new content id that the next delta names', async () => {
  const { port, renderer } = host();
  ack(renderer, frame(1)); await replies(port, 1);
  port.send({ kind: 'resync' }); await replies(port, 2);
  const first = raster(port, 0), repaint = raster(port, 1);
  expect(repaint.frameEpoch).toBe(1); expect(repaint.content).not.toBe(first.content);
  ack(renderer, frame(2, list('ab'))); await replies(port, 3);
  expect(raster(port, 2).baseEpoch).toBe(1); expect(raster(port, 2).baseContent).toBe(repaint.content);
});

test('worker fonts stay parked through acknowledged frames until the first view arrives', async () => {
  const constructed: string[] = [];
  set('FontFace', class {
    constructor(readonly family: string) { constructed.push(family); }
    async load() { return this; }
  });
  const port = new Port();
  const renderer = createPptxWorkerRenderer(port as unknown as MessagePort);
  renderers.push(renderer);
  renderer.defaultFonts([{ family: 'Default', bold: false, italic: false, bytes: new ArrayBuffer(1) }]);
  port.send({ kind: 'fonts', faces: [{ family: 'Host', source: new ArrayBuffer(1), descriptors: {} }] });
  ack(renderer, frame(1)); await tick(); await tick();
  expect(constructed).toEqual([]); expect(fonts).toEqual([]); expect(port.replies).toEqual([]);
  port.send({ kind: 'view', scale: 1, dpr: 1 });
  await replies(port, 1);
  expect(constructed).toEqual(['Default', 'Host']); expect(fonts).toEqual(['Default', 'Host']);
  expect(raster(port, 0).baseEpoch).toBeNull();
});

test('the worker keeps no copy of a picture after classifying it and still decodes it from its blob', async () => {
  let decodes = 0;
  set('createImageBitmap', async (source: unknown, _x: number, _y: number, width = 1, height = 1) => {
    if (source instanceof Blob) { decodes += 1; expect(new Uint8Array(await source.arrayBuffer())).toEqual(Uint8Array.of(137, 80, 78, 71)); }
    return new Bitmap(width, height);
  });
  const { port, renderer } = host();
  const Weak = (globalThis as unknown as { WeakRef: new (value: object) => { deref(): object | undefined } }).WeakRef;
  const refs: { deref(): object | undefined }[] = [];
  const source = (epoch: number) => ({ ...frame(epoch, pictureList()),
    mediaBytes: () => { const bytes = Uint8Array.of(137, 80, 78, 71); refs.push(new Weak(bytes.buffer)); return bytes; } });
  ack(renderer, source(1)); await replies(port, 1);
  ack(renderer, source(2)); await replies(port, 2);
  expect(refs).toHaveLength(1); expect(decodes).toBe(1);
  await tick(); Bun.gc(true); await tick(); Bun.gc(true);
  expect(refs[0]!.deref()).toBeUndefined();
});

function prerendering(source: (slideId: string) => { slideId: string; version: string; displayList: SlideDisplayList;
  mediaBytes(assetId: string): Uint8Array } | undefined) {
  const port = new Port();
  const renderer = createPptxWorkerRenderer(port as unknown as MessagePort, () => source);
  renderers.push(renderer);
  port.send({ kind: 'view', scale: 1, dpr: 1 });
  return { port, renderer };
}

test('a prerender paints one slide canvas from the slide source and leaves the active raster deltas intact', async () => {
  const { port, renderer } = prerendering((slideId) => slideId === 'next'
    ? { slideId, version: '7', displayList: list('ahead'), mediaBytes: () => new Uint8Array() } : undefined);
  ack(renderer, frame(1)); await replies(port, 1);
  const before = Canvas.all.length;
  port.send({ kind: 'prerender', slideId: 'next', token: 3 }); await replies(port, 2);
  expect(port.replies[1]).toMatchObject({ kind: 'prerender', slideId: 'next', token: 3, version: '7', fonts: 0, factor: 1,
    width: 160, height: 120 });
  expect(port.transfers[1]).toEqual([(port.replies[1] as { bitmap: Transferable }).bitmap]);
  const painted = Canvas.all.slice(before);
  expect(painted).toHaveLength(1);
  expect(painted[0]!.text).toEqual(['ahead', 'steady']);
  expect([painted[0]!.width, painted[0]!.height]).toEqual([0, 0]);
  ack(renderer, frame(2, list('ab'))); await replies(port, 3);
  expect(raster(port, 2).baseEpoch).toBe(1);
  expect(raster(port, 2).rect).toEqual({ x: 8, y: 10, width: 84, height: 14 });
});

test('only the latest prerender request is answered and unknown slides answer nothing', async () => {
  const { port } = prerendering((slideId) => slideId === 'missing' ? undefined
    : { slideId, version: '1', displayList: list(slideId), mediaBytes: () => new Uint8Array() });
  port.send({ kind: 'prerender', slideId: 'first', token: 1 });
  port.send({ kind: 'prerender', slideId: 'second', token: 2 });
  await replies(port, 1);
  for (let index = 0; index < 5; index++) await tick();
  expect(port.replies).toHaveLength(1);
  expect(port.replies[0]).toMatchObject({ kind: 'prerender', slideId: 'second', token: 2 });
  port.send({ kind: 'prerender', slideId: 'missing', token: 3 });
  for (let index = 0; index < 5; index++) await tick();
  expect(port.replies).toHaveLength(1);
});

test('a prerender with media not ready yet answers nothing and leaves the media to the active frame', async () => {
  const { port, renderer } = prerendering((slideId) => ({ slideId, version: '1', displayList: pictureList(),
    mediaBytes: () => { throw new Error('Slide media is not converted yet'); } }));
  port.send({ kind: 'prerender', slideId: 'next', token: 1 });
  for (let index = 0; index < 5; index++) await tick();
  expect(port.replies).toEqual([]);
  let reads = 0;
  ack(renderer, { ...frame(1, pictureList()), mediaBytes: () => { reads += 1; return Uint8Array.of(137, 80, 78, 71); } });
  await replies(port, 1);
  expect(reads).toBe(1);
  expect(raster(port, 0).frameEpoch).toBe(1);
});

test('worker font changes are announced so finished prerenders can be discarded', async () => {
  const { port } = prerendering(() => undefined);
  port.send({ kind: 'fonts', faces: [{ family: 'Late', source: new ArrayBuffer(1), descriptors: {} }] });
  port.send({ kind: 'fonts', faces: [{ family: 'Later', source: new ArrayBuffer(1), descriptors: {} }] });
  for (let index = 0; index < 10; index++) await tick();
  expect(port.generations).toEqual([1, 2]);
});

