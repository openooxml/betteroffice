import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, spyOn, test } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx';
import type { PptxRenderReply, PptxWorkerRenderClient } from '../../../../shared/office-session/render/pptxWorkerRender';
import { commandFrameRaster, retainCommandFrameRaster } from './commandFramePainter';
import { createWorkerFramePainter } from './workerFramePainter';
import { createCommandFrameImages } from './commandFramePainter';
import { slideRasterCache } from '../../../../shared/office-session/render/pptxSlideRaster';
import { createLocalEchoPainter } from './localEchoPainter';
import type { PeerlessConfirmedFrame, PeerlessViewSnapshot } from './peerlessViewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(() => { if (ownsDom) GlobalRegistrator.unregister(); });
class Bitmap {
  closed = 0;
  constructor(readonly width = 160, readonly height = 120, readonly pixels = 'worker') {}
  close() { this.closed += 1; }
}
class Client implements PptxWorkerRenderClient {
  supported = true;
  resyncs = 0;
  idles = 0;
  views: number[][] = [];
  listeners = new Set<(reply: PptxRenderReply) => void>();
  view(scale: number, dpr: number) { this.views.push([scale, dpr]); }
  fonts() {}
  resync() { this.resyncs += 1; }
  idle() { this.idles += 1; }
  subscribe(listener: (reply: PptxRenderReply) => void) { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; }
  dispose() {}
  send(reply: PptxRenderReply) { for (const listener of [...this.listeners]) listener(reply); }
}
let pixels: WeakMap<object, string>;
let cleanups: (() => void)[];
let painters: ReturnType<typeof createWorkerFramePainter>[];
let draws: { source: object; target: HTMLCanvasElement; args: number[] }[];
let clears: { target: HTMLCanvasElement; rect: number[] }[];
beforeEach(() => {
  pixels = new WeakMap(); cleanups = []; painters = []; draws = []; clears = [];
  const spy = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    const canvas = this;
    const context = { canvas, fillStyle: '', globalCompositeOperation: 'source-over',
      clearRect(...rect: number[]) { clears.push({ target: canvas, rect }); pixels.set(canvas, ''); },
      fillRect() { pixels.set(canvas, String(context.fillStyle)); },
      drawImage(source: object, ...args: number[]) {
        draws.push({ source, target: canvas, args });
        pixels.set(canvas, pixels.get(source) ?? ('pixels' in source ? String(source.pixels) : ''));
      }, measureText: (text: string) => ({ width: text.length * 4 }),
    };
    return new Proxy(context, { get: (target, key) => Reflect.get(target, key) ?? (() => {}) }) as unknown as CanvasRenderingContext2D;
  } as unknown as HTMLCanvasElement['getContext']);
  cleanups.push(() => spy.mockRestore());
});
afterEach(() => { for (const painter of painters) painter.dispose(); for (const cleanup of cleanups.reverse()) cleanup(); });
function frame(epoch = 1, slideId = 'slide'): PeerlessConfirmedFrame {
  const displayList: SlideDisplayList = { contractVersion: 1, width: 160, height: 120,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [] };
  return { displayList, frame: { slideId, slideIndex: 0, slideRevision: epoch, coversThroughSeq: epoch,
    documentEpoch: 1, incarnation: 'worker', epoch, version: String(epoch),
    displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer, media: [] } };
}
function view(confirmed: PeerlessConfirmedFrame): PeerlessViewSnapshot {
  return { confirmed, selection: null, focused: false, caret: null, selectionRects: [], composition: null,
    echo: null, activeLine: null, activeContextBytes: 0, pending: [], waitingInputCount: 0,
    appliedThroughSeq: 0, presentedThroughSeq: 0, thumbnails: new Map(), dirtyThumbnailIds: new Set(),
    editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: null, storedMarks: {},
      formatting: null, alignment: null, hyperlink: null, table: null, canEditText: true } } };
}
function host(options: Parameters<typeof createWorkerFramePainter>[2] = {}) {
  const client = new Client(), canvas = document.createElement('canvas');
  const commits: PeerlessConfirmedFrame[] = [];
  const painter = createWorkerFramePainter(canvas, client, { dpr: 1, ...options,
    onCommit(frame, view) { commits.push(frame); options.onCommit?.(frame, view); },
  });
  painters.push(painter);
  return { client, canvas, painter, commits };
}
function raster(confirmed: PeerlessConfirmedFrame, extra: Partial<Extract<PptxRenderReply, { kind: 'raster' }>> = {}) {
  const bitmap = new Bitmap();
  const reply = { kind: 'raster' as const, slideId: confirmed.frame.slideId, frameEpoch: confirmed.frame.epoch,
    factor: 1, baseEpoch: null as number | null, rect: { x: 0, y: 0, width: 160, height: 120 }, bitmap: bitmap as unknown as ImageBitmap,
    width: 160, height: 120, ...extra };
  return { content: reply.frameEpoch, baseContent: reply.baseEpoch, ...reply };
}
const closed = (reply: ReturnType<typeof raster>) => (reply.bitmap as unknown as Bitmap).closed;
async function commit(h: ReturnType<typeof host>, confirmed = frame(), reply = raster(confirmed)) {
  h.painter.paintFrame(confirmed, view(confirmed)); h.client.send(reply);
  expect(await h.painter.whenPainted()).toBe(true);
  return confirmed;
}

test('only matching frame, slide and factor commit; all bitmaps close once', async () => {
  const h = host(), latest = frame(2);
  h.painter.paintFrame(latest, view(latest));
  const discarded = [raster(frame(1)), raster(frame(2, 'other')), raster(latest, { factor: 2 })];
  for (const reply of discarded) h.client.send(reply);
  expect(h.commits).toEqual([]);
  expect(discarded.map(closed)).toEqual([1, 1, 1]);
  const matching = raster(latest); h.client.send(matching);
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([latest]); expect(closed(matching)).toBe(1);
  expect(pixels.get(h.canvas)).toBe('worker');
});

test('pending paint preserves visible pixels; delta bases require the exact backing epoch', async () => {
  const h = host(); const first = await commit(h);
  const next = frame(2); h.painter.paintFrame(next, view(next));
  expect(pixels.get(h.canvas)).toBe('worker');
  const resyncs = h.client.resyncs;
  const mismatch = raster(next, { baseEpoch: 0 }); h.client.send(mismatch);
  expect(h.client.resyncs).toBe(resyncs + 1); expect(closed(mismatch)).toBe(1);
  expect(h.commits).toEqual([first]);
  const patch = raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap });
  h.client.send(patch); expect(await h.painter.whenPainted()).toBe(true);
  expect(pixels.get(h.canvas)).toBe('delta'); expect(closed(patch)).toBe(1);
});

test.each(['main-only', 'error'] as const)('%s falls back per frame and later frames use worker pixels', async (kind) => {
  const h = host(), confirmed = frame();
  h.painter.paintFrame(confirmed, view(confirmed));
  h.client.send(kind === 'error' ? { kind, frameEpoch: 1, message: 'decode failed' }
    : { kind, frameEpoch: 1, slideId: 'slide', assetIds: ['vector.svg'] });
  expect(await h.painter.whenPainted()).toBe(true);
  expect(pixels.get(h.canvas)).toBe('#ffffff');
  expect(commandFrameRaster(confirmed, 1, 1)?.ready).toBe(true);
  await commit(h, frame(2));
  expect(pixels.get(h.canvas)).toBe('worker');
});

test('unsupported switches the session to the main painter', async () => {
  const h = host(); h.client.send({ kind: 'unsupported' });
  const confirmed = frame(); h.painter.paintFrame(confirmed, view(confirmed));
  expect(await h.painter.whenPainted()).toBe(true);
  expect(pixels.get(h.canvas)).toBe('#ffffff');
  const ignored = raster(confirmed); h.client.send(ignored); expect(closed(ignored)).toBe(1);
});

test('selection restoration uses a registered clean raster and held frames remain intact', async () => {
  const h = host(), first = await commit(h);
  const held = retainCommandFrameRaster(first, 1)!;
  const clean = commandFrameRaster(first, 1, 1)!.canvas;
  expect(pixels.get(clean)).toBe('worker');
  const echo = createLocalEchoPainter(h.canvas, { dpr: 1 });
  try {
    const snapshot = { ...view(first), selection: { slideId: 'slide', shapeId: 'box', storyId: 'story',
      anchor: 0, focus: 1, affinity: 'downstream' as const }, selectionRects: [{ x: 10, y: 10, width: 5, height: 5 }] };
    const before = draws.length;
    echo.paintEcho(snapshot);
    echo.restoreLine(snapshot);
    expect(draws.slice(before).some(({ source, target }) => source === clean && target === h.canvas)).toBe(true);
    await commit(h, frame(2), raster(frame(2), { bitmap: new Bitmap(160, 120, 'next') as unknown as ImageBitmap }));
    expect(pixels.get(clean)).toBe('worker'); expect(clean.width).toBe(160);
  } finally { echo.dispose(); held.release(); }
  expect(clean.width).toBe(0);
});

test('checkpoint restores pixels, leases and completed whenPainted state', async () => {
  const h = host(), first = await commit(h);
  const checkpoint = h.painter.checkpoint();
  await commit(h, frame(2), raster(frame(2), { bitmap: new Bitmap(160, 120, 'next') as unknown as ImageBitmap }));
  checkpoint();
  expect(pixels.get(h.canvas)).toBe('worker');
  expect(await h.painter.whenPainted()).toBe(true);
  expect(commandFrameRaster(first, 1, 1)?.ready).toBe(true);
  checkpoint(); checkpoint.commit?.();
});

test('checkpoint(false) preserves visible pixels and an old checkpoint cannot roll back later paints', async () => {
  const h = host(); await commit(h);
  const noPixels = h.painter.checkpoint(false);
  await commit(h, frame(2), raster(frame(2), { bitmap: new Bitmap(160, 120, 'next') as unknown as ImageBitmap }));
  noPixels(); expect(pixels.get(h.canvas)).toBe('next');
  const old = h.painter.checkpoint();
  await commit(h, frame(3)); await commit(h, frame(4));
  old(); expect(h.commits[h.commits.length - 1]!.frame.epoch).toBe(4);
});

test('cancellation during beforeCommit retains then closes the bitmap once', async () => {
  let finish!: () => void;
  const blocked = new Promise<void>((resolve) => { finish = resolve; });
  const h = host({ beforeCommit: () => blocked }), confirmed = frame(), reply = raster(confirmed);
  h.painter.paintFrame(confirmed, view(confirmed)); h.client.send(reply);
  expect(closed(reply)).toBe(0);
  const completion = h.painter.whenPainted(); h.painter.cancel(); finish();
  expect(await completion).toBe(false);
  await Promise.resolve(); await Promise.resolve();
  expect(closed(reply)).toBe(1); expect(h.commits).toEqual([]);
});

test('freeze discards pending rasters and disposal releases backing storage', async () => {
  const h = host(), first = await commit(h);
  const clean = commandFrameRaster(first, 1, 1)!.canvas;
  const next = frame(2); h.painter.paintFrame(next, view(next));
  const completion = h.painter.whenPainted(); h.painter.freeze();
  const ignored = raster(next); h.client.send(ignored);
  expect(await completion).toBe(false); expect(closed(ignored)).toBe(1);
  h.painter.dispose(); expect(clean.width).toBe(0); expect(h.client.listeners.size).toBe(0);
});

test('an unchanged confirmed frame follows a newer ack raster epoch without a new paint request', async () => {
  const h = host(), confirmed = frame();
  h.painter.paintFrame(confirmed, view(confirmed));
  confirmed.frame = { ...confirmed.frame, epoch: 2 };
  const reply = raster(confirmed); h.client.send(reply);
  expect(await h.painter.whenPainted()).toBe(true);
  expect(closed(reply)).toBe(1); expect(h.commits).toEqual([confirmed]);
});

test('a raster that outran its duplicate-epoch ack commits once the epoch is adopted, without a resync', async () => {
  const h = host(), confirmed = frame();
  h.painter.paintFrame(confirmed, view(confirmed));
  const early = raster(frame(2)); h.client.send(early);
  expect(closed(early)).toBe(0); expect(h.commits).toEqual([]);
  confirmed.frame = { ...confirmed.frame, epoch: 2 };
  const resyncs = h.client.resyncs;
  h.painter.resync?.();
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([confirmed]); expect(h.client.resyncs).toBe(resyncs);
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(closed(early)).toBe(1);
  h.painter.resync?.();
  expect(h.client.resyncs).toBe(resyncs);
});

test('restoring an unfinished checkpoint paints its saved frame even after the worker advanced', async () => {
  const h = host(); await commit(h);
  const saved = frame(2); h.painter.paintFrame(saved, view(saved));
  const checkpoint = h.painter.checkpoint();
  const newer = frame(3); h.painter.paintFrame(newer, view(newer));
  checkpoint();
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits[h.commits.length - 1]).toBe(saved);
  expect(pixels.get(h.canvas)).toBe('#ffffff');
  const late = raster(newer); h.client.send(late); expect(closed(late)).toBe(1);
});

test('a commit callback can freeze the painter without cancelling its completed paint', async () => {
  let freeze = () => {};
  const h = host({ onCommit: () => freeze() }); freeze = () => h.painter.freeze();
  await commit(h);
  expect(h.commits).toHaveLength(1);
});

test('sequential rect rasters reuse the backing and blit only their rects without outside leases', async () => {
  const h = host(), first = await commit(h);
  const backing = commandFrameRaster(first, 1, 1)!.canvas;
  const allocations = spyOn(document, 'createElement');
  cleanups.push(() => allocations.mockRestore());
  const drawStart = draws.length, clearStart = clears.length;
  const rect = { x: 8, y: 10, width: 84, height: 14 };
  for (let epoch = 2; epoch <= 6; epoch += 1) {
    const confirmed = frame(epoch);
    await commit(h, confirmed, raster(confirmed, { baseEpoch: epoch - 1, rect,
      bitmap: new Bitmap(rect.width, rect.height, String(epoch)) as unknown as ImageBitmap }));
    expect(commandFrameRaster(confirmed, 1, 1)!.canvas).toBe(backing);
  }
  expect(allocations.mock.calls.filter(([tag]) => tag === 'canvas')).toHaveLength(0);
  expect(draws.slice(drawStart).filter(({ target }) => target === h.canvas).map(({ args }) => args))
    .toEqual(Array.from({ length: 5 }, () => [8, 10, 84, 14, 8, 10, 84, 14]));
  expect(clears.slice(clearStart).filter(({ target }) => target === h.canvas).map(({ rect }) => rect))
    .toEqual(Array.from({ length: 5 }, () => [8, 10, 84, 14]));
  expect(draws.slice(drawStart).every(({ source }) => !(source instanceof HTMLCanvasElement) || source === backing)).toBe(true);
});

test('a held raster forces one backing copy while later rects reuse it and preserve the held pixels', async () => {
  const h = host(), first = await commit(h);
  const held = retainCommandFrameRaster(first, 1)!;
  const original = held.raster.canvas;
  const allocations = spyOn(document, 'createElement');
  cleanups.push(() => allocations.mockRestore());
  const drawStart = draws.length;
  let nextBacking: HTMLCanvasElement | undefined;
  try {
    for (let epoch = 2; epoch <= 5; epoch += 1) {
      const confirmed = frame(epoch);
      await commit(h, confirmed, raster(confirmed, { baseEpoch: epoch - 1,
        rect: { x: 8, y: 10, width: 84, height: 14 },
        bitmap: new Bitmap(84, 14, String(epoch)) as unknown as ImageBitmap }));
      const clean = commandFrameRaster(confirmed, 1, 1)!.canvas;
      nextBacking ??= clean;
      expect(clean).toBe(nextBacking); expect(clean).not.toBe(original);
      expect(pixels.get(original)).toBe('worker'); expect(original.width).toBe(160);
    }
    expect(allocations.mock.calls.filter(([tag]) => tag === 'canvas')).toHaveLength(1);
    expect(draws.slice(drawStart).filter(({ source, target }) => source === original && target !== h.canvas))
      .toEqual([{ source: original, target: nextBacking!, args: [0, 0] }]);
  } finally { held.release(); }
  expect(original.width).toBe(0);
});

test('selection paint invalidates a visible backing and the next rect raster blits the full clean slide', async () => {
  const h = host(), first = await commit(h);
  const echo = createLocalEchoPainter(h.canvas, { dpr: 1 });
  try {
    echo.paintEcho({ ...view(first), selection: { slideId: 'slide', shapeId: 'box', storyId: 'story',
      anchor: 0, focus: 1, affinity: 'downstream' }, selectionRects: [{ x: 10, y: 10, width: 5, height: 5 }] });
    const drawStart = draws.length;
    const next = frame(2);
    await commit(h, next, raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
      bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
    expect(draws.slice(drawStart).filter(({ target }) => target === h.canvas).map(({ args }) => args)).toEqual([[0, 0]]);
  } finally { echo.dispose(); }
});

test('a composed visible frame requires a full blit before subsequent rect blits', async () => {
  let compose = true;
  const h = host({ compose(raster) { return compose ? { canvas: raster.canvas, release() {} } : undefined; } });
  await commit(h); compose = false;
  const drawStart = draws.length;
  for (let epoch = 2; epoch <= 3; epoch += 1) {
    const confirmed = frame(epoch);
    await commit(h, confirmed, raster(confirmed, { baseEpoch: epoch - 1,
      rect: { x: 8, y: 10, width: 84, height: 14 }, bitmap: new Bitmap(84, 14) as unknown as ImageBitmap }));
  }
  expect(draws.slice(drawStart).filter(({ target }) => target === h.canvas).map(({ args }) => args))
    .toEqual([[0, 0], [8, 10, 84, 14, 8, 10, 84, 14]]);
});

test('a reply that arrives before its frame commits on paint without arming another view request', async () => {
  const h = host(), first = await commit(h), next = frame(2);
  const early = raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap });
  h.client.send(early);
  expect(h.commits).toEqual([first]); expect(closed(early)).toBe(0);
  const views = h.client.views.length;
  h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([first, next]); expect(pixels.get(h.canvas)).toBe('delta');
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(h.client.views).toHaveLength(views); expect(closed(early)).toBe(1);
});

test('an older same-epoch patch arriving after a newer stashed raster never replaces it', async () => {
  const h = host(); await commit(h);
  const next = frame(2);
  const newer = raster(next, { content: 5, bitmap: new Bitmap(160, 120, 'newer') as unknown as ImageBitmap });
  const older = raster(next, { baseEpoch: 1, content: 4, baseContent: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'older') as unknown as ImageBitmap });
  h.client.send(newer); h.client.send(older);
  h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  expect(pixels.get(h.canvas)).toBe('newer');
});

test('an older same-epoch unchanged reply arriving after a newer stashed raster never replaces it', async () => {
  const h = host(); await commit(h);
  const next = frame(2);
  const newer = raster(next, { content: 5, bitmap: new Bitmap(160, 120, 'newer') as unknown as ImageBitmap });
  h.client.send(newer);
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 1, baseContent: 1 });
  h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  expect(pixels.get(h.canvas)).toBe('newer');
});

test('a stashed reply for a superseded frame closes and the newer frame waits for the worker', async () => {
  const h = host(), first = await commit(h);
  const stale = raster(frame(2)); h.client.send(stale);
  const latest = frame(3), views = h.client.views.length;
  h.painter.paintFrame(latest, view(latest));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(closed(stale)).toBe(1); expect(h.commits).toEqual([first]);
  expect(h.client.views).toHaveLength(views + 1);
  await commit(h, latest, raster(latest));
  expect(h.commits).toEqual([first, latest]);
});

test('an echo lease alone patches the backing in place; disposal releases it', async () => {
  const h = host(), first = await commit(h);
  const clean = commandFrameRaster(first, 1, 1)!.canvas;
  const echo = createLocalEchoPainter(h.canvas, { dpr: 1 });
  try {
    echo.paintEcho(view(first));
    const next = frame(2);
    await commit(h, next, raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
      bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
    expect(commandFrameRaster(next, 1, 1)!.canvas).toBe(clean);
    expect(pixels.get(clean)).toBe('delta');
  } finally { echo.dispose(); }
  expect(clean.width).toBe(160);
});

test.each(['after', 'before'] as const)('the first frame paints on main and a worker raster arriving %s that commit becomes the patch base', async (order) => {
  const h = host({ mainFirstFrame: true }), first = frame();
  const full = raster(first);
  h.painter.paintFrame(first, view(first));
  if (order === 'before') h.client.send(full);
  expect(await h.painter.whenPainted()).toBe(true);
  const resyncs = h.client.resyncs;
  if (order === 'after') h.client.send(full);
  expect(pixels.get(h.canvas)).toBe('#ffffff'); expect(h.commits).toEqual([first]);
  expect(closed(full)).toBe(1);
  const next = frame(2);
  await commit(h, next, raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
  expect(h.client.resyncs).toBe(resyncs); expect(pixels.get(h.canvas)).toBe('delta');
});

test('until a worker raster lands, later frames keep painting on main; then deltas patch the shadowed backing', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const next = frame(2); h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([first, next]); expect(pixels.get(h.canvas)).toBe('#ffffff');
  const resyncs = h.client.resyncs, views = h.client.views.length;
  h.client.send(raster(next));
  const third = frame(3);
  await commit(h, third, raster(third, { baseEpoch: 2, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
  expect(pixels.get(h.canvas)).toBe('delta'); expect(h.client.resyncs).toBe(resyncs);
  expect(h.client.views).toHaveLength(views + 1);
});

test('an unchanged reply commits its frame from the backing without drawing', async () => {
  const h = host(), first = await commit(h), next = frame(2);
  h.painter.paintFrame(next, view(next));
  const before = draws.length;
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 1, baseContent: 1 });
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([first, next]);
  expect(draws.slice(before).filter(({ target }) => target === h.canvas)).toEqual([]);
  const third = frame(3);
  await commit(h, third, raster(third, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
  expect(pixels.get(h.canvas)).toBe('delta');
});

test('an unchanged reply against another content epoch asks for a full replacement', async () => {
  const h = host(), first = await commit(h), next = frame(2), resyncs = h.client.resyncs;
  h.painter.paintFrame(next, view(next));
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 0, baseContent: 0 });
  expect(h.client.resyncs).toBe(resyncs + 1); expect(h.commits).toEqual([first]);
  await commit(h, next, raster(next));
  expect(h.commits).toEqual([first, next]);
});

test('an older reply never replaces a newer stash, and consuming the stash clears the watchdog', async () => {
  const h = host(), first = await commit(h);
  const newer = raster(frame(3)), older = raster(frame(2));
  h.client.send(newer); h.client.send(older);
  expect(closed(newer)).toBe(0);
  const second = frame(2); h.painter.paintFrame(second, view(second));
  const idles = h.client.idles;
  const third = frame(3); h.painter.paintFrame(third, view(third));
  expect(h.client.idles).toBe(idles + 1);
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.commits).toEqual([first, third]);
});

test('the handover job adopts every newer full raster for its frame', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  const early = raster(first), later = raster(first, { bitmap: new Bitmap(160, 120, 'held') as unknown as ImageBitmap });
  h.client.send(early); h.client.send(later);
  expect(closed(early)).toBe(1);
  expect(await h.painter.whenPainted()).toBe(true);
  expect(draws.some(({ source }) => source === later.bitmap)).toBe(true);
  const refreshed = raster(first, { bitmap: new Bitmap(160, 120, 'fonts') as unknown as ImageBitmap });
  h.client.send(refreshed);
  expect(draws.some(({ source }) => source === refreshed.bitmap)).toBe(true);
  expect(pixels.get(h.canvas)).toBe('#ffffff');
});

test('a delta naming a newer repaint of the backing epoch waits for that repaint', async () => {
  const h = host(), first = await commit(h), next = frame(2), resyncs = h.client.resyncs;
  h.painter.paintFrame(next, view(next));
  h.client.send(raster(next, { baseEpoch: 1, baseContent: 7, content: 8, rect: { x: 8, y: 10, width: 84, height: 14 } }));
  expect(h.client.resyncs).toBe(resyncs + 1); expect(h.commits).toEqual([first]);
});

test('with main painting the first frame the worker gets its view only after that commit', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  expect(h.client.views).toEqual([]); expect(h.client.resyncs).toBe(0);
  h.painter.paintFrame(first, view(first));
  expect(h.client.views).toEqual([]);
  expect(await h.painter.whenPainted()).toBe(true);
  expect(h.client.views).toEqual([[1, 1]]); expect(h.client.resyncs).toBe(1);
  await commit(h, frame(2));
  expect(h.client.resyncs).toBe(1);
});

test('retiring the main painter after a worker commit empties its canvas pool', async () => {
  const images = createCommandFrameImages();
  const h = host({ images }), confirmed = frame();
  h.painter.paintFrame(confirmed, view(confirmed));
  h.client.send({ kind: 'main-only', frameEpoch: 1, slideId: 'slide', assetIds: ['vector.svg'] });
  expect(await h.painter.whenPainted()).toBe(true);
  await commit(h, frame(2));
  await commit(h, frame(3));
  expect(slideRasterCache(images).pool.canvases).toEqual([]);
  images.dispose();
});

test.each(['after', 'before'] as const)('the main painter retires and trims its decodes once the worker shadows the first frame (raster %s commit)', async (order) => {
  const images = createCommandFrameImages();
  const trim = spyOn(images, 'trim');
  const h = host({ mainFirstFrame: true, images }), first = frame(), full = raster(first);
  h.painter.paintFrame(first, view(first));
  if (order === 'before') h.client.send(full);
  expect(await h.painter.whenPainted()).toBe(true);
  if (order === 'after') h.client.send(full);
  await Promise.resolve();
  expect(trim).toHaveBeenCalledTimes(1);
  expect(slideRasterCache(images).pool.canvases).toEqual([]);
  expect(pixels.get(h.canvas)).toBe('#ffffff'); expect(h.commits).toEqual([first]); expect(closed(full)).toBe(1);
  const resyncs = h.client.resyncs, next = frame(2);
  await commit(h, next, raster(next, { baseEpoch: 1, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap }));
  expect(h.client.resyncs).toBe(resyncs); expect(pixels.get(h.canvas)).toBe('delta');
  expect(trim).toHaveBeenCalledTimes(1);
  images.dispose();
});

test('a shadow of an older frame keeps the main painter until its own frame is shadowed', async () => {
  const images = createCommandFrameImages();
  const trim = spyOn(images, 'trim');
  const h = host({ mainFirstFrame: true, images }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const next = frame(2); h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  h.client.send(raster(first));
  await Promise.resolve();
  expect(trim).not.toHaveBeenCalled();
  h.client.send(raster(next));
  expect(trim).toHaveBeenCalledTimes(1); expect(h.commits).toEqual([first, next]);
  images.dispose();
});

test('adoption lets a soft echo lease rebind to the worker raster, which frees the retired main raster', async () => {
  const images = createCommandFrameImages();
  let soft: ReturnType<typeof retainCommandFrameRaster>;
  const h = host({ mainFirstFrame: true, images, onCommit(frame) { soft ??= retainCommandFrameRaster(frame, 1, true); },
    onAdopt(frame) { soft?.release(); soft = retainCommandFrameRaster(frame, 1, true); } });
  const first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const main = commandFrameRaster(first, 1, 1)!.canvas;
  expect(main.width).toBe(160);
  h.client.send(raster(first));
  expect(main.width).toBe(0);
  const adopted = commandFrameRaster(first, 1, 1);
  expect(adopted?.canvas).not.toBe(main); expect(soft!.raster).toBe(adopted!);
  soft!.retain().release(); soft!.release();
  expect(commandFrameRaster(first, 1, 1)).toBe(adopted);
  expect(slideRasterCache(images).pool.canvases).toEqual([]);
  images.dispose();
});

test('without a rebind a soft lease keeps the retired main raster usable until it is released', async () => {
  const images = createCommandFrameImages();
  let soft: ReturnType<typeof retainCommandFrameRaster>;
  const h = host({ mainFirstFrame: true, images, onCommit(frame) { soft ??= retainCommandFrameRaster(frame, 1, true); } });
  const first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const main = soft!.raster;
  h.client.send(raster(first));
  expect(main.ready).toBe(true); expect(main.canvas.width).toBe(160);
  soft!.retain().release(); soft!.release();
  await commit(h, frame(2));
  expect(main.canvas.width).toBe(0); expect(slideRasterCache(images).pool.canvases).toEqual([]);
  images.dispose();
});

test('a worker raster older than the shadowed epoch is dropped and later frames keep their base', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const next = frame(2); h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  h.client.send(raster(next));
  const stale = raster(first, { bitmap: new Bitmap(160, 120, 'stale') as unknown as ImageBitmap });
  h.client.send(stale);
  expect(closed(stale)).toBe(1);
  const third = frame(3), resyncs = h.client.resyncs;
  h.painter.paintFrame(third, view(third));
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 3, baseEpoch: 2, baseContent: 2 });
  expect(h.commits).toEqual([first, next, third]); expect(h.client.resyncs).toBe(resyncs);
  expect(pixels.get(h.canvas)).toBe('worker');
});

test('a stale worker raster after adoption leaves the adopted raster and its pixels in place', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  h.client.send(raster(first));
  const adopted = commandFrameRaster(first, 1, 1);
  expect(adopted).toBeDefined(); expect(pixels.get(adopted!.canvas)).toBe('worker');
  first.frame = { ...first.frame, epoch: 2 };
  const stale = raster(frame(1), { bitmap: new Bitmap(160, 120, 'stale') as unknown as ImageBitmap });
  h.client.send(stale);
  expect(commandFrameRaster(first, 1, 1)).toBe(adopted); expect(pixels.get(adopted!.canvas)).toBe('worker');
  expect(closed(stale)).toBe(1);
  const next = frame(3);
  await commit(h, next, raster(next));
  expect(pixels.get(h.canvas)).toBe('worker');
});

test('a late repaint older than the shadowed content of the same epoch is dropped', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  h.client.send(raster(first, { content: 5 }));
  const adopted = commandFrameRaster(first, 1, 1)!;
  const stale = raster(first, { content: 4, bitmap: new Bitmap(160, 120, 'stale') as unknown as ImageBitmap });
  h.client.send(stale);
  expect(closed(stale)).toBe(1); expect(commandFrameRaster(first, 1, 1)).toBe(adopted);
  const next = frame(2), resyncs = h.client.resyncs;
  h.painter.paintFrame(next, view(next));
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 1, baseContent: 5 });
  expect(h.commits).toEqual([first, next]); expect(h.client.resyncs).toBe(resyncs);
  expect(pixels.get(h.canvas)).toBe('worker');
});

test('a main-painted frame whose worker reply is unchanged still hands over to the worker', async () => {
  const images = createCommandFrameImages();
  const trim = spyOn(images, 'trim');
  const h = host({ mainFirstFrame: true, images }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const next = frame(2); h.painter.paintFrame(next, view(next));
  expect(await h.painter.whenPainted()).toBe(true);
  h.client.send(raster(first));
  expect(trim).not.toHaveBeenCalled();
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 2, baseEpoch: 1, baseContent: 1 });
  expect(trim).toHaveBeenCalledTimes(1); expect(commandFrameRaster(next, 1, 1)).toBeDefined();
  const third = frame(3), resyncs = h.client.resyncs;
  h.painter.paintFrame(third, view(third));
  h.client.send({ kind: 'unchanged', slideId: 'slide', frameEpoch: 3, baseEpoch: 2, baseContent: 1 });
  expect(h.commits).toEqual([first, next, third]); expect(h.client.resyncs).toBe(resyncs);
  images.dispose();
});

test('a worker delta that cannot patch the handover backing asks for a full replacement', async () => {
  const h = host({ mainFirstFrame: true }), first = frame();
  h.painter.paintFrame(first, view(first));
  expect(await h.painter.whenPainted()).toBe(true);
  const resyncs = h.client.resyncs;
  const delta = raster(first, { baseEpoch: 0, baseContent: 9, rect: { x: 8, y: 10, width: 84, height: 14 },
    bitmap: new Bitmap(84, 14, 'delta') as unknown as ImageBitmap });
  h.client.send(delta);
  expect(h.client.resyncs).toBe(resyncs + 1); expect(closed(delta)).toBe(1);
});

test('a stale full raster of the committed epoch never replaces newer pixels', async () => {
  const h = host(), first = frame();
  await commit(h, first, raster(first, { content: 5 }));
  const stale = raster(first, { content: 4, bitmap: new Bitmap(160, 120, 'stale') as unknown as ImageBitmap });
  h.client.send(stale);
  await Promise.resolve();
  expect(pixels.get(h.canvas)).toBe('worker'); expect(closed(stale)).toBe(1); expect(h.commits).toEqual([first]);
});
