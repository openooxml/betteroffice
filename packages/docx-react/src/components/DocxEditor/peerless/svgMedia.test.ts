import { afterEach, expect, spyOn, test } from 'bun:test';
import { decodeSvgMedia, type DecodedSvgMedia } from './svgMedia';

const restores: Array<() => void> = [];
afterEach(() => { for (const restore of restores.splice(0).reverse()) restore(); });

function replaceGlobal(name: string, value: unknown) {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, name);
  Object.defineProperty(globalThis, name, { configurable: true, writable: true, value });
  restores.push(() => {
    if (descriptor) Object.defineProperty(globalThis, name, descriptor);
    else Reflect.deleteProperty(globalThis, name);
  });
}

async function microtasks() {
  for (let index = 0; index < 8; index++) await Promise.resolve();
}

function host() {
  let now = 0;
  let nextTimer = 0;
  let active = 0;
  let maximum = 0;
  let deferBitmap = false;
  const timers = new Map<number, { at: number; run: () => void }>();
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(((run: () => void, delay = 0) => {
    const id = ++nextTimer;
    timers.set(id, { at: now + delay, run });
    return id;
  }) as unknown as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation(id => { timers.delete(id as unknown as number); });
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  restores.push(() => { timeout.mockRestore(); clear.mockRestore(); clock.mockRestore(); });
  const pictures: Array<{ src: string; finish(width?: number, height?: number): void; fail(): void }> = [];
  replaceGlobal('Image', class {
    src = '';
    naturalWidth = 100;
    naturalHeight = 50;
    decode() {
      active += 1;
      maximum = Math.max(maximum, active);
      return new Promise<void>((resolve, reject) => {
        pictures.push({
          src: this.src,
          finish: (width = 100, height = 50) => {
            this.naturalWidth = width;
            this.naturalHeight = height;
            active -= 1;
            resolve();
          },
          fail: () => { active -= 1; reject(new Error('decode failed')); },
        });
      });
    }
  });
  replaceGlobal('document', {
    createElement: (name: string) => {
      expect(name).toBe('canvas');
      return { width: 0, height: 0, getContext: () => ({ drawImage() {} }) };
    },
  });
  const bitmaps: Array<{ width: number; height: number; closed: number; close(): void }> = [];
  const pendingBitmaps: Array<() => void> = [];
  replaceGlobal('createImageBitmap', (canvas: { width: number; height: number }) => {
    const bitmap = { width: canvas.width, height: canvas.height, closed: 0, close() { this.closed += 1; } };
    bitmaps.push(bitmap);
    return deferBitmap ? new Promise(resolve => pendingBitmaps.push(() => resolve(bitmap))) : Promise.resolve(bitmap);
  });
  return {
    pictures, bitmaps, pendingBitmaps, timers,
    get maximum() { return maximum; },
    deferBitmaps() { deferBitmap = true; },
    async advance(ms = 0) {
      now += ms;
      for (const [id, timer] of [...timers]) {
        if (timer.at <= now && timers.delete(id)) timer.run();
      }
      await microtasks();
    },
  };
}

const media = (count: number) => Array.from({ length: count }, (_, index) => ({ source: `media:${index}`, url: `svg:${index}` }));

test('SVG pictures decode one at a time with a macrotask between pictures', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(3));
  for (let index = 0; index < 3; index++) {
    expect(fake.pictures.length).toBe(index + 1);
    expect(fake.pictures[index]!.src).toBe(`svg:${index}`);
    fake.pictures[index]!.finish();
    await microtasks();
    expect(fake.pictures.length).toBe(index + 1);
    if (index < 2) {
      expect(fake.timers.size).toBe(1);
      await fake.advance();
    }
  }
  await decoding.done;
  expect(fake.maximum).toBe(1);
  expect(decoding.take().map(({ source }) => source)).toEqual(['media:0', 'media:1', 'media:2']);
});

test('SVG pictures after the first take arrive immediately, coalesce per interval, and flush on completion', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(8));
  fake.pictures[0]!.finish();
  await microtasks();
  expect(decoding.take().map(({ source }) => source)).toEqual(['media:0']);
  const batches: DecodedSvgMedia[][] = [];
  decoding.flushTo(batch => batches.push(batch));
  await fake.advance(40);
  fake.pictures[1]!.finish();
  await microtasks();
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([['media:1']]);
  await fake.advance(20);
  fake.pictures[2]!.finish();
  await microtasks();
  expect(fake.timers.size).toBe(2);
  await fake.advance(40);
  fake.pictures[3]!.finish();
  await microtasks();
  expect(fake.timers.size).toBe(2);
  await fake.advance(189);
  expect(batches.length).toBe(1);
  await fake.advance(1);
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([['media:1'], ['media:2', 'media:3']]);
  await fake.advance(300);
  fake.pictures[4]!.finish();
  await microtasks();
  expect(batches.length).toBe(3);
  await fake.advance(10);
  fake.pictures[5]!.finish();
  await microtasks();
  await fake.advance(239);
  fake.pictures[6]!.finish();
  await microtasks();
  expect(batches.length).toBe(3);
  await fake.advance(1);
  expect(batches.length).toBe(4);
  fake.pictures[7]!.finish();
  await decoding.done;
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([
    ['media:1'], ['media:2', 'media:3'], ['media:4'], ['media:5', 'media:6'], ['media:7'],
  ]);
  expect(fake.timers.size).toBe(0);
  expect(decoding.take()).toEqual([]);
});

test('SVG batch delivery flushes pending pictures immediately when attached during decoding', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(2));
  fake.pictures[0]!.finish();
  await microtasks();
  const batches: DecodedSvgMedia[][] = [];
  decoding.flushTo(batch => batches.push(batch));
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([['media:0']]);
  await fake.advance();
  fake.pictures[1]!.finish();
  await decoding.done;
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([['media:0'], ['media:1']]);
  expect(fake.timers.size).toBe(0);
});

test('SVG close stops starting pictures and closes a bitmap completed after close', async () => {
  const fake = host();
  fake.deferBitmaps();
  const decoding = decodeSvgMedia(media(3));
  const batches: DecodedSvgMedia[][] = [];
  decoding.flushTo(batch => batches.push(batch));
  fake.pictures[0]!.finish();
  await microtasks();
  expect(fake.pendingBitmaps.length).toBe(1);
  decoding.close();
  fake.pendingBitmaps[0]!();
  await decoding.done;
  expect(fake.bitmaps[0]!.closed).toBe(1);
  expect(fake.pictures.length).toBe(1);
  expect(fake.timers.size).toBe(0);
  expect(batches).toEqual([]);
  expect(decoding.take()).toEqual([]);
});

test('SVG close during Image decode stops rasterization and preserves taken bitmap ownership', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(4));
  fake.pictures[0]!.finish();
  await microtasks();
  const taken = decoding.take();
  await fake.advance();
  fake.pictures[1]!.finish();
  await microtasks();
  await fake.advance();
  decoding.close();
  decoding.close();
  fake.pictures[2]!.finish();
  await decoding.done;
  expect(taken.length).toBe(1);
  expect(fake.bitmaps.map(bitmap => bitmap.closed)).toEqual([0, 1]);
  expect(fake.pictures.length).toBe(3);
});

test('SVG decoding skips failures and empty pictures in document order at twice the device pixel ratio', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(4), 2);
  fake.pictures[0]!.finish();
  await microtasks();
  await fake.advance();
  fake.pictures[1]!.fail();
  await microtasks();
  await fake.advance();
  fake.pictures[2]!.finish(0, 50);
  await microtasks();
  await fake.advance();
  fake.pictures[3]!.finish(50, 100);
  await decoding.done;
  expect(decoding.take().map(({ source }) => source)).toEqual(['media:0', 'media:3']);
  expect(fake.bitmaps.map(({ width, height }) => [width, height])).toEqual([[400, 200], [200, 400]]);
});

test('SVG decoded pictures retain the edge and document pixel caps across take calls', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(4), 3);
  const sizes: number[] = [];
  for (let index = 0; index < 4; index++) {
    fake.pictures[index]!.finish(10000, 10000);
    await microtasks();
    for (const { image } of decoding.take()) {
      expect(image.width).toBeLessThanOrEqual(4096);
      expect(image.height).toBeLessThanOrEqual(4096);
      sizes.push(image.width * image.height);
    }
    if (index < 3) await fake.advance();
  }
  await decoding.done;
  expect(sizes).toEqual([4096 * 4096, 4096 * 4096, 4096 * 4096]);
  expect(sizes.reduce((sum, pixels) => sum + pixels, 0)).toBe(48 * 1024 * 1024);
});

test('SVG batch delivery flushes pictures when attached after decoding finishes', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(1));
  fake.pictures[0]!.finish();
  await decoding.done;
  const batches: DecodedSvgMedia[][] = [];
  decoding.flushTo(batch => batches.push(batch));
  expect(batches.map(batch => batch.map(({ source }) => source))).toEqual([['media:0']]);
  expect(fake.timers.size).toBe(0);
});

test('SVG close between pictures prevents the next Image from starting', async () => {
  const fake = host();
  const decoding = decodeSvgMedia(media(2));
  fake.pictures[0]!.finish();
  await microtasks();
  decoding.close();
  await fake.advance();
  await decoding.done;
  expect(fake.pictures.length).toBe(1);
  expect(fake.bitmaps[0]!.closed).toBe(1);
  expect(decoding.take()).toEqual([]);
});
