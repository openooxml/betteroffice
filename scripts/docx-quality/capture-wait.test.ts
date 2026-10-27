import { describe, expect, test } from 'bun:test';
import {
  hasInk,
  isUnpaintedCanvas,
  LAYOUT_WAIT_BASE_MS,
  LAYOUT_WAIT_MAX_MS,
  layoutWaitMs,
  settleCanvas,
  type CanvasFrame,
} from './capture-wait';

const portrait = (value = 'a'): CanvasFrame => ({ width: 1241, height: 1755, value });
const landscape = (value = 'b'): CanvasFrame => ({ width: 1755, height: 1241, value });

function clock(frames: (CanvasFrame | null)[]) {
  let time = 0;
  let reads = 0;
  return {
    read: () => frames[Math.min(reads++, frames.length - 1)],
    sleep: async (ms: number) => {
      time += ms;
    },
    now: () => time,
    reads: () => reads,
  };
}

describe('layoutWaitMs', () => {
  test('scales with the reference page count up to a cap', () => {
    expect(layoutWaitMs()).toBe(LAYOUT_WAIT_BASE_MS);
    expect(layoutWaitMs(0)).toBe(LAYOUT_WAIT_BASE_MS);
    expect(layoutWaitMs(22)).toBe(164_000);
    expect(layoutWaitMs(143)).toBe(406_000);
    expect(layoutWaitMs(5000)).toBe(LAYOUT_WAIT_MAX_MS);
  });
});

describe('isUnpaintedCanvas', () => {
  test('flags a blank canvas at the HTML default size', () => {
    expect(isUnpaintedCanvas(300, 150, false)).toBe(true);
    expect(isUnpaintedCanvas(1241, 1755, false)).toBe(false);
    expect(isUnpaintedCanvas(150, 300, false)).toBe(false);
  });

  test('keeps a painted 2 x 1 inch page, which is 300 x 150 at 150 DPI', () => {
    expect(isUnpaintedCanvas(300, 150, true)).toBe(false);
  });
});

describe('hasInk', () => {
  test('reports whether any pixel is not fully transparent', () => {
    expect(hasInk(new Uint8ClampedArray(4 * 6))).toBe(false);
    const pixels = new Uint8ClampedArray(4 * 6);
    pixels[4 * 5 + 3] = 1;
    expect(hasInk(pixels)).toBe(true);
  });
});

describe('settleCanvas', () => {
  const landscapeOnly = (width: number, height: number) => width > height;

  test('returns a matching page after three identical reads', async () => {
    const c = clock([null, landscape('x'), landscape('y'), landscape('y'), landscape('y')]);
    const frame = await settleCanvas({ ...c, matches: landscapeOnly });
    expect(frame).toEqual(landscape('y'));
    expect(c.reads()).toBe(5);
  });

  test('keeps waiting while a stale canvas still has the previous page size', async () => {
    const stale = Array.from({ length: 12 }, () => portrait());
    const c = clock([...stale, landscape(), landscape(), landscape()]);
    const frame = await settleCanvas({ ...c, matches: landscapeOnly });
    expect(frame).toEqual(landscape());
    expect(c.now()).toBe(14 * 200);
  });

  test('restarts the stable count when the size changes but the pixels do not', async () => {
    const c = clock([portrait('same'), portrait('same'), landscape('same'), landscape('same'), landscape('same')]);
    const frame = await settleCanvas({ ...c, matches: () => true });
    expect(frame).toEqual(landscape('same'));
    expect(c.reads()).toBe(5);
  });

  test('waits for an unpainted canvas to paint even when its size matches', async () => {
    const small = (value: string, painted: boolean): CanvasFrame => ({ width: 300, height: 150, value, painted });
    const c = clock([small('blank', false), small('blank', false), small('blank', false), small('blank', false), small('page', true), small('page', true), small('page', true)]);
    const frame = await settleCanvas({ ...c, matches: () => true });
    expect(frame).toEqual(small('page', true));
    expect(c.reads()).toBe(7);
  });

  test('returns a canvas that never paints after the grace period', async () => {
    const blank: CanvasFrame = { width: 300, height: 150, value: 'blank', painted: false };
    const c = clock([blank]);
    const frame = await settleCanvas({ ...c, matches: () => true, mismatchGraceMs: 1_000 });
    expect(frame).toEqual(blank);
    expect(c.now()).toBe(1_400);
  });

  test('gives up on a stable mismatched page after the grace period', async () => {
    const c = clock([portrait()]);
    const frame = await settleCanvas({ ...c, matches: landscapeOnly, mismatchGraceMs: 1_000 });
    expect(frame).toEqual(portrait());
    expect(c.now()).toBe(1_400);
  });

  test('restarts the grace period when the mismatched page changes', async () => {
    const frames = [
      ...Array.from({ length: 4 }, () => portrait('first')),
      ...Array.from({ length: 20 }, () => portrait('second')),
    ];
    const c = clock(frames);
    const frame = await settleCanvas({ ...c, matches: landscapeOnly, mismatchGraceMs: 1_000 });
    expect(frame).toEqual(portrait('second'));
    expect(c.now()).toBe(4 * 200 + 2 * 200 + 1_000);
  });
});
