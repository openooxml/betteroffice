import { expect, test } from 'bun:test';
import { CANVAS_PAGES_PADDING_PX, canvasPageStack, type DisplayList, type RetainedFrame } from '@betteroffice/docx/layout/render';
import { buildRemotePresencePageMetrics } from '../overlays/remotePresenceGeometry';
import { commandPageGeometry } from './commandPageGeometry';

const A4 = { width: 11906 / 15, height: 16838 / 15 };
const LETTER = { width: 816, height: 1056 };
const sizes = [A4, A4, LETTER, A4, { width: 16838 / 15, height: 11906 / 15 }, A4];
const onGrid = (value: number) => Math.abs(value - Math.round(value)) < 1e-6;

function frame(): RetainedFrame {
  return { pages: sizes.map((size, pageIndex) => ({
    pageId: BigInt(pageIndex + 1), pageIndex, primitiveIds: [],
    page: { pageIndex, ...size, primitives: [] },
  })) } as unknown as RetainedFrame;
}

test('worker page boxes come from the shared stack and sit on whole device pixels', () => {
  for (const ratio of [1, 1.5, 2]) for (const zoom of [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2]) {
    const rasterScale = ratio * zoom;
    const pages = commandPageGeometry(frame(), { documentWidth: 900 }, rasterScale);
    const stack = canvasPageStack(sizes, { scale: { rasterScale, unitScale: rasterScale } });
    expect(pages.map(({ bounds }) => ({ top: bounds.y, width: bounds.width, height: bounds.height }))).toEqual(stack);
    for (const { bounds, transform } of pages) {
      for (const value of [bounds.x, bounds.y, bounds.width, bounds.height]) expect(onGrid(value * rasterScale)).toBe(true);
      expect(transform).toEqual([1, 0, 0, 1, bounds.x, bounds.y]);
    }
  }
});

test('worker page boxes cover their canvas backing exactly', () => {
  const rasterScale = 1.5;
  const pages = commandPageGeometry(frame(), {}, rasterScale);
  for (const [index, { bounds }] of pages.entries()) {
    expect(bounds.width * rasterScale).toBeCloseTo(Math.ceil(sizes[index]!.width * rasterScale), 9);
    expect(bounds.height * rasterScale).toBeCloseTo(Math.ceil(sizes[index]!.height * rasterScale), 9);
  }
});

test('in-thread page metrics read the same stack and keep the exact page size', () => {
  const displayList = { pages: sizes.map((size, pageIndex) => ({ pageIndex, ...size, primitives: [] })) } as DisplayList;
  for (const zoom of [0.5, 1, 1.25]) {
    const metrics = buildRemotePresencePageMetrics(displayList, zoom);
    const stack = canvasPageStack(sizes, { zoom, padding: CANVAS_PAGES_PADDING_PX });
    expect(metrics.tops).toEqual(stack.map(({ top }) => top));
    expect(metrics.bottoms).toEqual(stack.map(({ top, height }) => top + height));
    for (const [index, bottom] of metrics.bottoms.entries()) expect(bottom - metrics.tops[index]!).toBeCloseTo(sizes[index]!.height * zoom, 9);
  }
});
