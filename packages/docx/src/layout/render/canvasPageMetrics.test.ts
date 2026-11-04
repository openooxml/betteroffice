import { describe, expect, test } from 'bun:test';
import {
  CANVAS_PAGE_GAP_PX,
  CANVAS_PAGES_PADDING_PX,
  canvasPageLength,
  canvasPageStack,
  canvasPageTops,
  snapCanvasLength,
} from './canvasPageMetrics';
import type { DisplayListQueries } from './displayListQueries';

const A4 = { width: 11906 / 15, height: 16838 / 15 };
const LETTER = { width: 816, height: 1056 };
const LANDSCAPE = { width: 16838 / 15, height: 11906 / 15 };
const ZOOMS = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2];
const RATIOS = [1, 1.25, 1.5, 2, 3];
const onGrid = (value: number) => Math.abs(value - Math.round(value)) < 1e-6;

describe('canvas page stack', () => {
  test('an A4 page box spans its canvas backing pixels exactly', () => {
    const [box] = canvasPageStack([A4], { scale: { rasterScale: 1, unitScale: 1 } });
    expect(box).toEqual({ top: 0, width: 794, height: 1123 });
    expect(canvasPageLength(A4.height, { rasterScale: 2, unitScale: 2 })).toBe(Math.ceil(A4.height * 2) / 2);
  });

  test('every box and top falls on whole device pixels for a transformed column', () => {
    const pages = [A4, LETTER, LANDSCAPE, A4, A4];
    for (const zoom of ZOOMS) for (const ratio of RATIOS) {
      const rasterScale = ratio * zoom;
      const scale = { rasterScale, unitScale: rasterScale };
      for (const [index, box] of canvasPageStack(pages, { scale }).entries()) {
        expect(onGrid(box.top * rasterScale)).toBe(true);
        expect(box.height * rasterScale).toBeCloseTo(Math.ceil(pages[index]!.height * rasterScale), 6);
        expect(box.width * rasterScale).toBeCloseTo(Math.ceil(pages[index]!.width * rasterScale), 6);
      }
    }
  });

  test('every box and top falls on whole device pixels for a column sized at the zoom', () => {
    const pages = [A4, LETTER, LANDSCAPE];
    for (const zoom of ZOOMS) for (const ratio of RATIOS) {
      const scale = { rasterScale: ratio * zoom, unitScale: ratio };
      for (const [index, box] of canvasPageStack(pages, { scale, padding: CANVAS_PAGES_PADDING_PX }).entries()) {
        expect(onGrid(box.top * ratio)).toBe(true);
        expect(box.height * ratio).toBeCloseTo(Math.ceil(pages[index]!.height * (ratio * zoom)), 6);
      }
    }
  });

  test('a box never shrinks below its page and grows by less than one device pixel', () => {
    for (const zoom of ZOOMS) for (const ratio of RATIOS) {
      const rasterScale = ratio * zoom;
      const height = canvasPageLength(A4.height, { rasterScale, unitScale: rasterScale });
      expect(height).toBeGreaterThanOrEqual(A4.height);
      expect((height - A4.height) * rasterScale).toBeLessThan(1);
    }
  });

  test('without a scale, boxes keep the exact page size at the zoom', () => {
    const boxes = canvasPageStack([A4, LETTER], { zoom: 1.25, padding: CANVAS_PAGES_PADDING_PX });
    expect(boxes).toEqual([
      { top: 24, width: A4.width * 1.25, height: A4.height * 1.25 },
      { top: 24 + (A4.height * 1.25 + CANVAS_PAGE_GAP_PX), width: 816 * 1.25, height: 1056 * 1.25 },
    ]);
  });

  test('gap and padding round to whole device pixels', () => {
    expect(snapCanvasLength(CANVAS_PAGE_GAP_PX, { rasterScale: 1.1, unitScale: 1.1 }) * 1.1).toBeCloseTo(18, 9);
    expect(snapCanvasLength(CANVAS_PAGE_GAP_PX, { rasterScale: 2, unitScale: 2 })).toBe(CANVAS_PAGE_GAP_PX);
    expect(snapCanvasLength(17.3)).toBe(17.3);
  });

  test('invalid scales fall back to exact boxes', () => {
    expect(canvasPageStack([A4], { scale: { rasterScale: 0, unitScale: Number.NaN } })).toEqual(canvasPageStack([A4]));
  });

  test('canvasPageTops reads the same stack from display list queries', () => {
    const pages = [A4, null, LETTER, A4];
    const queries = { pageCount: () => pages.length, pageSize: (index: number) => pages[index] ?? null } as unknown as DisplayListQueries;
    let top = CANVAS_PAGES_PADDING_PX;
    const expected = pages.map((page) => {
      const current = top;
      top += (page?.height ?? 0) + CANVAS_PAGE_GAP_PX;
      return current;
    });
    expect(canvasPageTops(queries)).toEqual(expected);
  });
});
