import { expect, test } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import type { CommandPainterProbe } from './pptx-command-painter-harness';

test.use({ deviceScaleFactor: 1 });

test.beforeEach(async ({ page }) => {
  await page.goto('/pptx-command-painter.html');
  await page.waitForFunction(() => '__pptxCommandPainter' in window);
});

test('p2_picture_stroke_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await pinnedEvaluate(page, () => {
    const probe: CommandPainterProbe = window.__pptxCommandPainter;
    return probe.pixels({ kind: 'image', objectId: 1, name: 'Picture', x: 100, y: 100, w: 100, h: 100,
      stroke: { color: '#000000', width: 20 } }, { x: 95, y: 150, width: 1, height: 1 });
  });
  expect(result.actual).toEqual([0, 0, 0, 255]);
  expect(result.actual).toEqual(result.reference);
});

test('p2_acute_miter_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await pinnedEvaluate(page, () => window.__pptxCommandPainter.pixels({
    kind: 'shape', objectId: 1, name: 'Triangle', geometry: 'triangle',
    x: 400, y: 100, w: 100, h: 200, stroke: { color: '#000000', width: 20 },
    path: [{ type: 'move', x: 0, y: 1 }, { type: 'line', x: 0.5, y: 0 },
      { type: 'line', x: 1, y: 1 }, { type: 'close' }],
  }, { x: 450, y: 70, width: 1, height: 1 }));
  expect(result.actual).toEqual([0, 0, 0, 255]);
  expect(result.actual).toEqual(result.reference);
});

test('p2_offset_text_overflow_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await pinnedEvaluate(page, () => window.__pptxCommandPainter.pixels({
    kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story',
    x: 20, y: 300, w: 300, h: 40, anchor: 'top', overflow: true, paragraphs: [],
    lines: [{ x: 50, y: 300, width: 400, height: 20, baseline: 315, start: 0, end: 40, caretStops: [],
      runs: [{ text: 'M'.repeat(40), start: 0, end: 40, x: 50, width: 400, fontId: 0,
        fontFamily: 'CommandPainterMono', fontSizePx: 16, bold: false, italic: false, underline: false,
        color: '#000000', glyphs: [] }] }],
  }, { x: 424, y: 300, width: 26, height: 20 }));
  expect(result.reference.some((channel, index) => index % 4 === 0 && channel < 128)).toBe(true);
  expect(result.actual).toEqual(result.reference);
});

const area = { x: 380, y: 180, width: 160, height: 140 };

/** Dark pixels of `pixels` (read over `area`) inside `part`, both in slide units. */
function darkIn(pixels: number[], factor: number, part: { x: number; y: number; width: number; height: number }) {
  const stride = Math.round(area.width * factor);
  let count = 0;
  for (let y = Math.round((part.y - area.y) * factor); y < Math.round((part.y + part.height - area.y) * factor); y += 1)
    for (let x = Math.round((part.x - area.x) * factor); x < Math.round((part.x + part.width - area.x) * factor); x += 1)
      if (pixels[(y * stride + x) * 4] < 128) count += 1;
  return count;
}

for (const dpr of [1, 2]) test.describe(`device pixel ratio ${dpr}`, () => {
  test.use({ deviceScaleFactor: dpr });
  for (const scale of [1, 1.5]) for (const kind of ['shape', 'image'] as const) {
    test(`p1_${kind}_shadow_lands_down_right_in_its_tile_at_dpr_${dpr}_zoom_${scale * 100}`, async ({ page }) => {
      const result = await pinnedEvaluate(page, ({ kind, scale, area }) => {
        const frame = { objectId: 1, name: 'Shadowed', x: 400, y: 200, w: 100, h: 100,
          shadow: { color: '#000000', dx: 12, dy: 9 } };
        const primitive = kind === 'shape'
          ? { ...frame, kind, geometry: 'rect', fill: { kind: 'solid' as const, color: '#ffffff' },
            path: [{ type: 'move' as const, x: 0, y: 0 }, { type: 'line' as const, x: 1, y: 0 },
              { type: 'line' as const, x: 1, y: 1 }, { type: 'line' as const, x: 0, y: 1 }, { type: 'close' as const }] }
          : { ...frame, kind, stroke: { color: '#ffffff', width: 2 } };
        return window.__pptxCommandPainter.pixels(primitive, area, scale);
      }, { kind, scale, area });
      const factor = dpr * scale;
      expect(darkIn(result.actual, factor, { x: 503, y: 215, width: 12, height: 90 })).toBeGreaterThan(0);
      expect(darkIn(result.actual, factor, { x: 415, y: 302, width: 90, height: 10 })).toBeGreaterThan(0);
      expect(darkIn(result.actual, factor, { x: 380, y: 180, width: 18, height: 140 })).toBe(0);
      expect(darkIn(result.actual, factor, { x: 380, y: 180, width: 160, height: 18 })).toBe(0);
      expect(result.actual).toEqual(result.reference);
    });
  }
});
