import { expect, test } from 'playwright/test';
import type { CommandPainterProbe } from './pptx-command-painter-harness';

test.use({ deviceScaleFactor: 1 });

test.beforeEach(async ({ page }) => {
  await page.goto('/pptx-command-painter.html');
  await page.waitForFunction(() => '__pptxCommandPainter' in window);
});

test('p2_picture_stroke_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await page.evaluate(() => {
    const probe: CommandPainterProbe = window.__pptxCommandPainter;
    return probe.pixels({ kind: 'image', objectId: 1, name: 'Picture', x: 100, y: 100, w: 100, h: 100,
      stroke: { color: '#000000', width: 20 } }, { x: 95, y: 150, width: 1, height: 1 });
  });
  expect(result.actual).toEqual([0, 0, 0, 255]);
  expect(result.actual).toEqual(result.reference);
});

test('p2_acute_miter_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await page.evaluate(() => window.__pptxCommandPainter.pixels({
    kind: 'shape', objectId: 1, name: 'Triangle', geometry: 'triangle',
    x: 400, y: 100, w: 100, h: 200, stroke: { color: '#000000', width: 20 },
    path: [{ type: 'move', x: 0, y: 1 }, { type: 'line', x: 0.5, y: 0 },
      { type: 'line', x: 1, y: 1 }, { type: 'close' }],
  }, { x: 450, y: 70, width: 1, height: 1 }));
  expect(result.actual).toEqual([0, 0, 0, 255]);
  expect(result.actual).toEqual(result.reference);
});

test('p2_offset_text_overflow_survives_tile_edges_in_pixels', async ({ page }) => {
  const result = await page.evaluate(() => window.__pptxCommandPainter.pixels({
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
