import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import type { PresentationHandle, SlideDisplayList } from '@betteroffice/pptx';
import { initWasm, openPresentation } from '@betteroffice/pptx';
import { farEdge, hoverTargetAtPoint } from './interactions';

const root = resolve(import.meta.dir, '../../..');
let handle: PresentationHandle;
let frame: SlideDisplayList;

beforeAll(async () => {
  const [wasm, sourceDeck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  const zip = await JSZip.loadAsync(sourceDeck);
  const slide = zip.file('ppt/slides/slide1.xml');
  if (!slide) throw new Error('fixture slide is missing');
  const xml = (await slide.async('string'))
    .replace(
      '<a:xfrm><a:off x="762000" y="533400"/>',
      '<a:xfrm rot="1020000" flipH="1"><a:off x="762000" y="533400"/>'
    )
    .replace(
      '<a:xfrm><a:off x="762000" y="1219200"/>',
      '<a:xfrm rot="5400000" flipV="1"><a:off x="762000" y="1219200"/>'
    );
  zip.file('ppt/slides/slide1.xml', xml);
  const deck = await zip.generateAsync({ type: 'uint8array' });
  await initWasm(wasm);
  handle = openPresentation(deck, {
    clientId: 9101,
    fonts: [{ family: 'Liberation Sans', bytes: font }],
  });
  frame = handle.layoutSlide(0);
});

afterAll(() => handle.dispose());

describe('hover target parity with the engine', () => {
  test('agrees with hitTest across the slide', () => {
    const disagreements: string[] = [];
    expect(
      frame.primitives.some(
        (primitive) => primitive.transform?.flipH && primitive.transform.rotationDeg === 17
      ) &&
        frame.primitives.some(
          (primitive) => primitive.transform?.flipV && primitive.transform.rotationDeg === 90
        )
    ).toBe(true);
    let checked = 0;
    for (let y = 0; y < frame.height; y += 3) {
      for (let x = 0; x < frame.width; x += 3) {
        checked += 1;
        const hovered = hoverTargetAtPoint(frame, { x, y });
        const engine = handle.hitTest(x, y)?.kind ?? null;
        if (hovered !== engine && disagreements.length < 25) {
          disagreements.push(`(${x},${y}) hover=${hovered} engine=${engine}`);
        }
      }
    }
    expect(checked).toBeGreaterThan(100_000);
    expect(disagreements).toEqual([]);
  });

  test('agrees on every primitive corner', () => {
    const disagreements: string[] = [];
    for (const primitive of frame.primitives) {
      if (!primitive.shapeId) continue;
      // the far edges are f32 sums; the whole-pixel values astride them are the
      // positions a real cursor actually reports, so probe both
      const [x0, y0] = [Math.fround(primitive.x), Math.fround(primitive.y)];
      const [x1, y1] = [
        farEdge(primitive.x, primitive.w),
        farEdge(primitive.y, primitive.h),
      ];
      const corners = [
        [x0, y0],
        [x1, y0],
        [x0, y1],
        [x1, y1],
        [Math.round(x0), Math.round(y0)],
        [Math.round(x1), Math.round(y1)],
        [Math.round(x0), Math.round(y1)],
        [Math.round(x1), Math.round(y0)],
        [x0 + primitive.w / 2, y0 + primitive.h / 2],
      ];
      // sub-pixel offsets straddle the f32 narrowing wasm-bindgen applies to
      // hitTest's arguments, which a whole-pixel probe cannot reach
      for (const [cx, cy] of corners) {
        for (const delta of [0, -1e-8, 1e-8, -1e-6, 1e-6]) {
          const [x, y] = [cx + delta, cy + delta];
          const hovered = hoverTargetAtPoint(frame, { x, y });
          const engine = handle.hitTest(x, y)?.kind ?? null;
          if (hovered !== engine && disagreements.length < 25) {
            disagreements.push(`${primitive.shapeId} (${x},${y}) hover=${hovered} engine=${engine}`);
          }
        }
      }
    }
    expect(disagreements).toEqual([]);
  });
});
