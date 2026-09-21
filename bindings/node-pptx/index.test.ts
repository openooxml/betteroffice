import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { openPresentation } from './index.js';

const fixture = readFileSync(
  new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)
);
const font = readFileSync(
  new URL('../../crates/pptx-raster/tests/assets/Carlito-Regular.ttf', import.meta.url)
);

describe('@betteroffice/pptx-native', () => {
  test('opens, inspects, renders, and saves a presentation', async () => {
    const presentation = await openPresentation(fixture);
    const snapshot = await presentation.snapshot();

    expect(snapshot.slides.length).toBeGreaterThan(0);
    expect(await presentation.slideCount).toBe(snapshot.slides.length);
    expect(await presentation.slideIds).toEqual(snapshot.slides.map((slide) => slide.id));
    expect((await presentation.slide(0)).id).toBe(snapshot.slides[0].id);
    expect(await presentation.canUndo).toBe(false);
    await Promise.all([
      presentation.registerFont({ family: 'Arial', data: font }),
      presentation.registerFont({ family: 'Calibri', data: font })
    ]);
    const rendered = await presentation.renderSlide(0);

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await presentation.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('guards collaboration methods on standalone presentations', async () => {
    const presentation = await openPresentation(fixture);
    await expect(presentation.encodeStateVector()).rejects.toThrow('collaborative presentation');
  });

  test('rejects invalid shadow pixel limits', async () => {
    const presentation = await openPresentation(fixture);

    await expect(presentation.renderSlide(0, { maxShadowPixels: -1 })).rejects.toThrow(
      'maxShadowPixels must be a non-negative safe integer'
    );
    await expect(presentation.renderSlide(0, { maxShadowPixels: Number.NaN })).rejects.toThrow(
      'maxShadowPixels must be a non-negative safe integer'
    );
  });
});
