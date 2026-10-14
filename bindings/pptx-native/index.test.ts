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
    expect(presentation.author).toBe('node');
    expect(presentation.origin).toBe('local');
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

  test('preserves explicit edit context and rejects unknown origins and text enums', async () => {
    const presentation = await openPresentation(fixture, { author: 'editor', origin: 'agent' });
    expect(presentation.author).toBe('editor');
    expect(presentation.origin).toBe('agent');
    await expect((async () => openPresentation(fixture, { origin: 'bogus' }))()).rejects.toThrow();
    expect(() => { presentation.origin = 'bogus'; }).toThrow();
    const slideId = (await presentation.slideIds)[0];
    await expect((async () => presentation.addTextBox(slideId, {
      name: 'Invalid style', rect: { x: 0, y: 0, width: 1000000, height: 1000000 },
      text: 'text', style: { caps: 'bogus' }
    }))()).rejects.toThrow();
  });

  test('requires an explicit comment status', async () => {
    const presentation = await openPresentation(fixture);
    await presentation.setCommentFlavor('modern');
    const comment = await presentation.addComment({
      slideId: (await presentation.slideIds)[0], author: 'editor',
      text: 'comment', created: '2026-01-01T00:00:00Z'
    });
    const before = await presentation.comments;
    await expect((async () => presentation.setCommentStatus(comment.commentId, undefined as unknown as boolean))()).rejects.toThrow();
    expect(await presentation.comments).toEqual(before);
    await presentation.setCommentStatus(comment.commentId, true);
  });
});
