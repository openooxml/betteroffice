import { describe, expect, test } from 'bun:test';
import { displayListNeedsHostImages } from './canvasPresentation';
import type { DisplayList } from '@betteroffice/docx/layout/render';

type Page = DisplayList['pages'][number];

function textPage(pageIndex: number): Page {
  return {
    pageIndex,
    width: 100,
    height: 100,
    primitives: [
      {
        kind: 'glyphRun',
        glyphs: [{ id: 1, x: 0, y: 10 }],
      } as unknown as Page['primitives'][number],
    ],
  };
}

function imagePrimitive(): Page['primitives'][number] {
  return { kind: 'image', x: 0, y: 0, w: 10, h: 10 } as unknown as Page['primitives'][number];
}

describe('displayListNeedsHostImages', () => {
  test('detects image primitives and text-only lists', () => {
    expect(displayListNeedsHostImages({ pages: [textPage(0)] })).toBe(false);
    const withImage: Page = { ...textPage(1), primitives: [imagePrimitive()] };
    expect(displayListNeedsHostImages({ pages: [textPage(0), withImage] })).toBe(true);
  });

  test('caches per page object across frame publications', () => {
    const retained = textPage(0);
    expect(displayListNeedsHostImages({ pages: [retained] })).toBe(false);
    retained.primitives.push(imagePrimitive());
    expect(displayListNeedsHostImages({ pages: [retained] })).toBe(false);
    const replacement: Page = { ...retained, primitives: [...retained.primitives] };
    expect(displayListNeedsHostImages({ pages: [replacement] })).toBe(true);
  });
});
