import { expect, test } from 'bun:test';
import type { DisplayList, DisplayPage, ImagePrimitive } from './displayList';
import { findImagePrimitiveAtPoint } from './displayListImages';

function pageWith(images: ImagePrimitive[]): DisplayList {
  return {
    pages: [{ pageIndex: 0, width: 400, height: 400, primitives: images } as DisplayPage],
  };
}

function image(docStart: number, clip?: ImagePrimitive['clipGroup']): ImagePrimitive {
  return {
    kind: 'image',
    relId: `rId${docStart}`,
    x: 100,
    y: 20,
    w: 60,
    h: 40,
    docStart,
    ...(clip ? { clipGroup: clip } : {}),
  };
}

test('an image is found only where its clip paints it', () => {
  const list = pageWith([image(10, { clip: { x: 100, y: 50, w: 60, h: 30 } })]);
  expect(findImagePrimitiveAtPoint(list, 0, 120, 35)).toBeNull();
  expect(findImagePrimitiveAtPoint(list, 0, 120, 55)?.primitive.docStart).toBe(10);
});

test('a clip without a size paints nothing, as the canvas painter draws it', () => {
  const list = pageWith([image(10, { clip: { x: 100, y: 20 } })]);
  expect(findImagePrimitiveAtPoint(list, 0, 120, 35)).toBeNull();
});

test('an image clipped out at the point lets the one beneath it take the point', () => {
  const list = pageWith([image(5), image(10, { clip: { x: 100, y: 50, w: 60, h: 30 } })]);
  expect(findImagePrimitiveAtPoint(list, 0, 120, 35)?.primitive.docStart).toBe(5);
});
