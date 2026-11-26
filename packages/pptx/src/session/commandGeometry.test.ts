import { expect, test } from 'bun:test';
import type { SlideDisplayList } from '../types';
import { hitGeometry, indexHitGeometry } from './commandGeometry';

function displayListWithEmptyBoxes(): SlideDisplayList {
  return {
    contractVersion: 1,
    width: 960,
    height: 540,
    primitives: [
      {
        kind: 'textBox',
        objectId: 2,
        shapeId: 'shape-body',
        storyId: 'story:shape-body:0',
        placeholderType: 'body',
        x: 64,
        y: 140,
        w: 832,
        h: 350,
        anchor: 'top',
        paragraphs: [{ align: 'left', level: 0, runs: [] }],
        lines: [
          {
            x: 72,
            y: 148,
            width: 0,
            height: 24,
            baseline: 168,
            start: 0,
            end: 0,
            runs: [],
            caretStops: [{ position: 0, x: 72 }],
          },
        ],
        overflow: false,
      },
      {
        kind: 'textBox',
        objectId: 3,
        shapeId: 'shape-decorative',
        storyId: 'story:shape-decorative:0',
        x: 64,
        y: 20,
        w: 300,
        h: 100,
        anchor: 'top',
        paragraphs: [{ align: 'left', level: 0, runs: [] }],
        lines: [
          {
            x: 72,
            y: 28,
            width: 0,
            height: 20,
            baseline: 44,
            start: 0,
            end: 0,
            runs: [],
            caretStops: [{ position: 0, x: 72 }],
          },
        ],
        overflow: false,
      },
    ],
  };
}

test('hitGeometry enters text editing anywhere inside an empty editable placeholder, while border stays shape hit', () => {
  const index = indexHitGeometry(displayListWithEmptyBoxes());

  // Deep inside the empty body placeholder well below line 1 (y=148..172)
  expect(hitGeometry(index, { x: 400, y: 420 }, undefined, 1)).toEqual({
    kind: 'text',
    shapeId: 'shape-body',
    storyId: 'story:shape-body:0',
    position: 0,
    affinity: 'upstream',
  });

  // Edge grab band (within 6px of top border) still selects the shape for moving/deleting
  expect(hitGeometry(index, { x: 400, y: 142 }, undefined, 1)).toEqual({
    kind: 'shape',
    shapeId: 'shape-body',
  });

  // Non-placeholder empty AutoShape below line 1 keeps shape hit semantics
  expect(hitGeometry(index, { x: 200, y: 100 }, undefined, 1)).toEqual({
    kind: 'shape',
    shapeId: 'shape-decorative',
  });
});
