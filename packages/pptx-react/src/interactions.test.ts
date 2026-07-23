import { describe, expect, it } from 'bun:test';
import type { DeckSnapshot, ShapeSnapshot, SlideDisplayList } from '@betteroffice/pptx';
import {
  canMoveShape,
  findShape,
  findTopLevelShape,
  frameBoundsForShape,
  indexShapes,
  movedShapePosition,
  passedDragThreshold,
  slidePoint,
  textPositionAtPoint,
} from './interactions';

const child = shape('child', 20, 30, 40, 50);
const group = { ...shape('group', 10, 20, 100, 100), kind: 'group' as const, children: [child] };
const picture = { ...shape('picture', 100, 200, 300, 400), kind: 'picture' as const };
const deck: DeckSnapshot = {
  widthEmu: 12_192_000,
  heightEmu: 6_858_000,
  slides: [{ id: 'slide', sourcePartPath: null, layoutPartPath: null, name: null, shapes: [group, picture] }],
};
const frame: SlideDisplayList = {
  contractVersion: 1,
  width: 1280,
  height: 720,
  primitives: [
    { kind: 'image', objectId: 1, shapeId: 'child', name: 'child', x: 40, y: 50, w: 80, h: 90 },
    {
      kind: 'image',
      objectId: 2,
      shapeId: 'picture',
      name: 'picture',
      x: 200,
      y: 210,
      w: 120,
      h: 130,
      transform: { rotationDeg: 90 },
    },
    {
      kind: 'textBox',
      objectId: 3,
      shapeId: 'text',
      storyId: 'story',
      x: 100,
      y: 100,
      w: 300,
      h: 100,
      anchor: 'top',
      paragraphs: [],
      lines: [
        {
          x: 110,
          y: 110,
          width: 100,
          height: 20,
          baseline: 125,
          start: 0,
          end: 5,
          runs: [],
          caretStops: [
            { position: 0, x: 110 },
            { position: 5, x: 210 },
          ],
        },
        {
          x: 110,
          y: 140,
          width: 100,
          height: 20,
          baseline: 155,
          start: 6,
          end: 10,
          runs: [],
          caretStops: [
            { position: 6, x: 110 },
            { position: 10, x: 210 },
          ],
        },
      ],
    },
  ],
};

describe('pptx interactions', () => {
  it('maps client coordinates into slide coordinates', () => {
    expect(slidePoint({ left: 200, top: 100, width: 640, height: 360 }, frame, 520, 280)).toEqual({
      x: 640,
      y: 360,
    });
    expect(slidePoint({ left: 0, top: 0, width: 0, height: 360 }, frame, 0, 0)).toBeNull();
  });

  it('resolves descendants to their movable top-level shape', () => {
    expect(findShape(deck.slides[0].shapes, 'child')?.id).toBe('child');
    expect(findTopLevelShape(deck.slides[0], 'child')?.id).toBe('group');
    expect(findTopLevelShape(deck.slides[0], 'missing')).toBeNull();
    expect([...indexShapes(deck.slides[0].shapes).keys()].sort()).toEqual([
      'child',
      'group',
      'picture',
    ]);
  });

  it('only moves shapes with a local transform', () => {
    expect(canMoveShape(picture)).toBe(true);
    expect(canMoveShape({ ...picture, width: 0, height: 0 })).toBe(false);
  });

  it('uses descendant primitive bounds for a group', () => {
    expect(frameBoundsForShape(deck, frame, group)).toEqual({ x: 40, y: 50, width: 80, height: 90 });
  });

  it('includes primitive rotation in selection bounds', () => {
    const bounds = frameBoundsForShape(deck, frame, picture);
    expect(bounds?.x).toBeCloseTo(195);
    expect(bounds?.y).toBeCloseTo(215);
    expect(bounds?.width).toBeCloseTo(130);
    expect(bounds?.height).toBeCloseTo(120);
  });

  it('converts one final frame delta to absolute EMU coordinates', () => {
    expect(movedShapePosition(deck, frame, picture, { x: 0, y: 200 })).toEqual({
      x: 100,
      y: 1_905_200,
    });
  });

  it('uses a client-pixel drag threshold', () => {
    expect(passedDragThreshold(10, 10, 12, 12)).toBe(false);
    expect(passedDragThreshold(10, 10, 14, 10)).toBe(true);
    expect(passedDragThreshold(10, 10, 16, 10, 8)).toBe(false);
  });

  it('clamps captured text dragging to the nearest caret', () => {
    expect(textPositionAtPoint(frame, 'text', 'story', { x: -100, y: -100 })).toBe(0);
    expect(textPositionAtPoint(frame, 'text', 'story', { x: 999, y: 999 })).toBe(10);
    expect(textPositionAtPoint(frame, 'missing', 'story', { x: 110, y: 110 })).toBeNull();
  });
});

function shape(id: string, x: number, y: number, width: number, height: number): ShapeSnapshot {
  return {
    id,
    sourceId: 1,
    kind: 'shape',
    name: id,
    x,
    y,
    width,
    height,
    rotationDeg: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    placeholder: null,
    fill: null,
    outline: null,
    mediaPartPath: null,
    graphic: null,
    textStories: [],
    children: [],
  };
}
