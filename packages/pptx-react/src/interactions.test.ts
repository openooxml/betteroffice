import { describe, expect, it } from 'bun:test';
import type {
  DeckSnapshot,
  ShapeSnapshot,
  SlideDisplayList,
  TextBoxPrimitive,
} from '@betteroffice/pptx';
import {
  MIN_SHAPE_SIZE_EMU,
  canResizeShape,
  canMoveShape,
  findShape,
  findTopLevelShape,
  frameBoundsForShape,
  gestureOwnsPointer,
  handleAnchor,
  hoverTargetAtPoint,
  indexShapes,
  pointerTargetAtPoint,
  movedShapePosition,
  passedDragThreshold,
  resizeCommitDelta,
  resizeCursor,
  resizedBounds,
  resizedShapeBox,
  resizedShapeBounds,
  slidePoint,
  shapeTargetExists,
  textLocationAtPoint,
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
          start: 5,
          end: 10,
          runs: [],
          caretStops: [
            { position: 5, x: 110 },
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

  it('invalidates a resize target when a remote refresh removes it', () => {
    const target = { slideId: 'slide', shapeId: 'picture' };
    expect(shapeTargetExists(deck.slides[0], target)).toBe(true);
    expect(shapeTargetExists({ ...deck.slides[0], shapes: [group] }, target)).toBe(false);
    expect(shapeTargetExists({ ...deck.slides[0], id: 'other' }, target)).toBe(false);
  });

  it('scopes a resize gesture to its captured pointer', () => {
    expect(gestureOwnsPointer({ pointerId: 7 }, 7)).toBe(true);
    expect(gestureOwnsPointer({ pointerId: 7 }, 8)).toBe(false);
    expect(gestureOwnsPointer(null, 7)).toBe(false);
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

  it('reports the hovered primitive for cursor feedback', () => {
    expect(hoverTargetAtPoint(frame, { x: 300, y: 150 })).toBe('text');
    expect(hoverTargetAtPoint(frame, { x: 50, y: 60 })).toBe('shape');
    expect(hoverTargetAtPoint(frame, { x: 600, y: 600 })).toBeNull();
  });

  it('reads a text box edge as the shape, so the border moves it', () => {
    // the box spans x 100-400, y 100-200; a few px inside any edge grabs it
    expect(pointerTargetAtPoint(frame, { x: 102, y: 150 })).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 398, y: 150 })).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 300, y: 102 })).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 300, y: 198 })).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 300, y: 150 })).toBe('text');
  });

  it('keeps the text-box edge band six screen pixels wide', () => {
    expect(pointerTargetAtPoint(frame, { x: 111, y: 150 }, 0.5)).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 113, y: 150 }, 0.5)).toBe('text');
    expect(pointerTargetAtPoint(frame, { x: 101.4, y: 150 }, 4)).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 101.6, y: 150 }, 4)).toBe('text');
  });

  it('leaves the engine mirror without a border band', () => {
    expect(hoverTargetAtPoint(frame, { x: 102, y: 150 })).toBe('text');
    expect(pointerTargetAtPoint(frame, { x: 50, y: 60 })).toBe('shape');
    expect(pointerTargetAtPoint(frame, { x: 600, y: 600 })).toBeNull();
  });

  it('prefers the topmost primitive where they overlap', () => {
    expect(hoverTargetAtPoint(frame, { x: 110, y: 120 })).toBe('text');
  });

  it('respects primitive rotation when hovering', () => {
    expect(hoverTargetAtPoint(frame, { x: 198, y: 275 })).toBe('shape');
    expect(hoverTargetAtPoint(frame, { x: 260, y: 338 })).toBeNull();
  });

  it('treats a text box with nowhere to put a caret as a movable shape', () => {
    const textBox = frame.primitives[2] as TextBoxPrimitive;
    const only = (patch: Partial<TextBoxPrimitive>): SlideDisplayList => ({
      ...frame,
      primitives: [{ ...textBox, ...patch }],
    });
    expect(hoverTargetAtPoint(only({ storyId: undefined }), { x: 300, y: 150 })).toBe('shape');
    expect(hoverTargetAtPoint(only({ lines: [] }), { x: 300, y: 150 })).toBe('shape');
    expect(
      hoverTargetAtPoint(
        only({ lines: textBox.lines.map((line) => ({ ...line, caretStops: [] })) }),
        { x: 300, y: 150 }
      )
    ).toBe('shape');
  });

  it('resolves a drag inside a rotated text box in its unrotated frame', () => {
    const textBox = frame.primitives[2] as TextBoxPrimitive;
    const rotated: SlideDisplayList = {
      ...frame,
      primitives: [{ ...textBox, transform: { rotationDeg: 90 } }],
    };
    // (285,10) lands on the first caret once the 90° turn is undone; read in
    // slide coordinates it would resolve to the far end of the line instead
    expect(textPositionAtPoint(rotated, 'text', 'story', { x: 285, y: 10 })).toBe(0);
    expect(textPositionAtPoint(frame, 'text', 'story', { x: 285, y: 10 })).toBe(5);
  });

  it('clamps captured text dragging to the nearest caret', () => {
    expect(textPositionAtPoint(frame, 'text', 'story', { x: -100, y: -100 })).toBe(0);
    expect(textPositionAtPoint(frame, 'text', 'story', { x: 999, y: 999 })).toBe(10);
    expect(textPositionAtPoint(frame, 'missing', 'story', { x: 110, y: 110 })).toBeNull();
  });

  it('keeps a shared endpoint on the line under the pointer', () => {
    expect(textLocationAtPoint(frame, 'text', 'story', { x: 110, y: 115 })).toEqual({
      position: 0,
      lineIndex: 0,
    });
    expect(textLocationAtPoint(frame, 'text', 'story', { x: 110, y: 145 })).toEqual({
      position: 5,
      lineIndex: 1,
    });
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
    adjustValues: {},
    placeholder: null,
    fill: null,
    resolvedFillColor: null,
    outline: null,
    resolvedOutlineColor: null,
    mediaPartPath: null,
    graphic: null,
    textStories: [],
    children: [],
  };
}

describe('pptx resize handles', () => {
  const bounds = { x: 100, y: 200, width: 300, height: 100 };

  it('anchors each grip on the selection box', () => {
    expect(handleAnchor(bounds, 'nw')).toEqual({ x: 100, y: 200 });
    expect(handleAnchor(bounds, 'se')).toEqual({ x: 400, y: 300 });
    expect(handleAnchor(bounds, 'n')).toEqual({ x: 250, y: 200 });
    expect(handleAnchor(bounds, 'w')).toEqual({ x: 100, y: 250 });
  });

  it('holds the edge opposite the dragged handle', () => {
    expect(resizedBounds(bounds, 'se', { x: 50, y: 20 }, 0)).toEqual({
      x: 100,
      y: 200,
      width: 350,
      height: 120,
    });
    expect(resizedBounds(bounds, 'nw', { x: 50, y: 20 }, 0)).toEqual({
      x: 150,
      y: 220,
      width: 250,
      height: 80,
    });
    expect(resizedBounds(bounds, 'n', { x: 999, y: 20 }, 0)).toEqual({
      x: 100,
      y: 220,
      width: 300,
      height: 80,
    });
  });

  it('never collapses the box past the minimum', () => {
    const collapsed = resizedBounds(bounds, 'e', { x: -1000, y: 0 }, 8);
    expect(collapsed.width).toBe(8);
    const fromWest = resizedBounds(bounds, 'w', { x: 1000, y: 0 }, 8);
    expect(fromWest.width).toBe(8);
    expect(fromWest.x).toBe(392);
  });

  it('converts a drag into the shape box, in EMU', () => {
    const box = resizedShapeBox(deck, frame, picture, 'se', { x: 128, y: 72 });
    if (!box) throw new Error('shape should be resizable');
    expect(box.width).toBe(picture.width + 1_219_200);
    expect(box.height).toBe(picture.height + 685_800);
    expect(box.x).toBe(picture.x);
  });

  it('uses the same EMU floor for preview and commit', () => {
    const box = resizedShapeBox(deck, frame, picture, 'e', { x: -1_000, y: 0 });
    const preview = resizedShapeBounds(deck, frame, picture, 'e', { x: -1_000, y: 0 });
    if (!box || !preview) throw new Error('shape should be resizable');
    expect(MIN_SHAPE_SIZE_EMU).toBe(8 * (deck.widthEmu / frame.width));
    expect(box.width).toBe(MIN_SHAPE_SIZE_EMU);
    expect(preview.width).toBe((box.width * frame.width) / deck.widthEmu);
    expect(preview.x).toBe((box.x * frame.width) / deck.widthEmu);
  });

  it('projects the committed rectangle exactly into the preview', () => {
    const sizable = { ...picture, width: 3_000_000, height: 2_000_000 };
    const delta = { x: 12.345, y: -6.789 };
    const box = resizedShapeBox(deck, frame, sizable, 'nw', delta);
    const preview = resizedShapeBounds(deck, frame, sizable, 'nw', delta);
    if (!box) throw new Error('shape should be resizable');
    expect(preview).toEqual({
      x: (box.x * frame.width) / deck.widthEmu,
      y: (box.y * frame.height) / deck.heightEmu,
      width: (box.width * frame.width) / deck.widthEmu,
      height: (box.height * frame.height) / deck.heightEmu,
    });
  });

  it('offers neither preview nor commit for a 90-degree shape', () => {
    const rotated = { ...picture, rotationDeg: 90 };
    const preview = resizedShapeBounds(deck, frame, rotated, 'e', { x: 50, y: 0 });
    const commit = resizedShapeBox(deck, frame, rotated, 'e', { x: 50, y: 0 });
    expect(canResizeShape(rotated)).toBe(false);
    expect(preview).toBe(commit);
    expect(commit).toBeNull();
  });

  it('names the cursor each grip carries', () => {
    expect(resizeCursor('nw')).toBe('nwse-resize');
    expect(resizeCursor('ne')).toBe('nesw-resize');
    expect(resizeCursor('n')).toBe('ns-resize');
    expect(resizeCursor('w')).toBe('ew-resize');
  });
});

describe('pptx resize commit', () => {
  const start = { x: 100, y: 100 };
  const tracked = { x: 5, y: 5 };

  it('commits the release position, not the last rendered delta', () => {
    // the pointer travelled to 180,140 after the move React last rendered
    expect(resizeCommitDelta(start, tracked, { x: 180, y: 140 })).toEqual({ x: 80, y: 40 });
  });

  it('falls back to the tracked delta when the release resolves nowhere', () => {
    expect(resizeCommitDelta(start, tracked, null)).toEqual(tracked);
  });
});
