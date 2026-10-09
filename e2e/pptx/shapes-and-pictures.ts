import { expect } from 'bun:test';

import { PNG_BASE64, profiledLayout } from './context';
import type { PptxScenario } from './context';
import type { SlideDisplayList } from '../../packages/pptx/src/types';

const SLIDE = 0;
const RECT = { x: 457_200, y: 457_200, width: 1_828_800, height: 1_371_600 };
const PICTURE = {
  x: 3_200_400,
  y: 457_200,
  width: 1_143_000,
  height: 1_143_000,
};
const PRESETS = ['ellipse', 'roundRect', 'triangle'];

function images(layout: SlideDisplayList): unknown[] {
  return layout.primitives.filter((primitive) => primitive.kind === 'image');
}

export const shapesAndPictures: PptxScenario = {
  name: 'shapes-and-pictures',
  description:
    'Adds three preset shapes and an embedded PNG, restyles fill and stroke, walks the z-order, resizes and removes, and checks the slide layout gains and loses the matching primitives.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const slide = recorder.op('snapshot', () => handle.snapshot()).slides[
      SLIDE
    ];
    const baseShapes = slide.shapes.length;
    const baseImages = images(
      profiledLayout(handle, recorder, 'layoutSlide:before', SLIDE)
    ).length;

    const shapes = PRESETS.map((geometry, index) => {
      const receipt = recorder.op(`addShape:${geometry}`, () =>
        handle.addShape(slide.id, {
          name: `preset ${geometry}`,
          geometry,
          rect: { ...RECT, x: RECT.x + index * 200_000 },
          fill: '#dd8800',
        })
      );
      expect(receipt.slideId).toBe(slide.id);
      return receipt.shapeId;
    });
    const picture = recorder.op('addPicture', () =>
      handle.addPicture(slide.id, {
        name: 'probe',
        rect: PICTURE,
        contentType: 'image/png',
        mediaBase64: PNG_BASE64,
      })
    );

    const withShapes = recorder.op('snapshot:afterAdd', () => handle.snapshot())
      .slides[SLIDE];
    expect(withShapes.shapes.length).toBe(baseShapes + PRESETS.length + 1);
    expect(
      withShapes.shapes.find((shape) => shape.id === shapes[0])
    ).toMatchObject({ geometry: 'ellipse', kind: 'shape' });
    const embedded = withShapes.shapes.find(
      (shape) => shape.id === picture.shapeId
    );
    expect(embedded?.kind).toBe('picture');
    expect(embedded?.mediaPartPath).toBeNull();
    expect(embedded?.pendingMedia?.contentType).toBe('image/png');

    const saved = recorder.op('save:withPicture', () => handle.save());
    const reopened = recorder.op('reopen:withPicture', () => open(saved));
    const stored = recorder
      .op('snapshot:stored', () => reopened.snapshot())
      .slides[SLIDE].shapes.find((shape) => shape.name === 'probe');
    expect(stored?.mediaPartPath).toBeTruthy();
    expect(
      recorder.op('mediaBytes', () =>
        reopened.mediaBytes(stored!.mediaPartPath!)
      ).byteLength
    ).toBeGreaterThan(100);
    expect(
      images(
        recorder.op('layoutSlide:stored', () => reopened.layoutSlide(SLIDE))
      ).length
    ).toBe(baseImages + 1);

    const added = profiledLayout(
      handle,
      recorder,
      'layoutSlide:afterAdd',
      SLIDE
    );
    expect(images(added).length).toBe(baseImages + 1);

    expect(
      recorder.op('setShapeFill', () =>
        handle.setShapeFill(slide.id, shapes[1], '#2244ff')
      ).shapeId
    ).toBe(shapes[1]);
    expect(
      recorder.op('setShapeStroke', () =>
        handle.setShapeStroke(slide.id, shapes[1], {
          color: '#111111',
          widthPt: 3,
        })
      ).shapeId
    ).toBe(shapes[1]);
    const restyled = recorder
      .op('snapshot:restyled', () => handle.snapshot())
      .slides[SLIDE].shapes.find((shape) => shape.id === shapes[1]);
    expect(restyled?.resolvedFillColor?.toLowerCase()).toBe('#2244ff');
    expect(restyled?.resolvedOutlineColor?.toLowerCase()).toBe('#111111');

    const paintOrder = () =>
      recorder
        .op('snapshot:zOrder', () => handle.snapshot())
        .slides[SLIDE].shapes.map((shape) => shape.id);
    expect(
      recorder.op('sendShapeToBack', () =>
        handle.sendShapeToBack(slide.id, shapes[2])
      ).shapeId
    ).toBe(shapes[2]);
    expect(paintOrder()[0]).toBe(shapes[2]);
    expect(
      recorder.op('bringShapeForward', () =>
        handle.bringShapeForward(slide.id, shapes[2])
      ).shapeId
    ).toBe(shapes[2]);
    expect(paintOrder()[1]).toBe(shapes[2]);
    expect(
      recorder.op('bringShapeToFront', () =>
        handle.bringShapeToFront(slide.id, shapes[2])
      ).shapeId
    ).toBe(shapes[2]);
    expect(paintOrder().at(-1)).toBe(shapes[2]);

    expect(
      recorder.op('setShapeRect', () =>
        handle.setShapeRect(slide.id, shapes[0], {
          x: 100_000,
          y: 100_000,
          width: 900_000,
          height: 900_000,
        })
      ).after
    ).toMatchObject({ x: 100_000, width: 900_000 });
    expect(
      recorder.op('removeShape:picture', () =>
        handle.removeShape(slide.id, picture.shapeId)
      ).shapeId
    ).toBe(picture.shapeId);
    expect(
      images(profiledLayout(handle, recorder, 'layoutSlide:afterRemove', SLIDE))
        .length
    ).toBe(baseImages);

    expect(recorder.op('undo:removePicture', () => handle.undo()).applied).toBe(
      true
    );
    expect(
      images(profiledLayout(handle, recorder, 'layoutSlide:afterUndo', SLIDE))
        .length
    ).toBe(baseImages + 1);
    expect(recorder.op('redo:removePicture', () => handle.redo()).applied).toBe(
      true
    );
    expect(
      recorder.op('snapshot:final', () => handle.snapshot()).slides[SLIDE]
        .shapes.length
    ).toBe(baseShapes + PRESETS.length);
  },
};
