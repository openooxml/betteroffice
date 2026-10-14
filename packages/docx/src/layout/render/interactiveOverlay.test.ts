import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type {
  DecorationPrimitive,
  DisplayPage,
  DisplayPrimitive,
  ImagePrimitive,
  ShapePrimitive,
  TextRunPrimitive,
} from './displayList';
import { buildInteractiveOverlayPage, interactiveOverlayHasTabStops } from './interactiveOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function checkbox(groupId = 'header'): DisplayPrimitive {
  return {
    kind: 'rect', x: 100, y: 20, w: 20, h: 20, fill: '#fff',
    inlineSdtWidget: { kind: 'checkbox', groupId, pos: 1 },
  };
}

function image(extra: Partial<ImagePrimitive> = {}): ImagePrimitive {
  return { kind: 'image', relId: 'rIdImage', x: 100, y: 20, w: 60, h: 40, ...extra };
}

function pageWith(body: DisplayPrimitive[]): DisplayPage {
  return {
    pageIndex: 0, width: 400, height: 400, primitives: body,
    header: {
      rId: 'rIdHeader', kind: 'header', y: 0, height: 80, primitives: [checkbox()],
    },
  };
}

function headerWidget(body: DisplayPrimitive[]): HTMLButtonElement {
  return buildInteractiveOverlayPage(pageWith(body))
    .querySelector<HTMLButtonElement>('.layout-inline-sdt-widget')!;
}

function expectHidden(button: HTMLButtonElement): void {
  expect(button.style.visibility).toBe('hidden');
  expect(button.style.pointerEvents).toBe('none');
  expect(button.tabIndex).toBe(-1);
  expect(button.disabled).toBe(true);
  expect(button.hasAttribute('inert')).toBe(true);
  expect(button.getAttribute('aria-hidden')).toBe('true');
  expect(button.getAttribute('aria-disabled')).toBe('true');
}

function expectActive(button: HTMLButtonElement): void {
  expect(button.style.visibility).not.toBe('hidden');
  expect(button.style.pointerEvents).toBe('auto');
  expect(button.tabIndex).toBe(0);
  expect(button.disabled).toBe(false);
  expect(button.hasAttribute('inert')).toBe(false);
  expect(button.hasAttribute('aria-hidden')).toBe(false);
  expect(button.hasAttribute('aria-disabled')).toBe(false);
}

test('body images hide covered header checkboxes', () => {
  expectHidden(headerWidget([image()]));
});

test('body image occlusion honors primitive bounds and clips', () => {
  expectHidden(headerWidget([image({ clipGroup: { clip: { x: 100, y: 20, w: 60, h: 40 } } })]));
  for (const clip of [
    { x: 100, y: 40, w: 60, h: 20 },
    { x: 120, y: 20, w: 40, h: 40 },
    { x: 100, y: 20, w: 0, h: 40 },
    { x: 100, y: 20, w: 60, h: -1 },
    { x: 100, y: 20, h: 40 },
    { x: 100, y: 20, w: 60 },
    { x: 100, y: 20, w: Infinity, h: 40 },
  ]) {
    expectActive(headerWidget([image({ clipGroup: { clip } })]));
  }
  expectHidden(headerWidget([image({
    clipGroup: { clip: { x: NaN, y: Infinity, w: 200, h: 100 } },
  })]));
  expectActive(headerWidget([image({
    x: 130, clipGroup: { clip: { x: 100, y: 20, w: 60, h: 40 } },
  })]));
});

test('body text occlusion honors horizontal paint slots', () => {
  const text: TextRunPrimitive = {
    kind: 'text', text: 'hello', x: 100, baselineY: 35, width: 60,
    font: '16px sans-serif', color: '#000',
  };
  const runs: DisplayPrimitive[] = [text, {
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000', text: 'hello',
    glyphs: [{ id: 1, x: 100, y: 35, cluster: 0, advance: 60 }],
  }];
  for (const run of runs) {
    expectHidden(headerWidget([run]));
    expectHidden(headerWidget([{ ...run, paintClip: { x: 100, w: 15 } } as DisplayPrimitive]));
    for (const paintClip of [{ x: 100, w: 5 }, { x: 120, w: 40 }, { x: 100, w: 0 }]) {
      expectActive(headerWidget([{ ...run, paintClip } as DisplayPrimitive]));
    }
  }
});

test('only body paint hides lower-layer controls', () => {
  const repeat: DisplayPrimitive = {
    kind: 'rect', x: 100, y: 20, w: 40, h: 20, fill: '#fff',
    sdt: { groupId: 'repeat', sdtType: 'repeatingSection', repeatingItem: true },
  };
  const page = pageWith([checkbox('watermark'), image(), checkbox('body')]);
  page.watermarkPrimitiveCount = 1;
  page.header!.primitives.push(repeat);
  page.footer = {
    rId: 'rIdFooter', kind: 'footer', y: 0, height: 80, primitives: [checkbox('footer')],
  };
  page.noteAreas = [{ y: 20, height: 40, primitives: [checkbox('note')] }];
  const overlay = buildInteractiveOverlayPage(page);
  for (const groupId of ['watermark', 'header', 'footer']) {
    expectHidden(overlay.querySelector<HTMLButtonElement>(
      `.layout-inline-sdt-widget[data-sdt-group-id="${groupId}"]`
    )!);
  }
  const buttons = overlay.querySelectorAll<HTMLButtonElement>('.layout-sdt-repeat-btn');
  expect(buttons.length).toBe(2);
  for (const button of buttons) expectHidden(button);
  expect(overlay.querySelector<HTMLElement>('.layout-sdt-repeat-controls')!.style.pointerEvents)
    .toBe('none');
  for (const groupId of ['body', 'note']) {
    expectActive(overlay.querySelector<HTMLButtonElement>(
      `.layout-inline-sdt-widget[data-sdt-group-id="${groupId}"]`
    )!);
  }
  const watermarkOnly = pageWith([image()]);
  watermarkOnly.watermarkPrimitiveCount = 1;
  expectActive(buildInteractiveOverlayPage(watermarkOnly)
    .querySelector<HTMLButtonElement>('.layout-inline-sdt-widget')!);
});

test('only primitives with paint and positive area occlude controls', () => {
  expectHidden(headerWidget([{ kind: 'rect', x: 100, y: 20, w: 20, h: 20, fill: '#fff' }]));
  expectHidden(headerWidget([{
    kind: 'shape', x: 100, y: 20, w: 20, h: 20, fill: '#fff',
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 20 },
      { type: 'line', x: 120, y: 40 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
  }]));
  for (const primitive of [
    image({ w: 0 }),
    image({ h: 0 }),
    image({ opacity: 0 }),
    { kind: 'rect', x: 100, y: 20, w: 20, h: 20, fill: 'transparent' },
    {
      kind: 'shape', x: 100, y: 20, w: 20, h: 20, geometryPath: [],
      stroke: { color: '#000', width: 1 },
    },
  ] as DisplayPrimitive[]) {
    expectActive(headerWidget([primitive]));
  }
});

test('a filled body shape hides header controls only where its path is its own box', () => {
  const box = (extra: Partial<ShapePrimitive> = {}): ShapePrimitive => ({
    kind: 'shape', x: 100, y: 20, w: 20, h: 20, fill: '#fff',
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 20 },
      { type: 'line', x: 120, y: 40 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
    ...extra,
  });
  expectHidden(headerWidget([box({ transform: { rotation: 180 } })]));
  const triangle: ShapePrimitive = {
    kind: 'shape', x: 0, y: 0, w: 200, h: 200, fill: '#fff',
    geometryPath: [
      { type: 'move', x: 0, y: 0 }, { type: 'line', x: 200, y: 200 },
      { type: 'line', x: 0, y: 200 }, { type: 'close' },
    ],
  };
  const bowtie = box({
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 40 },
      { type: 'line', x: 120, y: 20 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
  });
  // Its diagonals move less than the corner tolerance.
  const narrowBowtie = box({
    x: 109.995, w: 0.011,
    geometryPath: [
      { type: 'move', x: 109.997, y: 20 }, { type: 'line', x: 110.006, y: 40 },
      { type: 'line', x: 110.006, y: 20 }, { type: 'line', x: 109.997, y: 40 }, { type: 'close' },
    ],
  });
  // Its left side leans by a hundredth of a pixel, past the control's center.
  const skewed = box({
    x: 109.998, w: 0.004,
    geometryPath: [
      { type: 'move', x: 109.995, y: 20 }, { type: 'line', x: 110.011, y: 20 },
      { type: 'line', x: 110.011, y: 40 }, { type: 'line', x: 110.007, y: 40 }, { type: 'close' },
    ],
  });
  // A rectangle filling the left half of its box, over the control's center until flipped.
  const half = (transform?: ShapePrimitive['transform']) =>
    box({
      x: 105, w: 20, transform,
      geometryPath: [
        { type: 'move', x: 105, y: 20 }, { type: 'line', x: 115, y: 20 },
        { type: 'line', x: 115, y: 40 }, { type: 'line', x: 105, y: 40 }, { type: 'close' },
      ],
    });
  expectHidden(headerWidget([half()]));
  expectHidden(headerWidget([half({ flipH: true, rotation: 180 })]));
  for (const shape of [
    triangle,
    bowtie,
    narrowBowtie,
    skewed,
    box({ transform: { rotation: 45 } }),
    half({ flipH: true }),
    half({ rotation: 180 }),
  ]) {
    expectActive(headerWidget([shape]));
  }
});

test('a stretched picture fill inset from a side hides no header control', () => {
  const picture = (fillPaint: ShapePrimitive['fillPaint']): ShapePrimitive => ({
    kind: 'shape', x: 100, y: 20, w: 20, h: 20, fillPaint,
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 20 },
      { type: 'line', x: 120, y: 40 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
  });
  expectActive(headerWidget([
    picture({ kind: 'picture', pictureRelId: 'rId9', pictureStretchRect: { left: 0.75 } }),
  ]));
  expectHidden(headerWidget([
    picture({ kind: 'picture', pictureRelId: 'rId9', pictureStretchRect: { left: -0.1, top: 0 } }),
  ]));
  expectActive(headerWidget([
    picture({
      kind: 'picture', pictureRelId: 'rId9', pictureFillMode: 'tile',
      pictureStretchRect: { left: 0.75 },
    }),
  ]));
});

test('a picture cropped past its source hides no header control', () => {
  const picture = (fillPaint: ShapePrimitive['fillPaint']): ShapePrimitive => ({
    kind: 'shape', x: 100, y: 20, w: 20, h: 20, fillPaint,
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 20 },
      { type: 'line', x: 120, y: 40 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
  });
  const crop = (left: number, right: number) => ({ top: 0, right, bottom: 0, left });
  expectActive(headerWidget([
    picture({ kind: 'picture', pictureRelId: 'rId9', pictureSrcRect: { left: -1 } }),
  ]));
  expectActive(headerWidget([image({ crop: crop(-0.5, 0) })]));
  expectActive(headerWidget([image({ crop: crop(0.6, 0.4) })]));
  expectHidden(headerWidget([
    picture({ kind: 'picture', pictureRelId: 'rId9', pictureSrcRect: { left: 0.25, right: 0.25 } }),
  ]));
  expectActive(headerWidget([
    picture({
      kind: 'picture', pictureRelId: 'rId9', pictureFillMode: 'tile',
      pictureSrcRect: { left: -1 },
    }),
  ]));
  expectHidden(headerWidget([image({ crop: crop(0.25, 0.25) })]));
});

test('a body image hides header controls only inside the rectangle it paints', () => {
  for (const extra of [
    { shapeType: 'ellipse' },
    { rotationDeg: 45 },
    { rotationDeg: 90 },
    { contentFrame: { x: 130, y: 20, w: 30, h: 40 } },
  ] as Partial<ImagePrimitive>[]) {
    expectActive(headerWidget([image(extra)]));
  }
  for (const extra of [
    { shapeType: 'rect' },
    { rotationDeg: 180 },
    { contentFrame: { x: 100, y: 20, w: 30, h: 40 } },
  ] as Partial<ImagePrimitive>[]) {
    expectHidden(headerWidget([image(extra)]));
  }
});

test('a body decoration hides header controls only where it fills its box', () => {
  const decoration = (extra: Partial<DecorationPrimitive>): DecorationPrimitive => ({
    kind: 'decoration', deco: 'highlight', x: 100, y: 20, w: 60, h: 40, color: '#ffff00', ...extra,
  });
  for (const extra of [
    { dashed: true },
    { dotted: true },
    { style: 'double' },
    { color: 'transparent' },
    { color: '' },
  ] as Partial<DecorationPrimitive>[]) {
    expectActive(headerWidget([decoration(extra)]));
  }
  expectHidden(headerWidget([decoration({})]));
  expectHidden(headerWidget([decoration({ style: 'solid' })]));
});

test('body picture fills with zero opacity do not hide header controls', () => {
  const shape: ShapePrimitive = {
    kind: 'shape', x: 100, y: 20, w: 20, h: 20,
    geometryPath: [
      { type: 'move', x: 100, y: 20 }, { type: 'line', x: 120, y: 20 },
      { type: 'line', x: 120, y: 40 }, { type: 'line', x: 100, y: 40 }, { type: 'close' },
    ],
    fillPaint: { kind: 'picture', pictureRelId: 'rIdPicture', pictureOpacity: 0 },
  };
  expectHidden(headerWidget([{
    ...shape, fillPaint: { ...shape.fillPaint, pictureOpacity: 1 },
  }]));
  expectActive(headerWidget([shape]));
});

test('whitespace-only body text does not hide header controls', () => {
  for (const text of [' ', '\t\n', '\u00a0\u2003']) {
    const runs: DisplayPrimitive[] = [{
      kind: 'text', text, x: 100, baselineY: 35, width: 60,
      font: '16px sans-serif', color: '#000',
    }, {
      kind: 'glyphRun', fontId: 1, size: 16, color: '#000', text,
      glyphs: [{ id: 1, x: 100, y: 35, cluster: 0, advance: 60 }],
    }];
    for (const run of runs) expectActive(headerWidget([run]));
  }
});

test('body text with no fill does not hide header controls', () => {
  const none = { modernEffects: { textFill: { kind: 'none' as const } } };
  const runs: DisplayPrimitive[] = [{
    kind: 'text', text: 'covered', x: 100, baselineY: 35, width: 60,
    font: '16px sans-serif', color: '#000',
  }, {
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000', text: 'covered',
    glyphs: [{ id: 1, x: 100, y: 35, cluster: 0, advance: 60 }],
  }];
  for (const run of runs) {
    expectHidden(headerWidget([run]));
    expectActive(headerWidget([{ ...run, ...none } as DisplayPrimitive]));
    expectActive(headerWidget([{ ...run, ...none, textOutline: true } as DisplayPrimitive]));
  }
});

test('turned or compressed body text does not hide header controls', () => {
  const runs: DisplayPrimitive[] = [{
    kind: 'text', text: 'covered', x: 100, baselineY: 35, width: 60,
    font: '16px sans-serif', color: '#000',
  }, {
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000', text: 'covered',
    glyphs: [{ id: 1, x: 100, y: 35, cluster: 0, advance: 60 }],
  }];
  for (const run of runs) {
    expectActive(headerWidget([{ ...run, rotationDeg: 90 } as DisplayPrimitive]));
    expectActive(headerWidget([{ ...run, horizontalScale: 25 } as DisplayPrimitive]));
    expectHidden(headerWidget([{ ...run, rotationDeg: 360 } as DisplayPrimitive]));
    expectHidden(headerWidget([{ ...run, horizontalScale: 150 } as DisplayPrimitive]));
  }
});

test('a page whose only inline control is covered has no tab stop', () => {
  const page = pageWith([image()]);
  expectHidden(buildInteractiveOverlayPage(page)
    .querySelector<HTMLButtonElement>('.layout-inline-sdt-widget')!);
  expect(interactiveOverlayHasTabStops(page)).toBe(false);
  page.primitives = [];
  expect(interactiveOverlayHasTabStops(page)).toBe(true);
  expectActive(buildInteractiveOverlayPage(page)
    .querySelector<HTMLButtonElement>('.layout-inline-sdt-widget')!);
});

test('covered boundary controls do not contribute tab stops', () => {
  for (const sdtType of ['checkbox', 'repeatingSection']) {
    const page = pageWith([image()]);
    page.header!.primitives = [{
      kind: 'rect', x: 100, y: 20, w: 60, h: 40, fill: '#fff',
      sdt: { groupId: 'header', sdtType, repeatingItem: sdtType === 'repeatingSection' },
    }];
    const buttons = buildInteractiveOverlayPage(page).querySelectorAll<HTMLButtonElement>('button');
    expect(buttons.length).toBe(sdtType === 'checkbox' ? 1 : 2);
    for (const button of buttons) expectHidden(button);
    expect(interactiveOverlayHasTabStops(page)).toBe(false);
    page.primitives = [];
    expect(interactiveOverlayHasTabStops(page)).toBe(true);
  }
});

test('boundary controls use their own centers for occlusion', () => {
  const page = pageWith([image({ x: 140, y: 42, w: 20, h: 16 })]);
  page.header!.primitives = [{
    kind: 'rect', x: 100, y: 20, w: 60, h: 40, fill: '#fff',
    sdt: { groupId: 'repeat', sdtType: 'checkbox', repeatingItem: true },
  }];
  const overlay = buildInteractiveOverlayPage(page);
  expectHidden(overlay.querySelector<HTMLButtonElement>('[data-sdt-repeat="remove"]')!);
  expectActive(overlay.querySelector<HTMLButtonElement>('[data-sdt-repeat="add"]')!);
  expectActive(overlay.querySelector<HTMLButtonElement>('.layout-sdt-widget')!);
  expect(interactiveOverlayHasTabStops(page)).toBe(true);

  page.primitives = [image({ x: 124, y: 36, w: 12, h: 8 })];
  const centerCovered = buildInteractiveOverlayPage(page);
  for (const button of centerCovered.querySelectorAll<HTMLButtonElement>('button')) {
    expectActive(button);
  }

  page.primitives = [image({ x: 140, y: 22, w: 18, h: 18 })];
  expectHidden(buildInteractiveOverlayPage(page)
    .querySelector<HTMLButtonElement>('.layout-sdt-widget')!);
});
