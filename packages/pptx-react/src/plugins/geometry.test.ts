import { describe, expect, test } from 'bun:test';
import type {
  DeckSnapshot,
  ShapeSnapshot,
  SlideDisplayList,
  SlidePrimitive,
  TextBoxPrimitive,
} from '@betteroffice/pptx';
import type { PptxGeometryTarget, PptxPluginSelection } from './types';
import {
  createPluginGeometry,
  EMU_PER_PIXEL,
  pluginLayout,
  pluginGeometryIndex,
  slideRectToOverlay,
  slideViewport,
  type PresentedSlide,
  type SelectionCaret,
} from './geometry';

function box(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left,
    top,
    width,
    height,
    x: left,
    y: top,
    right: left + width,
    bottom: top + height,
  } as DOMRect;
}

function element(rect: DOMRect, extra: Partial<HTMLElement> = {}): HTMLElement {
  return {
    getBoundingClientRect: () => rect,
    clientLeft: 0,
    clientTop: 0,
    scrollLeft: 0,
    scrollTop: 0,
    isConnected: true,
    ...extra,
  } as unknown as HTMLElement;
}

function shape(id: string, children: ShapeSnapshot[] = []): ShapeSnapshot {
  return { id, name: id, children, textStories: [] } as unknown as ShapeSnapshot;
}

function primitive(shapeId: string, x: number, y: number, w: number, h: number, rotationDeg = 0) {
  return {
    kind: 'shape',
    shapeId,
    x,
    y,
    w,
    h,
    ...(rotationDeg ? { transform: { rotationDeg } } : {}),
  } as unknown as SlidePrimitive;
}

function presented(scale = 0.5): PresentedSlide {
  const frame = {
    width: 1280,
    height: 720,
    primitives: [
      primitive('title', 100, 50, 400, 80),
      primitive('child-a', 600, 300, 100, 50),
      primitive('child-b', 800, 400, 60, 60, 90),
      primitive('rotated', 200, 400, 200, 100, 90),
    ],
  } as unknown as SlideDisplayList;
  const snapshot = {
    widthEmu: 12_192_000,
    heightEmu: 6_858_000,
    slides: [
      {
        id: 'slide:0:256',
        shapes: [
          shape('title'),
          shape('group', [shape('child-a'), shape('child-b')]),
          shape('rotated'),
          shape('hidden'),
        ],
      },
    ],
  } as unknown as DeckSnapshot;
  return { frame, snapshot, slideIndex: 0, version: 'v1', scale };
}

describe('plugin layouts', () => {
  test('exist only while the presented frame shows the current version', () => {
    const slide = presented(0.75);
    const layout = pluginLayout(slide, 'v1');
    expect(layout).toMatchObject({
      version: 'v1',
      slideId: 'slide:0:256',
      slide: 1,
      zoom: 0.75,
      width: 1280,
      height: 720,
    });
    expect(pluginLayout(slide, 'v1')?.id).toBe(layout!.id);
    expect(pluginLayout(presented(0.75), 'v1')?.id).not.toBe(layout!.id);
    expect(pluginLayout(slide, 'v2')).toBeNull();
    expect(pluginLayout(null, 'v1')).toBeNull();
  });
});

describe('overlay coordinates', () => {
  test('follow the measured canvas box inside the layer, not the device pixel ratio', () => {
    const canvas = element(box(140, 90, 640, 360));
    const layer = element(box(100, 60, 800, 500), { clientLeft: 2, clientTop: 3 } as never);
    expect(
      slideRectToOverlay(
        canvas,
        layer,
        { width: 1280, height: 720 },
        {
          x: 100,
          y: 50,
          width: 400,
          height: 80,
        }
      )
    ).toEqual({ x: 38 + 50, y: 27 + 25, width: 200, height: 40 });
    const scrolled = element(box(100, 60, 800, 500), { scrollLeft: 10, scrollTop: 20 } as never);
    expect(
      slideRectToOverlay(
        canvas,
        scrolled,
        { width: 1280, height: 720 },
        {
          x: 0,
          y: 0,
          width: 1,
          height: 1,
        }
      )
    ).toEqual({ x: 50, y: 50, width: 0.5, height: 0.5 });
    expect(
      slideRectToOverlay(
        element(box(0, 0, 0, 0)),
        layer,
        { width: 1280, height: 720 },
        {
          x: 0,
          y: 0,
          width: 1,
          height: 1,
        }
      )
    ).toBeNull();
  });

  test('scale by the displayed frame on nonstandard slide extents', () => {
    const slide = presented();
    const frame = { ...slide.frame, width: 961, height: 541 } as SlideDisplayList;
    const odd = { ...slide, frame };
    const geometry = createPluginGeometry(
      pluginLayout(odd, 'v1')!,
      odd,
      element(box(10, 20, 1922, 1082)),
      element(box(0, 0, 2000, 1200)),
      () => true,
      () => null
    );
    expect(
      geometry.toOverlayRect({
        space: 'slide-emu',
        rect: { x: 9525 * 100, y: 9525 * 10, width: 9525 * 4, height: 9525 * 2 },
      })
    ).toEqual({ x: 210, y: 40, width: 8, height: 4 });
  });

  test('convert slide EMU through 9525 EMU per pixel and report rendered shape bounds', () => {
    const slide = presented();
    const layout = pluginLayout(slide, 'v1')!;
    let current = true;
    const geometry = createPluginGeometry(
      layout,
      slide,
      element(box(0, 0, 640, 360)),
      element(box(0, 0, 640, 360)),
      () => current,
      () => ({ kind: 'shape', shapeId: 'title', slide: 1, slideId: 'slide:0:256' } as never)
    );
    expect(
      geometry.toOverlayRect({
        space: 'slide-emu',
        rect: { x: 100 * EMU_PER_PIXEL, y: 50 * EMU_PER_PIXEL, width: 9525, height: 9525 },
      })
    ).toEqual({ x: 50, y: 25, width: 0.5, height: 0.5 });
    expect(
      geometry.toOverlayRect({ space: 'slide-px', rect: { x: 100, y: 50, width: 2, height: 2 } })
    ).toEqual({ x: 50, y: 25, width: 1, height: 1 });
    expect(
      geometry.toOverlayRect({
        space: 'slide-px',
        rect: { x: Number.NaN, y: 0, width: 1, height: 1 },
      })
    ).toBeNull();

    expect(geometry.getShapeRect('title')).toEqual({ x: 50, y: 25, width: 200, height: 40 });
    expect(geometry.getShapeRect('group')).toEqual({ x: 300, y: 150, width: 130, height: 80 });
    const rotated = geometry.getShapeRect('rotated')!;
    expect([rotated.x, rotated.y, rotated.width, rotated.height].map(Math.round)).toEqual([
      125, 175, 50, 100,
    ]);
    expect(geometry.getShapeRect('hidden')).toBeNull();
    expect(geometry.getShapeRect('missing')).toBeNull();
    expect(geometry.getPositionAtPoint(10, 10)).toMatchObject({
      shapeId: 'title',
      version: 'v1',
      layoutId: layout.id,
    });

    current = false;
    expect(geometry.getShapeRect('title')).toBeNull();
    expect(
      geometry.toOverlayRect({ space: 'slide-px', rect: { x: 0, y: 0, width: 1, height: 1 } })
    ).toBeNull();
    expect(geometry.getPositionAtPoint(10, 10)).toBeNull();
  });
});

test('viewport measures the stage client box relative to the overlay and follows scrolling', () => {
  const stage = element(box(100, 60, 804, 604), { clientLeft: 2, clientTop: 2, clientWidth: 780, clientHeight: 580 });
  const layer = element(box(140, -20, 1280, 720), { clientLeft: 1, clientTop: 1, scrollLeft: 4, scrollTop: 6 });
  expect(slideViewport(stage, layer)).toEqual({ x: -35, y: 87, width: 780, height: 580 });
  layer.getBoundingClientRect = () => box(140, -100, 1280, 720);
  expect(slideViewport(stage, layer)).toEqual({ x: -35, y: 167, width: 780, height: 580 });
});

function textGeometry(textValue = 'abcdefghij', caret: SelectionCaret = {}) {
  const slide = presented();
  const text: TextBoxPrimitive = {
    kind: 'textBox', objectId: 1, shapeId: 'title', storyId: 'story',
    x: 100, y: 50, w: 400, h: 80, anchor: 'top',
    paragraphs: [{ level: 0, runs: [{ text: textValue, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
    lines: [0, 5].map((start, row) => ({
      x: 100, y: 50 + row * 20, width: 50, height: 20, baseline: 65 + row * 20,
      start, end: start + 5, runs: [],
      caretStops: Array.from({ length: 6 }, (_, index) => ({ position: start + index, x: 100 + index * 10 })),
    })),
  };
  slide.frame = { ...slide.frame, primitives: [text] };
  const layer = element(box(10, 20, 640, 360));
  let current = true;
  let selection: PptxPluginSelection = null;
  const layout = pluginLayout(slide, 'v1')!;
  const geometry = createPluginGeometry(layout, slide, element(box(50, 60, 640, 360)), layer,
    () => current, () => null, layer, () => selection, caret);
  const range = { kind: 'range', version: 'v1', slideId: layout.slideId,
    shapeId: 'title', storyId: 'story', anchor: 1, focus: 9 } as const;
  return {
    slide, geometry, range,
    select(target: NonNullable<PptxPluginSelection>['target'] | null) {
      selection = target ? { slideId: layout.slideId, slide: 1, target } : null;
    },
    invalidate() { current = false; },
  };
}

describe('plugin anchor geometry', () => {
  test('maps a two-line text range and its focus-end anchor into overlay pixels', () => {
    const { geometry, range } = textGeometry();
    expect(geometry.getAnchorGeometry(range)).toEqual({
      ok: true,
      rects: [
        { x: 95, y: 65, width: 20, height: 10 },
        { x: 90, y: 75, width: 20, height: 10 },
      ],
      anchor: { x: 90, y: 75, width: 20, height: 10 },
      pageRect: { x: 40, y: 40, width: 640, height: 360 },
      version: 'v1', layoutId: geometry.layout.id,
    });
    expect(geometry.getAnchorGeometry({ ...range, anchor: 9, focus: 1 })).toMatchObject({
      ok: true, anchor: { x: 95, y: 65, width: 20, height: 10 },
    });
  });

  test('wrapped ranges anchor to the selected fragment at the focus end', () => {
    const { geometry, range } = textGeometry();
    const firstLine = { x: 95, y: 65, width: 20, height: 10 };
    expect(geometry.getAnchorGeometry({ ...range, anchor: 1, focus: 5 })).toMatchObject({
      ok: true, rects: [firstLine], anchor: firstLine,
    });
    expect(geometry.getAnchorGeometry({ ...range, anchor: 5, focus: 1 })).toMatchObject({
      ok: true, rects: [firstLine], anchor: firstLine,
    });
    const focusLine = { x: 110, y: 65, width: 5, height: 10 };
    expect(geometry.getAnchorGeometry({ ...range, anchor: 6, focus: 4 })).toMatchObject({
      ok: true,
      rects: [focusLine, { x: 90, y: 75, width: 5, height: 10 }],
      anchor: focusLine,
    });
  });

  test('collapsed text has no rectangles and anchors to its caret', () => {
    const { geometry, range } = textGeometry();
    expect(geometry.getAnchorGeometry({ ...range, anchor: 2, focus: 2 })).toMatchObject({
      ok: true, rects: [], anchor: { x: 100, y: 65, width: 0.5, height: 10 },
    });
  });

  test('text ranges use UTF-16 story offsets', () => {
    const { geometry, range } = textGeometry('a😀bcdefgh');
    expect(geometry.getAnchorGeometry({ ...range, anchor: 1, focus: 3 })).toMatchObject({
      ok: true,
      rects: [{ x: 95, y: 65, width: 10, height: 10 }],
      anchor: { x: 95, y: 65, width: 10, height: 10 },
    });
  });

  test('refuses ranges read at another version', () => {
    const { geometry, range } = textGeometry();
    expect(geometry.getAnchorGeometry({ ...range, version: 'v0' })).toMatchObject({
      ok: false, failure: { code: 'stale-version' },
    });
  });

  test('unknown shapes, stories and invalid offsets are missing targets', () => {
    const { geometry, range } = textGeometry();
    for (const target of [
      { ...range, shapeId: 'unknown' }, { ...range, storyId: 'unknown' },
      { ...range, anchor: -1 }, { ...range, focus: 11 }, { ...range, focus: Number.NaN },
    ]) {
      expect(geometry.getAnchorGeometry(target)).toMatchObject({
        ok: false, failure: { code: 'missing-target' },
      });
    }
  });

  test('selection resolves text, shapes and slides, and refuses no selection', () => {
    const { geometry, range, select } = textGeometry();
    const target = { kind: 'selection' } as const;
    expect(geometry.getAnchorGeometry(target)).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    select({ kind: 'text', shapeId: range.shapeId, storyId: range.storyId, anchor: 9, focus: 1 });
    expect(geometry.getAnchorGeometry(target)).toEqual(geometry.getAnchorGeometry({ ...range, anchor: 9, focus: 1 }));
    select({ kind: 'shape', shapeId: 'title' });
    const rect = geometry.getShapeRect('title');
    expect(rect).not.toBeNull();
    expect(geometry.getAnchorGeometry(target)).toMatchObject({ ok: true, rects: [rect], anchor: rect });
    select({ kind: 'shape', shapeId: 'unknown' });
    expect(geometry.getAnchorGeometry(target)).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    select({ kind: 'slide' });
    const pageRect = { x: 40, y: 40, width: 640, height: 360 };
    expect(geometry.getAnchorGeometry(target)).toMatchObject({ ok: true, rects: [pageRect], anchor: pageRect, pageRect });
  });

  test('retained anchor geometry refuses once its frame is no longer presented', async () => {
    const { geometry, range, invalidate } = textGeometry();
    invalidate();
    for (const target of [range, { kind: 'selection' } as const]) {
      expect(geometry.getAnchorGeometry(target)).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    }
    expect(await geometry.readAnchorGeometries([range])).toEqual([geometry.getAnchorGeometry(range)]);
  });

  test('collapsed selection anchors preserve the editor caret line at a wrap boundary', async () => {
    for (const caret of [{ focusLine: 0 }, { affinity: 'upstream' as const }]) {
      const { geometry, range, select } = textGeometry('abcdefghij', caret);
      select({ kind: 'text', shapeId: range.shapeId, storyId: range.storyId, anchor: 5, focus: 5 });
      const target = { kind: 'selection' } as const;
      const answer = geometry.getAnchorGeometry(target);
      expect(answer).toMatchObject({
        ok: true, rects: [], anchor: { x: 115, y: 65, width: 0.5, height: 10 },
      });
      expect(await geometry.readAnchorGeometries([target])).toEqual([answer]);
      expect(geometry.getAnchorGeometry({ ...range, anchor: 5, focus: 5 })).toMatchObject({
        ok: true, rects: [], anchor: { x: 90, y: 75, width: 0.5, height: 10 },
      });
    }
    const { geometry, range, select } = textGeometry('abcdefghij', { focusLine: 1 });
    select({ kind: 'text', shapeId: range.shapeId, storyId: range.storyId, anchor: 5, focus: 5 });
    expect(geometry.getAnchorGeometry({ kind: 'selection' })).toMatchObject({
      ok: true, rects: [], anchor: { x: 90, y: 75, width: 0.5, height: 10 },
    });
  });

  test('batch reads return the synchronous answers in target order', async () => {
    const { geometry, range, select } = textGeometry();
    select({ kind: 'text', shapeId: range.shapeId, storyId: range.storyId, anchor: 1, focus: 9 });
    const targets: PptxGeometryTarget[] = [
      { kind: 'selection' }, range, { ...range, anchor: 2, focus: 2 },
      { ...range, version: 'old' }, { ...range, shapeId: 'missing' },
    ];
    expect(await geometry.readAnchorGeometries(targets)).toEqual(targets.map((target) => geometry.getAnchorGeometry(target)));
    expect(await geometry.readAnchorGeometries([])).toEqual([]);
  });

  test('reuses the geometry index of a presented frame', () => {
    const { slide } = textGeometry();
    const index = pluginGeometryIndex(slide.frame);
    expect(pluginGeometryIndex(slide.frame)).toBe(index);
    expect(pluginGeometryIndex({ ...slide.frame })).not.toBe(index);
  });
});
