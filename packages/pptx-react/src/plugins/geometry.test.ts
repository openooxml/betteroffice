import { describe, expect, test } from 'bun:test';
import type {
  DeckSnapshot,
  ShapeSnapshot,
  SlideDisplayList,
  SlidePrimitive,
} from '@betteroffice/pptx';
import {
  createPluginGeometry,
  EMU_PER_PIXEL,
  pluginLayout,
  slideRectToOverlay,
  type PresentedSlide,
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
