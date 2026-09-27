import type { DeckSnapshot, SlideDisplayList } from '@betteroffice/pptx';
import { findShape, renderedShapeBounds } from '../interactions';
import type { PptxPointPosition } from '../PptxEditor';
import type { PptxPluginGeometry, PptxPluginLayout, PptxPluginRect } from './types';

/** EMU per unzoomed display-list pixel, the scale the renderer applies to shape geometry. */
export const EMU_PER_PIXEL = 9525;

/** A slide frame the canvas has finished painting, with the version it was laid out at. */
export interface PresentedSlide {
  frame: SlideDisplayList;
  snapshot: DeckSnapshot;
  slideIndex: number;
  version: string;
  scale: number;
}

const layoutIds = new WeakMap<PresentedSlide, string>();
let nextLayoutId = 0;

function layoutIdOf(presented: PresentedSlide): string {
  let id = layoutIds.get(presented);
  if (!id) {
    nextLayoutId += 1;
    id = `slide-layout-${nextLayoutId}`;
    layoutIds.set(presented, id);
  }
  return id;
}

/** The layout a presented frame shows, or null unless its pixels show `version`. */
export function pluginLayout(
  presented: PresentedSlide | null,
  version: string | null
): PptxPluginLayout | null {
  if (!presented || version === null || presented.version !== version) return null;
  const slide = presented.snapshot.slides[presented.slideIndex];
  if (!slide) return null;
  return {
    id: layoutIdOf(presented),
    version,
    slideId: slide.id,
    slide: presented.slideIndex + 1,
    zoom: presented.scale,
    width: presented.frame.width,
    height: presented.frame.height,
  };
}

function finiteRect(rect: unknown): rect is PptxPluginRect {
  const value = rect as PptxPluginRect | null | undefined;
  return (
    !!value &&
    Number.isFinite(value.x) &&
    Number.isFinite(value.y) &&
    Number.isFinite(value.width) &&
    Number.isFinite(value.height)
  );
}

/**
 * Converts a rectangle in unzoomed slide pixels into pixels of the unscaled `layer`, measuring
 * where the canvas shows `frame` and at what scale now.
 */
export function slideRectToOverlay(
  canvas: Element,
  layer: HTMLElement,
  frame: { width: number; height: number },
  rect: PptxPluginRect
): PptxPluginRect | null {
  const canvasRect = canvas.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  if (!(frame.width > 0 && frame.height > 0 && canvasRect.width > 0 && canvasRect.height > 0)) {
    return null;
  }
  const scaleX = canvasRect.width / frame.width;
  const scaleY = canvasRect.height / frame.height;
  const originX = canvasRect.left - layerRect.left - layer.clientLeft + layer.scrollLeft;
  const originY = canvasRect.top - layerRect.top - layer.clientTop + layer.scrollTop;
  return {
    x: originX + rect.x * scaleX,
    y: originY + rect.y * scaleY,
    width: rect.width * scaleX,
    height: rect.height * scaleY,
  };
}

export function createPluginGeometry(
  layout: PptxPluginLayout,
  presented: PresentedSlide,
  canvas: Element,
  layer: HTMLElement,
  current: () => boolean,
  pointAt: (frame: SlideDisplayList, clientX: number, clientY: number) => PptxPointPosition | null
): PptxPluginGeometry {
  const toOverlay = (rect: PptxPluginRect) =>
    current() ? slideRectToOverlay(canvas, layer, presented.frame, rect) : null;
  return {
    layout,
    toOverlayRect(input) {
      const rect = input?.rect;
      if (!finiteRect(rect)) return null;
      if (input.space === 'slide-px') return toOverlay(rect);
      if (input.space !== 'slide-emu') return null;
      return toOverlay({
        x: rect.x / EMU_PER_PIXEL,
        y: rect.y / EMU_PER_PIXEL,
        width: rect.width / EMU_PER_PIXEL,
        height: rect.height / EMU_PER_PIXEL,
      });
    },
    getShapeRect(shapeId) {
      const slide = presented.snapshot.slides[presented.slideIndex];
      const shape = slide && typeof shapeId === 'string' ? findShape(slide.shapes, shapeId) : null;
      const bounds = shape ? renderedShapeBounds(presented.frame, shape) : null;
      return bounds ? toOverlay(bounds) : null;
    },
    getPositionAtPoint(clientX, clientY) {
      if (!current()) return null;
      const hit = pointAt(presented.frame, clientX, clientY);
      return hit ? { ...hit, version: layout.version, layoutId: layout.id } : null;
    },
  };
}
