import type { DeckSnapshot, SlideDisplayList } from '@betteroffice/pptx';
import { indexGeometry, selectionGeometry, type CommandGeometryIndex } from '@betteroffice/pptx/worker';
import { findShape, renderedShapeBounds } from '../interactions';
import type { PptxPointPosition } from '../PptxEditor';
import type {
  PptxAnchorGeometryResult,
  PptxGeometryTarget,
  PptxPluginGeometry,
  PptxPluginLayout,
  PptxPluginRect,
  PptxPluginSelection,
} from './types';

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
const geometryIndexes = new WeakMap<SlideDisplayList, CommandGeometryIndex>();
let nextLayoutId = 0;

export function pluginGeometryIndex(frame: SlideDisplayList): CommandGeometryIndex {
  let index = geometryIndexes.get(frame);
  if (!index) {
    index = indexGeometry(frame);
    geometryIndexes.set(frame, index);
  }
  return index;
}

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

export function slideViewport(stage: HTMLElement, layer: HTMLElement): PptxPluginRect {
  const bounds = stage.getBoundingClientRect();
  const overlay = layer.getBoundingClientRect();
  return {
    x: bounds.left + stage.clientLeft - overlay.left - layer.clientLeft + layer.scrollLeft,
    y: bounds.top + stage.clientTop - overlay.top - layer.clientTop + layer.scrollTop,
    width: stage.clientWidth,
    height: stage.clientHeight,
  };
}

export function createAnchorGeometry(
  layout: PptxPluginLayout,
  index: CommandGeometryIndex,
  current: () => boolean,
  toOverlay: (rect: PptxPluginRect) => PptxPluginRect | null,
  getShapeRect: (shapeId: string) => PptxPluginRect | null,
  selection: () => PptxPluginSelection
): Pick<PptxPluginGeometry, 'getAnchorGeometry' | 'readAnchorGeometries'> {
  const failure = (
    code: 'stale-version' | 'missing-target' | 'layout-unavailable',
    message: string
  ): PptxAnchorGeometryResult => ({ ok: false, failure: { code, message } });
  const getAnchorGeometry = (target: PptxGeometryTarget): PptxAnchorGeometryResult => {
    if (!current()) return failure('layout-unavailable', 'The slide layout is no longer presented.');
    if (target.kind === 'range' && target.version !== layout.version) {
      return failure('stale-version', 'The range belongs to a different presentation version.');
    }
    const selected: PptxPluginSelection = target.kind === 'selection' ? selection() : {
      slideId: target.slideId,
      slide: layout.slide,
      target: { ...target, kind: 'text' },
    };
    if (!selected) return failure('missing-target', 'There is no selection.');
    if (selected.slideId !== layout.slideId) {
      return failure('layout-unavailable', 'The target slide is not presented.');
    }
    const pageRect = toOverlay({ x: 0, y: 0, width: layout.width, height: layout.height });
    if (!pageRect) return failure('layout-unavailable', 'The slide has no displayed bounds.');
    const answer = (rects: readonly PptxPluginRect[], anchor: PptxPluginRect): PptxAnchorGeometryResult =>
      ({ ok: true, rects, anchor, pageRect, version: layout.version, layoutId: layout.id });
    const value = selected.target;
    if (value.kind === 'slide') return answer([pageRect], pageRect);
    if (value.kind === 'shape') {
      const rect = getShapeRect(value.shapeId);
      return rect ? answer([rect], rect) : failure('missing-target', 'The shape is not displayed.');
    }
    const story = index.hits.find(({ primitive }) => primitive.kind === 'textBox' &&
      primitive.shapeId === value.shapeId && primitive.storyId === value.storyId);
    if (!story || !Number.isInteger(value.anchor) || !Number.isInteger(value.focus) ||
      Math.min(value.anchor, value.focus) < 0 || Math.max(value.anchor, value.focus) > (story.textLength ?? 0)) {
      return failure('missing-target', 'The text range does not exist.');
    }
    const geometry = selectionGeometry(index, { slideId: selected.slideId, shapeId: value.shapeId,
      storyId: value.storyId, anchor: value.anchor, focus: value.focus, affinity: 'downstream' });
    const rects: PptxPluginRect[] = [];
    if (value.anchor !== value.focus) {
      for (const rect of geometry.selectionRects) {
        const mapped = toOverlay(rect);
        if (!mapped) return failure('layout-unavailable', 'The slide has no displayed bounds.');
        rects.push(mapped);
      }
    }
    const anchor = value.anchor === value.focus
      ? geometry.caret && toOverlay(geometry.caret)
      : rects[value.focus > value.anchor ? rects.length - 1 : 0];
    if (!anchor) return failure('missing-target', 'The text range is not displayed.');
    return answer(rects, anchor);
  };
  return {
    getAnchorGeometry,
    readAnchorGeometries: (targets) => Promise.resolve(targets.map(getAnchorGeometry)),
  };
}

export function createPluginGeometry(
  layout: PptxPluginLayout,
  presented: PresentedSlide,
  canvas: Element,
  layer: HTMLElement,
  current: () => boolean,
  pointAt: (frame: SlideDisplayList, clientX: number, clientY: number) => PptxPointPosition | null,
  stage: HTMLElement = layer,
  selection: () => PptxPluginSelection = () => null
): PptxPluginGeometry {
  const toOverlay = (rect: PptxPluginRect) =>
    current() ? slideRectToOverlay(canvas, layer, presented.frame, rect) : null;
  const getShapeRect = (shapeId: string) => {
    const slide = presented.snapshot.slides[presented.slideIndex];
    const shape = slide && typeof shapeId === 'string' ? findShape(slide.shapes, shapeId) : null;
    const bounds = shape ? renderedShapeBounds(presented.frame, shape) : null;
    return bounds ? toOverlay(bounds) : null;
  };
  return {
    layout,
    ...createAnchorGeometry(layout, pluginGeometryIndex(presented.frame), current, toOverlay, getShapeRect, selection),
    get viewport() { return slideViewport(stage, layer); },
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
    getShapeRect,
    getPositionAtPoint(clientX, clientY) {
      if (!current()) return null;
      const hit = pointAt(presented.frame, clientX, clientY);
      return hit ? { ...hit, version: layout.version, layoutId: layout.id } : null;
    },
  };
}
