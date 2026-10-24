import type { DocxSelectionPointIntent } from '@betteroffice/docx/yrs';
import type { PeerlessPropertyIntent } from './inputEvents';
import type { PeerlessFrameGeometry } from './viewStore';

const EMU_PER_PX = 9525;
export const PICTURE_DRAG_SLOP_PX = 3;

export interface PictureDragTarget {
  handle: string;
  version: string;
  pageIndex: number;
  floating?: boolean;
  rotation?: number;
  rect: { x: number; y: number; width: number; height: number };
}

export interface PictureDrag {
  target: PictureDragTarget;
  startX: number;
  startY: number;
  dx: number;
  dy: number;
  dragging: boolean;
}

/** The floating picture under a page-local point, from a frame whose object geometry matches its layout. Text over a behind-text picture keeps the press. */
export function floatingPictureAt(geometry: PeerlessFrameGeometry, pageIndex: number, x: number, y: number): PictureDragTarget | null {
  const target = pictureAt(geometry, pageIndex, x, y, true);
  if (!target) return null;
  const { noChangeAspect: _lock, floating: _floating, rotation: _rotation, ...drag } = target;
  return drag;
}

export function inlinePictureAt(geometry: PeerlessFrameGeometry, pageIndex: number, x: number, y: number): PictureDragTarget | null {
  const target = pictureAt(geometry, pageIndex, x, y);
  return target && target.floating !== true ? { ...target, floating: false } : null;
}

export function pictureGeometryMatches(geometry: PeerlessFrameGeometry): boolean {
  const objects = geometry.objects;
  return !!objects && objects.version === geometry.documentVersion && objects.frameEpoch === geometry.frameEpoch &&
    objects.layoutRevision === geometry.layoutRevision;
}

export function selectedPicture(geometry: PeerlessFrameGeometry, handle: string): (PictureDragTarget & { noChangeAspect?: boolean }) | null {
  if (!pictureGeometryMatches(geometry)) return null;
  const object = geometry.objects!.objects.find(object => object.kind === 'image' && object.handle === handle);
  const rect = object?.nested !== true && object?.rects.length === 1 ? object.rects[0] : null;
  return rect && object ? { handle, version: geometry.objects!.version, pageIndex: rect.pageIndex,
    rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }, noChangeAspect: object.noChangeAspect,
    floating: object.floating, rotation: object.rotation } : null;
}

/** The picture whose embed the worker selection covers exactly, whether a click or the keyboard selected it. */
export function pictureInSelection(geometry: PeerlessFrameGeometry): ReturnType<typeof selectedPicture> {
  if (!pictureGeometryMatches(geometry)) return null;
  const object = geometry.objects!.objects.find(object => object.kind === 'image' && object.selected === true);
  return object ? selectedPicture(geometry, object.handle) : null;
}

export function pictureAt(geometry: PeerlessFrameGeometry, pageIndex: number, x: number, y: number, floatingOnly = false): ReturnType<typeof selectedPicture> {
  if (!pictureGeometryMatches(geometry)) return null;
  const objects = geometry.objects!;
  const hit = (rect: { pageIndex: number; x: number; y: number; width: number; height: number }) =>
    rect.pageIndex === pageIndex && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height;
  // A paragraph's selection rects include the full box of each picture anchored in it; those are not text.
  const pictureBox = (rect: { pageIndex: number; x: number; y: number; width: number; height: number }) =>
    objects.objects.some(object => object.kind === 'image' && object.rects.some(box => box.pageIndex === rect.pageIndex &&
      Math.abs(box.x - rect.x) < 0.5 && Math.abs(box.y - rect.y) < 0.5 &&
      Math.abs(box.width - rect.width) < 0.5 && Math.abs(box.height - rect.height) < 0.5));
  let object: (typeof objects.objects)[number] | undefined;
  let topOrder = -Infinity;
  for (const [index, candidate] of objects.objects.entries()) {
    if (candidate.kind !== 'image' || !candidate.rects.some(hit)) continue;
    const order = candidate.paintOrder ?? index;
    if (order >= topOrder) { object = candidate; topOrder = order; }
  }
  if (!object || (floatingOnly && object.floating !== true) || object.nested === true || object.rects.length !== 1) return null;
  if (object.behindText && objects.objects.some(other => other.kind === 'text' && other.rects.some(rect => hit(rect) && !pictureBox(rect)))) return null;
  return selectedPicture(geometry, object.handle);
}

export function beginPictureDrag(target: PictureDragTarget, clientX: number, clientY: number): PictureDrag {
  return { target, startX: clientX, startY: clientY, dx: 0, dy: 0, dragging: false };
}

/** Moves the drag to a client point; `zoom` maps client pixels to page pixels. Dragging starts past the slop. */
export function movePictureDrag(drag: PictureDrag, clientX: number, clientY: number, zoom: number): PictureDrag {
  if (!Number.isFinite(zoom) || zoom <= 0) return drag;
  const clientDx = clientX - drag.startX;
  const clientDy = clientY - drag.startY;
  const dragging = drag.dragging || Math.hypot(clientDx, clientDy) > PICTURE_DRAG_SLOP_PX;
  return { ...drag, dragging, dx: clientDx / zoom, dy: clientDy / zoom };
}

/** A picture move, or null without a drag or a valid inline drop. */
export function pictureDropIntent(drag: PictureDrag, drop: DocxSelectionPointIntent | null = null): PeerlessPropertyIntent | null {
  if (!drag.dragging) return null;
  if (drag.target.floating === false) return drop ? { kind: 'moveEmbed', payload: {
    handle: drag.target.handle, expectVersion: drag.target.version, drop,
  } } : null;
  if (drag.dx === 0 && drag.dy === 0) return null;
  const { rect, handle, version } = drag.target;
  return { kind: 'moveObject', payload: {
    handle, expectVersion: version,
    hOffsetEmu: Math.round((rect.x + drag.dx) * EMU_PER_PX),
    vOffsetEmu: Math.round((rect.y + drag.dy) * EMU_PER_PX),
  } };
}
