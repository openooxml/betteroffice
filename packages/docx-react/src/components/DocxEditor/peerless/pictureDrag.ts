import type { PeerlessPropertyIntent } from './inputEvents';
import type { PeerlessFrameGeometry } from './viewStore';

const EMU_PER_PX = 9525;
export const PICTURE_DRAG_SLOP_PX = 3;

export interface PictureDragTarget {
  handle: string;
  version: string;
  pageIndex: number;
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
  const { noChangeAspect: _lock, ...drag } = target;
  return drag;
}

export function pictureGeometryMatches(geometry: PeerlessFrameGeometry): boolean {
  const objects = geometry.objects;
  return !!objects && objects.version === geometry.documentVersion && objects.frameEpoch === geometry.frameEpoch &&
    objects.layoutRevision === geometry.layoutRevision;
}

export function selectedPicture(geometry: PeerlessFrameGeometry, handle: string): (PictureDragTarget & { noChangeAspect?: boolean }) | null {
  if (!pictureGeometryMatches(geometry)) return null;
  const object = geometry.objects!.objects.find(object => object.kind === 'image' && object.handle === handle);
  const rect = object?.rects.length === 1 ? object.rects[0] : null;
  return rect && object ? { handle, version: geometry.objects!.version, pageIndex: rect.pageIndex,
    rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }, noChangeAspect: object.noChangeAspect } : null;
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
  let overText: boolean | undefined;
  for (let index = objects.objects.length - 1; index >= 0; index--) {
    const object = objects.objects[index]!;
    if (object.kind !== 'image' || (floatingOnly && object.floating !== true) || object.rects.length !== 1 || !hit(object.rects[0]!)) continue;
    if (object.behindText) {
      overText ??= objects.objects.some(other => other.kind === 'text' && other.rects.some(rect => hit(rect) && !pictureBox(rect)));
      if (overText) continue;
    }
    return selectedPicture(geometry, object.handle);
  }
  return null;
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

/** The same page-relative move the object dialog applies, or null when the press never became a drag. */
export function pictureDropIntent(drag: PictureDrag): PeerlessPropertyIntent | null {
  if (!drag.dragging || (drag.dx === 0 && drag.dy === 0)) return null;
  const { rect, handle, version } = drag.target;
  return { kind: 'moveObject', payload: {
    handle, expectVersion: version,
    hOffsetEmu: Math.round((rect.x + drag.dx) * EMU_PER_PX),
    vOffsetEmu: Math.round((rect.y + drag.dy) * EMU_PER_PX),
  } };
}
