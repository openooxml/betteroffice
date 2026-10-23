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

/** The topmost floating picture at a page-local point. */
export function floatingPictureAt(geometry: PeerlessFrameGeometry, pageIndex: number, x: number, y: number): PictureDragTarget | null {
  const objects = geometry.objects;
  if (!objects || objects.version !== geometry.documentVersion || objects.frameEpoch !== geometry.frameEpoch ||
    objects.layoutRevision !== geometry.layoutRevision) return null;
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
  if (!object || object.floating !== true || object.nested === true || object.rects.length !== 1) return null;
  if (object.behindText && objects.objects.some(other => other.kind === 'text' && other.rects.some(rect => hit(rect) && !pictureBox(rect)))) return null;
  const rect = object.rects[0]!;
  return { handle: object.handle, version: objects.version, pageIndex, rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height } };
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
