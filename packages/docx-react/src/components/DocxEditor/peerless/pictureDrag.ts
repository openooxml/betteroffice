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
  const objects = geometry.objects;
  if (!objects || objects.version !== geometry.documentVersion || objects.frameEpoch !== geometry.frameEpoch ||
    objects.layoutRevision !== geometry.layoutRevision) return null;
  const hit = (rect: { pageIndex: number; x: number; y: number; width: number; height: number }) =>
    rect.pageIndex === pageIndex && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height;
  let overText: boolean | undefined;
  for (let index = objects.objects.length - 1; index >= 0; index--) {
    const object = objects.objects[index]!;
    if (object.kind !== 'image' || object.floating !== true || object.rects.length !== 1 || !hit(object.rects[0]!)) continue;
    if (object.behindText) {
      overText ??= objects.objects.some(other => other.kind === 'text' && other.rects.some(hit));
      if (overText) continue;
    }
    const rect = object.rects[0]!;
    return { handle: object.handle, version: objects.version, pageIndex, rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height } };
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
