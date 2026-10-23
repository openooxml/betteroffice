import type { PeerlessPropertyIntent } from './inputEvents';
import type { PictureDragTarget } from './pictureDrag';

export type PictureResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export interface PictureResize {
  target: PictureDragTarget & { noChangeAspect?: boolean };
  handle: PictureResizeHandle;
  startX: number;
  startY: number;
  rect: PictureDragTarget['rect'];
}

export function beginPictureResize(target: PictureResize['target'], handle: PictureResizeHandle, clientX: number, clientY: number): PictureResize {
  return { target, handle, startX: clientX, startY: clientY, rect: target.rect };
}

export function movePictureResize(resize: PictureResize, clientX: number, clientY: number, zoom: number, shift: boolean): PictureResize {
  if (!Number.isFinite(zoom) || zoom <= 0) return resize;
  const { handle, target: { rect, noChangeAspect } } = resize;
  const horizontal = handle.includes('w') || handle.includes('e');
  const vertical = handle.includes('n') || handle.includes('s');
  let width = horizontal ? rect.width + (clientX - resize.startX) / zoom * (handle.includes('w') ? -1 : 1) : rect.width;
  let height = vertical ? rect.height + (clientY - resize.startY) / zoom * (handle.includes('n') ? -1 : 1) : rect.height;
  if (horizontal && vertical && (noChangeAspect === true || shift)) {
    const widthScale = width / rect.width;
    const heightScale = height / rect.height;
    const dragScale = Math.abs(widthScale - 1) >= Math.abs(heightScale - 1) ? widthScale : heightScale;
    const scale = Math.max(dragScale, 1 / rect.width, 1 / rect.height);
    width = rect.width * scale;
    height = rect.height * scale;
  } else {
    width = Math.max(1, width);
    height = Math.max(1, height);
  }
  return { ...resize, rect: {
    x: handle.includes('w') ? rect.x + rect.width - width : rect.x,
    y: handle.includes('n') ? rect.y + rect.height - height : rect.y,
    width, height,
  } };
}

export function pictureResizeIntent(resize: PictureResize): PeerlessPropertyIntent | null {
  const widthEmu = Math.round(resize.rect.width * 9525);
  const heightEmu = Math.round(resize.rect.height * 9525);
  if (widthEmu === Math.round(resize.target.rect.width * 9525) && heightEmu === Math.round(resize.target.rect.height * 9525)) return null;
  const position = resize.target.floating && (resize.rect.x !== resize.target.rect.x || resize.rect.y !== resize.target.rect.y) ? {
    hOffsetEmu: Math.round(resize.rect.x * 9525), vOffsetEmu: Math.round(resize.rect.y * 9525),
  } : {};
  return { kind: 'resizeObject', payload: { handle: resize.target.handle, expectVersion: resize.target.version, widthEmu, heightEmu, ...position } };
}
