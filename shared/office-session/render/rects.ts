import type { DeviceRect } from './types';

/** @internal Empty bounds contain no device pixels. */
export function isEmpty(rect: DeviceRect): boolean { return rect.width <= 0 || rect.height <= 0; }

/** @internal Snap fractional bounds outward to whole device pixels. */
export function snapOut(rect: DeviceRect): DeviceRect {
  const x = Math.floor(rect.x), y = Math.floor(rect.y);
  return { x, y, width: isEmpty(rect) ? 0 : Math.ceil(rect.x + rect.width) - x,
    height: isEmpty(rect) ? 0 : Math.ceil(rect.y + rect.height) - y };
}

/** @internal Smallest enclosing bounds, ignoring empty rectangles. */
export function union(a: DeviceRect, b: DeviceRect): DeviceRect {
  if (isEmpty(a)) return { ...b };
  if (isEmpty(b)) return { ...a };
  const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y);
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x,
    height: Math.max(a.y + a.height, b.y + b.height) - y };
}

/** @internal True only when two nonempty bounds overlap in area. */
export function intersects(a: DeviceRect, b: DeviceRect): boolean {
  return !isEmpty(a) && !isEmpty(b) && a.x < b.x + b.width && b.x < a.x + a.width &&
    a.y < b.y + b.height && b.y < a.y + a.height;
}

/** @internal Clip device bounds to a surface's dimensions. */
export function clampToSurface(rect: DeviceRect, width: number, height: number): DeviceRect {
  const x = Math.max(0, Math.min(width, rect.x)), y = Math.max(0, Math.min(height, rect.y));
  return { x, y, width: Math.max(0, Math.min(width, rect.x + rect.width) - x),
    height: Math.max(0, Math.min(height, rect.y + rect.height) - y) };
}
