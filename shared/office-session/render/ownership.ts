import type { ClosableBitmap } from './types';

/** @internal Last release closes once; closed bitmaps cannot be retained again. */
export interface BitmapOwnership<B extends ClosableBitmap> {
  retain(bitmap: B, holder: object): void;
  release(bitmap: B, holder: object): void;
  readonly openCount: number;
}

/** @internal Reference-counted ownership shared by presentation and echo. */
export function createBitmapOwnership<B extends ClosableBitmap = ClosableBitmap>(
  close: (bitmap: B) => void = (bitmap) => bitmap.close()
): BitmapOwnership<B> {
  const holders = new WeakMap<B, Set<object>>();
  let openCount = 0;
  return {
    retain(bitmap, holder) {
      const current = holders.get(bitmap);
      if (current) { if (current.size > 0) current.add(holder); }
      else { holders.set(bitmap, new Set([holder])); openCount += 1; }
    },
    release(bitmap, holder) {
      const current = holders.get(bitmap);
      if (current?.delete(holder) && current.size === 0) { openCount -= 1; close(bitmap); }
    },
    get openCount() { return openCount; },
  };
}
