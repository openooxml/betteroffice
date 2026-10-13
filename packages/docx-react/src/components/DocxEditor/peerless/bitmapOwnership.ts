import { closeCommandBitmap } from '@betteroffice/docx/peerless';

const holders = new WeakMap<ImageBitmap, Set<object>>();

/** @internal A staged bitmap stays open while any holder retains it; the last release closes it once. */
export function retainBitmap(bitmap: ImageBitmap, holder: object): void {
  const current = holders.get(bitmap);
  if (current) current.add(holder);
  else holders.set(bitmap, new Set([holder]));
}

/** @internal */
export function releaseBitmap(bitmap: ImageBitmap, holder: object): void {
  const current = holders.get(bitmap);
  if (current?.delete(holder) && current.size === 0) closeCommandBitmap(bitmap);
}
