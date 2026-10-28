import type { ClosableBitmap, Patch } from './types';

/** @internal Minimal compositor context satisfied by a canvas 2D context. */
export interface PatchContext<B extends ClosableBitmap> {
  save(): void;
  beginPath(): void;
  rect(x: number, y: number, width: number, height: number): void;
  clip(): void;
  globalCompositeOperation: GlobalCompositeOperation;
  drawImage(bitmap: B, x: number, y: number): void;
  restore(): void;
}

/** @internal Replace pixels with clipped partial copies; partial bitmap dimensions must match rect dimensions. */
export function applyPatch<B extends ClosableBitmap>(ctx: PatchContext<B>, patch: Patch<B>): void {
  ctx.save();
  try {
    if (!patch.full) {
      ctx.beginPath();
      ctx.rect(patch.rect.x, patch.rect.y, patch.rect.width, patch.rect.height);
      ctx.clip();
    }
    ctx.globalCompositeOperation = 'copy';
    ctx.drawImage(patch.bitmap, patch.full ? 0 : patch.rect.x, patch.full ? 0 : patch.rect.y);
  } finally { ctx.restore(); }
}
