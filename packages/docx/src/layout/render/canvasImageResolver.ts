/**
 * Image resolver for the canvas replay backend, shared by both adapters.
 *
 * v0 display lists carry the flow-block image `src` as the primitive's
 * relId — for embedded media that is a blob:/data: URL minted by the parser.
 * Shape picture fills resolve through the same gate via
 * `fillPaint.pictureSrc`. Only those schemes are decoded; anything else
 * (notably remote http urls from external-mode relationships, or a raw
 * unresolved `rId`) resolves to null so opening a document never triggers a
 * network fetch (the no-zero-click-external-fetch security contract). Decode
 * results are cached per source so repaints reuse the same HTMLImageElement.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { ImageResolver } from './canvasBackend';

export function createCanvasImageResolver(): ImageResolver {
  const cache = new Map<string, Promise<CanvasImageSource | null>>();
  return (relId: string) => {
    if (!relId.startsWith('blob:') && !relId.startsWith('data:')) return null;
    let pending = cache.get(relId);
    if (!pending) {
      pending = new Promise<CanvasImageSource | null>((resolve) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = relId;
      });
      cache.set(relId, pending);
    }
    return pending;
  };
}
