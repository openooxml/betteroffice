/**
 * Image resolver for the canvas replay backend, shared by both adapters.
 *
 * v0 display lists carry the flow-block image `src` as the primitive's
 * relId — for embedded media that is a blob:/data: URL minted by the parser,
 * or a `media:{n}` token naming the package part, whose bytes `media` reads.
 * Shape picture fills resolve through the same gate via
 * `fillPaint.pictureSrc`. Only those sources are decoded; anything else
 * (notably remote http urls from external-mode relationships, or a raw
 * unresolved `rId`) resolves to null so opening a document never triggers a
 * network fetch (the no-zero-click-external-fetch security contract). A token
 * whose part this session cannot read paints a grey box. Decode
 * results are cached per source so repaints reuse the same HTMLImageElement.
 *
 */

import type { ImageResolver } from './canvasBackend';

export interface CanvasImageResolverOptions {
  /** The bytes and media type of a `media:{n}` source, or null. */
  media?: (token: string) => { bytes: Uint8Array; mimeType: string } | null;
  /** Identifies the package `media` reads; a new value decodes tokens anew. */
  mediaScope?: () => number;
}

const MEDIA_TOKEN = /^media:(0|[1-9]\d*)$/;

/** What an image whose part this session cannot read paints: a grey box. */
function missingMedia(): CanvasImageSource | null {
  const canvas =
    typeof OffscreenCanvas !== 'undefined'
      ? new OffscreenCanvas(1, 1)
      : typeof document !== 'undefined'
        ? Object.assign(document.createElement('canvas'), { width: 1, height: 1 })
        : null;
  const context = canvas?.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null
    | undefined;
  if (!canvas || !context) return null;
  context.fillStyle = '#e6e6e6';
  context.fillRect(0, 0, 1, 1);
  return canvas;
}

export function createCanvasImageResolver(
  options: CanvasImageResolverOptions = {}
): ImageResolver {
  const cache = new Map<string, Promise<CanvasImageSource | null>>();
  const { media, mediaScope } = options;
  let scope = mediaScope?.();
  return (relId: string) => {
    const token = media && MEDIA_TOKEN.test(relId);
    if (!token && !relId.startsWith('blob:') && !relId.startsWith('data:')) return null;
    if (mediaScope && mediaScope() !== scope) {
      scope = mediaScope();
      cache.clear();
    }
    let pending = cache.get(relId);
    if (!pending) {
      pending = new Promise<CanvasImageSource | null>((resolve) => {
        let url = relId;
        if (token) {
          const source = media(relId);
          if (!source) {
            resolve(missingMedia());
            return;
          }
          url = URL.createObjectURL(
            new Blob([source.bytes as Uint8Array<ArrayBuffer>], { type: source.mimeType })
          );
        }
        const img = new Image();
        const settle = (image: CanvasImageSource | null) => {
          if (token) URL.revokeObjectURL(url);
          resolve(image);
        };
        img.onload = () => settle(img);
        img.onerror = () => settle(token ? missingMedia() : null);
        img.src = url;
      });
      cache.set(relId, pending);
    }
    return pending;
  };
}
