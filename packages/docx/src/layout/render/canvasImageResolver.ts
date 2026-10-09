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
 * results are cached per source so repaints reuse decoded images.
 *
 */

import type { ImageResolver } from './canvasBackend';

export interface CanvasImageResolverOptions {
  /** The bytes and media type of a `media:{n}` source, or null. */
  media?: (token: string) => { bytes: Uint8Array; mimeType: string } | null;
  /** Identifies the package `media` reads; a new value decodes tokens anew. */
  mediaScope?: () => number;
  /** Maximum cached decoded bytes; defaults to 64 MiB. */
  maxCacheBytes?: number;
}

const MEDIA_TOKEN = /^media:(0|[1-9]\d*)$/;

interface CachedImage {
  pending: Promise<CanvasImageSource | null>;
  bytes: number;
}

function closeImage(image: CanvasImageSource | null | undefined): void {
  if (image && 'close' in image) image.close();
}

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
  const cache = new Map<string, CachedImage>();
  const { media, mediaScope } = options;
  const maxCacheBytes =
    options.maxCacheBytes !== undefined && Number.isFinite(options.maxCacheBytes)
      ? Math.max(0, options.maxCacheBytes)
      : 64 * 1024 * 1024;
  let cachedBytes = 0;
  let scope = mediaScope?.();
  const decode = async (relId: string, token: boolean): Promise<CanvasImageSource | null> => {
    let blob: Blob | undefined;
    if (token) {
      const source = media?.(relId);
      if (!source) return missingMedia();
      blob = new Blob([source.bytes as Uint8Array<ArrayBuffer>], { type: source.mimeType });
    }
    if (typeof Image === 'undefined') {
      try {
        const source = blob ?? (await (await fetch(relId)).blob());
        return await createImageBitmap(source);
      } catch {
        return token ? missingMedia() : null;
      }
    }
    return new Promise<CanvasImageSource | null>((resolve) => {
      const url = blob ? URL.createObjectURL(blob) : relId;
      const img = new Image();
      const settle = (image: CanvasImageSource | null) => {
        if (blob) URL.revokeObjectURL(url);
        resolve(image);
      };
      img.onload = () => settle(img);
      img.onerror = () => settle(token ? missingMedia() : null);
      img.src = url;
    });
  };
  return (relId: string) => {
    const token = !!media && MEDIA_TOKEN.test(relId);
    if (!token && !relId.startsWith('blob:') && !relId.startsWith('data:')) return null;
    const nextScope = mediaScope?.();
    if (nextScope !== scope) {
      scope = nextScope;
      cache.clear();
      cachedBytes = 0;
    }
    const cached = cache.get(relId);
    if (cached) {
      cache.delete(relId);
      cache.set(relId, cached);
      return cached.pending;
    }
    const entry: CachedImage = {
      bytes: 0,
      pending: decode(relId, token).then(
        (image) => {
          if (cache.get(relId) !== entry) {
            closeImage(image);
            return null;
          }
          cache.delete(relId);
          const dimensions = image as {
            naturalWidth?: number;
            naturalHeight?: number;
            width?: number;
            height?: number;
          } | null;
          const bytes = Math.max(
            4,
            (dimensions?.naturalWidth ?? dimensions?.width ?? 1) *
              (dimensions?.naturalHeight ?? dimensions?.height ?? 1) *
              4
          );
          entry.bytes = bytes;
          cache.set(relId, entry);
          cachedBytes += bytes;
          for (const [key, oldest] of cache) {
            if (cachedBytes <= maxCacheBytes) break;
            if (oldest === entry) break;
            if (oldest.bytes === 0) continue;
            cache.delete(key);
            cachedBytes -= oldest.bytes;
          }
          return image;
        },
        (error: unknown) => {
          if (cache.get(relId) === entry) cache.delete(relId);
          throw error;
        }
      ),
    };
    cache.set(relId, entry);
    return entry.pending;
  };
}
