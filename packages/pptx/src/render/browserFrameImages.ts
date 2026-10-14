import { isTiff } from '../../../../shared/media';
import type { CanvasImageResolver } from './canvas';
import { decodeBrowserPresentationImage } from './browserImage';

interface CachedImage {
  promise: Promise<CanvasImageSource | null>;
  source: CanvasImageSource | null;
  references: number;
  retained: boolean;
  closed: boolean;
}

export function browserFrameImages(): {
  resolve(frame: { media: ReadonlyMap<string, Uint8Array> }): CanvasImageResolver & { release(): void };
  dispose(): void;
} {
  const cache = new Map<string, CachedImage>();
  let disposed = false;
  const close = (image: CachedImage) => {
    if (image.retained || image.references || image.closed || !image.source) return;
    image.closed = true;
    if ('close' in image.source && typeof image.source.close === 'function') image.source.close();
    image.source = null;
  };
  return {
    resolve(frame) {
      const held = new Map<string, CachedImage>();
      let released = false;
      const resolve: CanvasImageResolver = (id) => {
        if (disposed || released) return Promise.resolve(null);
        const bytes = frame.media.get(id);
        if (!bytes || isTiff(bytes)) return Promise.resolve(null);
        let image = held.get(id) ?? cache.get(id);
        if (!image) {
          const entry: CachedImage = {
            promise: Promise.resolve(null), source: null, references: 0, retained: true, closed: false,
          };
          entry.promise = decodeBrowserPresentationImage(bytes, 'Unable to decode slide image').catch(() => null)
            .then((source) => { entry.source = source; close(entry); return source; });
          image = entry;
          cache.set(id, image);
        } else if (cache.get(id) === image) {
          cache.delete(id);
          cache.set(id, image);
        }
        if (!held.has(id)) { held.set(id, image); image.references += 1; }
        while (cache.size > 25) {
          const oldest = cache.entries().next().value;
          if (!oldest) break;
          cache.delete(oldest[0]);
          oldest[1].retained = false;
          close(oldest[1]);
        }
        return image.promise;
      };
      return Object.assign(resolve, {
        release() {
          if (released) return;
          released = true;
          for (const image of held.values()) { image.references -= 1; close(image); }
          held.clear();
        },
      });
    },
    dispose() {
      disposed = true;
      for (const image of cache.values()) { image.retained = false; close(image); }
      cache.clear();
    },
  };
}
