import { decodePresentationImage } from '@betteroffice/pptx';
import type { CanvasImageResolver, PptxFontFace, PresentationFrame } from '@betteroffice/pptx';
import { useRef } from 'react';
import { isTiff } from '../../../../shared/media';

export function useStableFonts(fonts: readonly PptxFontFace[]): readonly PptxFontFace[] {
  const stable = useRef(fonts);
  if (stable.current === fonts) return stable.current;
  if (stable.current.length !== fonts.length || !stable.current.every((left, index) => {
    const right = fonts[index];
    return left.family === right.family && (left.bold ?? false) === (right.bold ?? false) &&
      (left.italic ?? false) === (right.italic ?? false) &&
      (left.bytes === right.bytes || (left.bytes.length === right.bytes.length &&
        left.bytes.every((byte, offset) => byte === right.bytes[offset])));
  })) stable.current = fonts;
  return stable.current;
}

export async function installFonts(
  fonts: readonly PptxFontFace[], installed: FontFace[], current: () => boolean
): Promise<void> {
  if (typeof FontFace === 'undefined') return;
  await Promise.all(fonts.map(async (font) => {
    const face = await new FontFace(font.family, font.bytes.slice().buffer as ArrayBuffer, {
      style: font.italic ? 'italic' : 'normal', weight: font.bold ? '700' : '400',
    }).load();
    if (!current()) return;
    document.fonts.add(face);
    installed.push(face);
  }));
}

type PaintImageResolver = CanvasImageResolver & { release(): void };

interface CachedImage {
  promise: Promise<CanvasImageSource | null>;
  source: CanvasImageSource | null;
  references: number;
  retained: boolean;
  closed: boolean;
  bytes: number;
}

export function frameImages(options: { maxDecodedBytes?: number } = {}): {
  resolve(frame: PresentationFrame): PaintImageResolver;
  resolver(read: (id: string) => Uint8Array | undefined, errorMessage: string): CanvasImageResolver;
  clear(): void;
  dispose(): void;
} {
  const budget = options.maxDecodedBytes ?? 128 * 1024 * 1024;
  if (!Number.isSafeInteger(budget) || budget < 0) throw new Error('invalid decoded image budget');
  const cache = new Map<string, CachedImage>();
  let retainedBytes = 0;
  let disposed = false;
  const close = (image: CachedImage) => {
    if (image.retained || image.references || image.closed || !image.source) return;
    image.closed = true;
    if ('close' in image.source && typeof image.source.close === 'function') image.source.close();
    image.source = null;
  };
  const trim = () => {
    while (cache.size > 25 || retainedBytes > budget) {
      const oldest = cache.entries().next().value;
      if (!oldest) break;
      cache.delete(oldest[0]);
      retainedBytes -= oldest[1].bytes;
      oldest[1].retained = false;
      close(oldest[1]);
    }
  };
  const clear = () => {
    for (const image of cache.values()) {
      image.retained = false;
      close(image);
    }
    cache.clear();
    retainedBytes = 0;
  };
  const acquire = (
    read: (id: string) => Uint8Array | undefined, errorMessage: string, ignoreErrors = false,
    contains?: (id: string) => boolean
  ): PaintImageResolver => {
    const held = new Map<string, CachedImage>();
    let released = false;
    const resolve: CanvasImageResolver = (id) => {
      if (disposed || released || (contains && !contains(id))) return Promise.resolve(null);
      let image = held.get(id) ?? cache.get(id);
      if (!image) {
        const bytes = read(id);
        if (!bytes) return Promise.resolve(null);
        const entry: CachedImage = {
          promise: Promise.resolve(null), source: null, references: 0, retained: true, closed: false, bytes: 0,
        };
        entry.promise = decodePresentationImage(bytes, errorMessage).catch((error: unknown) => {
          if (ignoreErrors) return null;
          throw error;
        })
          .then((source) => {
            entry.source = source;
            if (source) {
              const size = source as { naturalWidth?: number; naturalHeight?: number; width?: number; height?: number };
              const bytes = (size.naturalWidth ?? size.width ?? 0) * (size.naturalHeight ?? size.height ?? 0) * 4;
              entry.bytes = Number.isSafeInteger(bytes) && bytes >= 0 ? bytes : budget + 1;
              if (entry.retained) retainedBytes += entry.bytes;
            }
            trim();
            close(entry);
            return source;
          });
        image = entry;
        cache.set(id, image);
      } else if (cache.get(id) === image) {
        cache.delete(id);
        cache.set(id, image);
      }
      if (!held.has(id)) {
        held.set(id, image);
        image.references += 1;
      }
      trim();
      return image.promise;
    };
    return Object.assign(resolve, {
      release() {
        if (released) return;
        released = true;
        for (const image of held.values()) {
          image.references -= 1;
          close(image);
        }
        held.clear();
      },
    });
  };
  return {
    resolve: (frame) => acquire((id) => {
      const bytes = frame.media.get(id);
      return bytes && !isTiff(bytes) ? bytes : undefined;
    }, 'Unable to decode slide image', true, (id) => frame.media.has(id)),
    resolver: (read, errorMessage) => Object.assign(async (id: string) => {
      const resolve = acquire(read, errorMessage);
      try { return await resolve(id); } finally { resolve.release(); }
    }, { acquire: () => acquire(read, errorMessage) }),
    clear,
    dispose() {
      disposed = true;
      clear();
    },
  };
}

export function currentContext(
  ctx: CanvasRenderingContext2D, current: () => boolean
): CanvasRenderingContext2D {
  const methods = new Map<PropertyKey, (...args: unknown[]) => unknown>();
  return new Proxy(ctx, {
    get(target, key) {
      const value: unknown = Reflect.get(target, key, target);
      if (typeof value !== 'function') return value;
      let method = methods.get(key);
      if (!method) {
        method = (...args: unknown[]) => {
          if (!current()) throw new Error('Paint superseded');
          return Reflect.apply(value, target, args);
        };
        methods.set(key, method);
      }
      return method;
    },
    set(target, key, value) {
      if (!current()) throw new Error('Paint superseded');
      return Reflect.set(target, key, value, target);
    },
  });
}
