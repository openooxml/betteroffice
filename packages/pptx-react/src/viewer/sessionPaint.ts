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
}

export function frameImages(): {
  resolve(frame: PresentationFrame): PaintImageResolver;
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
    resolve: (frame) => {
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
          entry.promise = decodePresentationImage(bytes, 'Unable to decode slide image').catch(() => null)
            .then((source) => { entry.source = source; close(entry); return source; });
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
          for (const image of held.values()) {
            image.references -= 1;
            close(image);
          }
          held.clear();
        },
      });
    },
    dispose() {
      disposed = true;
      for (const image of cache.values()) {
        image.retained = false;
        close(image);
      }
      cache.clear();
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
