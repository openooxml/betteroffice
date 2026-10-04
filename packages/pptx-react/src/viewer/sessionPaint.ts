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

export function frameImages(): {
  resolve(frame: PresentationFrame): CanvasImageResolver;
  dispose(): void;
} {
  const cache = new Map<string, Promise<CanvasImageSource | null>>();
  let disposed = false;
  return {
    resolve: (frame) => (id) => {
      if (disposed) return Promise.resolve(null);
      const bytes = frame.media.get(id);
      if (!bytes || isTiff(bytes)) return Promise.resolve(null);
      let image = cache.get(id);
      if (!image) {
        image = decodePresentationImage(bytes, 'Unable to decode slide image').catch(() => null);
        cache.set(id, image);
      }
      return image;
    },
    dispose() {
      disposed = true;
      for (const image of cache.values()) void image.then((source) => {
        if (source && 'close' in source && typeof source.close === 'function') source.close();
      });
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
