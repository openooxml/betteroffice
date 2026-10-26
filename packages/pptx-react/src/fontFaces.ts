import type { PptxFontFace } from '@betteroffice/pptx';
import { useMemo, useRef } from 'react';

export const FONT_FALLBACK_MESSAGE = 'Font faces failed to load; text uses fallback fonts.';

function warnFontFallback(): void { console.warn(FONT_FALLBACK_MESSAGE); }

export function useFontWarning(session: unknown): () => void {
  return useMemo(() => {
    let warned = false;
    return () => {
      if (warned) return;
      warned = true;
      warnFontFallback();
    };
  }, [session]);
}

export async function installFonts(
  fonts: readonly PptxFontFace[], installed: FontFace[], current: () => boolean, report: () => void = warnFontFallback,
): Promise<void> {
  if (!fonts.length) return;
  if (typeof FontFace === 'undefined') return;
  const results = await Promise.allSettled(fonts.map(async (font) => {
    const face = await new FontFace(font.family, font.bytes.slice().buffer as ArrayBuffer, {
      style: font.italic ? 'italic' : 'normal', weight: font.bold ? '700' : '400',
    }).load();
    if (!current()) return;
    document.fonts.add(face);
    installed.push(face);
  }));
  if (current() && results.some((result) => result.status === 'rejected')) {
    report();
  }
}

export async function installBrowserFonts(fonts: readonly PptxFontFace[], current: () => boolean = () => true, report?: () => void): Promise<FontFace[]> {
  const installed: FontFace[] = [];
  if (typeof document !== 'undefined') await installFonts(fonts, installed, current, report);
  return installed;
}

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
