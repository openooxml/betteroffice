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
  const resolved = await Promise.all(fonts.map(async (font) => {
    const traits = await import('./fontTraits').then((module) => module.readFontTraits(font.bytes)).catch(() => undefined);
    const weight = traits ? String(traits.weight) : font.bold ? '700' : '400';
    const style = traits ? traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal' : font.italic ? 'italic' : 'normal';
    return { font, weight, style };
  }));
  const groups = new Map<string, typeof resolved>();
  for (const entry of resolved) {
    const key = JSON.stringify([entry.font.family, entry.weight, entry.style]);
    const group = groups.get(key);
    if (group) group.push(entry);
    else groups.set(key, [entry]);
  }
  const results = await Promise.allSettled([...groups.values()].map(async (candidates) => {
    let failure: unknown;
    for (const { font, weight, style } of candidates) {
      let face: FontFace;
      try {
        face = await new FontFace(font.family, font.bytes.slice().buffer as ArrayBuffer, { style, weight }).load();
      } catch (error) {
        failure = error;
        continue;
      }
      if (!current()) return;
      document.fonts.add(face);
      installed.push(face);
      return;
    }
    throw failure;
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
