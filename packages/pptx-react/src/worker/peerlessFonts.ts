import type { PptxFontFace } from '@betteroffice/pptx';
import { useRef } from 'react';

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
  fonts: readonly PptxFontFace[], installed: FontFace[], current: () => boolean,
): Promise<void> {
  if (typeof FontFace === 'undefined') return;
  await Promise.all(fonts.map(async (font) => {
    const face = await new FontFace(font.family, font.bytes.slice().buffer, {
      style: font.italic ? 'italic' : 'normal', weight: font.bold ? '700' : '400',
    }).load();
    if (!current()) return;
    document.fonts.add(face);
    installed.push(face);
  }));
}
