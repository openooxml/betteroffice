export { installFonts, installBrowserFonts, useStableFonts } from '../fontFaces';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';

export async function installDefaultFonts(
  fonts: readonly DefaultFontFace[], installed: FontFace[], current: () => boolean, report: () => void,
): Promise<void> {
  if (!current() || typeof document === 'undefined' || typeof FontFace === 'undefined' || !document.fonts) return;
  await Promise.all(fonts.map(async (font) => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      const loading = new FontFace(font.family, font.bytes, {
        weight: font.bold ? '700' : '400', style: font.italic ? 'italic' : 'normal',
      }).load();
      const deadline = new Promise<undefined>((resolve) => {
        timer = setTimeout(() => resolve(undefined), 5000);
      });
      const face = await Promise.race([loading, deadline]);
      if (!face) { if (current()) report(); return; }
      if (!current()) return;
      document.fonts.add(face);
      installed.push(face);
    } catch { if (current()) report(); }
    finally { if (timer !== undefined) clearTimeout(timer); }
  }));
}
