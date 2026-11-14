import { pptxWorkerRenderClients, type PptxPaintFont, type PptxWorkerRenderClient } from '../../../../shared/office-session/render/pptxWorkerRender';
export { installFonts, installBrowserFonts, useStableFonts } from '../fontFaces';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';

export async function installDefaultFonts(
  fonts: readonly DefaultFontFace[], installed: FontFace[], current: () => boolean, report: () => void,
  allowed: (family: string) => boolean = () => true,
): Promise<DefaultFontFace[]> {
  if (!current() || typeof document === 'undefined' || typeof FontFace === 'undefined' || !document.fonts) return [];
  const results = await Promise.all(fonts.map(async (font) => {
    if (!allowed(font.family)) return undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      const traits = font.ownWeight ? undefined : await import('../fontTraits')
        .then((module) => module.readFontTraits(new Uint8Array(font.bytes))).catch(() => undefined);
      if (!current() || !allowed(font.family)) return;
      const loading = new FontFace(font.family, font.bytes, {
        weight: traits ? String(traits.weight) : font.bold ? '700' : '400',
        style: traits ? traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal' : font.italic ? 'italic' : 'normal',
      }).load();
      const deadline = new Promise<undefined>((resolve) => {
        timer = setTimeout(() => resolve(undefined), 5000);
      });
      const face = await Promise.race([loading, deadline]);
      if (!face) { if (current()) report(); return; }
      if (!current() || !allowed(font.family)) return;
      document.fonts.add(face);
      installed.push(face);
      return font;
    } catch { if (current()) report(); }
    finally { if (timer !== undefined) clearTimeout(timer); }
  }));
  return results.filter((face): face is DefaultFontFace => face !== undefined);
}

export function createPptxPaintFontForwarder() {
  const faces: PptxPaintFont[] = [];
  let client: PptxWorkerRenderClient | undefined;
  let disposed = false;
  return {
    installed(face: PptxPaintFont) {
      if (disposed) return;
      faces.push(face);
      client?.fonts([face]);
    },
    attach(owner: object) {
      if (disposed) return;
      client = pptxWorkerRenderClients.get(owner);
      client?.fonts(faces);
    },
    dispose() { disposed = true; faces.length = 0; client = undefined; },
  };
}
