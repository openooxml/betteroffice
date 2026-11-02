import type { BundledFontFace } from '@betteroffice/fonts';
import type { PptxFontFace } from '@betteroffice/pptx';
import type { preparePresentationFontDecoder } from '@betteroffice/pptx/worker';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useFontWarning, useStableFonts } from './fontFaces';

/** A bundled face decoded before opening the presentation. */
export interface PptxBundledFont {
  face: BundledFontFace;
  baseUrl?: string | URL;
}

type Font = PptxFontFace | PptxBundledFont;
export type ResolvedFontProps<T> = Omit<T, 'fonts'> & { fonts: readonly PptxFontFace[] };

export function useResolvedFonts(file: Uint8Array | undefined, fonts: readonly Font[]) {
  const stable = useStableFonts(fonts);
  const identity = useMemo(() => ({}), [file, stable]);
  const bytes = useMemo(() => {
    const faces = stable.filter((font): font is PptxFontFace => !('face' in font));
    return faces.length === stable.length ? stable as readonly PptxFontFace[] : faces;
  }, [stable]);
  const pending = !!file && bytes.length !== stable.length;
  const warnFonts = useFontWarning(identity);
  const owner = useRef<{ identity: object; release(): void } | null>(null);
  const [resolved, setResolved] = useState<{ identity: object; fonts: readonly PptxFontFace[] } | null>(null);
  const release = useCallback(() => {
    if (owner.current?.identity === identity) owner.current.release();
  }, [identity]);

  useEffect(() => {
    if (!file || !pending) return;
    let active = true;
    let prepared: ReturnType<typeof preparePresentationFontDecoder> | undefined;
    const current = () => active;
    const release = () => { prepared?.release(); prepared = undefined; };
    owner.current = { identity, release };
    void (async () => {
      try {
        const [{ preparePresentationFontDecoder: prepare }, { loadBundledFontBytes }] = await Promise.all([
          import('@betteroffice/pptx/worker'), import('@betteroffice/fonts'),
        ]);
        if (!current()) return;
        prepared = prepare(file);
        const { decoder } = prepared;
        const loaded = await Promise.all(stable.map(async (font): Promise<PptxFontFace | undefined> => {
          if (!('face' in font)) return font;
          try {
            const bytes = await loadBundledFontBytes(font.face, { baseUrl: font.baseUrl }, { decoder });
            return { family: font.face.family, bold: font.face.weight === 700,
              italic: font.face.style === 'italic', bytes: new Uint8Array(bytes) };
          } catch {
            if (current()) warnFonts();
            return undefined;
          }
        }));
        if (current()) setResolved({ identity, fonts: loaded.filter((font): font is PptxFontFace => font !== undefined) });
      } catch {
        release();
        if (current()) { warnFonts(); setResolved({ identity, fonts: bytes }); }
      }
    })();
    return () => { active = false; release(); };
  }, [file, pending, stable, identity, bytes, warnFonts]);

  return { fonts: pending ? resolved?.identity === identity ? resolved.fonts : undefined : bytes, release };
}
