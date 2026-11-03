import type { BundledFontFace } from '@betteroffice/fonts';
import type { PptxFontFace } from '@betteroffice/pptx';
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
    let prepared: { release(): void } | undefined;
    const current = () => active;
    const release = () => { prepared?.release(); prepared = undefined; };
    owner.current = { identity, release };
    void (async () => {
      try {
        const { loadBundledFonts } = await import('./bundledFontsLoad');
        if (!current()) return;
        const loading = loadBundledFonts(file, stable, () => { if (current()) warnFonts(); });
        prepared = loading;
        const loaded = await loading.fonts;
        if (current()) setResolved({ identity, fonts: loaded });
      } catch {
        release();
        if (current()) { warnFonts(); setResolved({ identity, fonts: bytes }); }
      }
    })();
    return () => { active = false; release(); };
  }, [file, pending, stable, identity, bytes, warnFonts]);

  return { fonts: pending ? resolved?.identity === identity ? resolved.fonts : undefined : bytes, release };
}
