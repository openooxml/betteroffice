import type { BundledFontFace } from '@betteroffice/fonts';
import type { PptxFontFace } from '@betteroffice/pptx';
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { useFontWarning } from './fontFaces';

/**
 * A bundled face decoded before opening the presentation. It serves the font it substitutes
 * (`face.metricCompatWith`, e.g. Arial for Liberation Sans), or its own family when it substitutes none.
 */
export interface PptxBundledFont {
  face: BundledFontFace;
  baseUrl?: string | URL;
}

type Font = PptxFontFace | PptxBundledFont;
export type ResolvedFontProps<T> = Omit<T, 'fonts'> & { fonts: readonly PptxFontFace[] };

function sameFont(left: Font, right: Font): boolean {
  if ('face' in left || 'face' in right) {
    return 'face' in left && 'face' in right && left.face.file === right.face.file &&
      left.face.family === right.face.family && left.face.weight === right.face.weight &&
      left.face.style === right.face.style && left.baseUrl?.toString() === right.baseUrl?.toString();
  }
  return left.family === right.family && (left.bold ?? false) === (right.bold ?? false) &&
    (left.italic ?? false) === (right.italic ?? false) &&
    (left.bytes === right.bytes || (left.bytes.length === right.bytes.length &&
      left.bytes.every((byte, offset) => byte === right.bytes[offset])));
}

function useStableDescriptors(fonts: readonly Font[]): readonly Font[] {
  const stable = useRef(fonts);
  if (stable.current === fonts) return stable.current;
  if (stable.current.length !== fonts.length || !stable.current.every((left, index) => sameFont(left, fonts[index]!))) {
    stable.current = fonts;
  }
  return stable.current;
}

export interface BundledFontsResolverProps {
  file: Uint8Array | undefined;
  fonts: readonly Font[];
  fallback: ReactNode;
  defaultFonts?: boolean;
  initialSlide?: number;
  onReady?: (api: never) => unknown;
  onError?: (error: Error) => void;
  children: (fonts: readonly PptxFontFace[], onReady: (api: never) => unknown, onError: (error: Error) => void) => ReactNode;
}

/** Loads bundled face descriptors with the prepared session worker's decoder, then renders the editor with byte faces. */
export function BundledFontsResolver({ file, fonts, fallback, defaultFonts, initialSlide, onReady, onError, children }: BundledFontsResolverProps) {
  const stable = useStableDescriptors(fonts);
  const identity = useMemo(() => ({}), [file, stable]);
  const bytes = useMemo(() => stable.filter((font): font is PptxFontFace => !('face' in font)), [stable]);
  const pending = !!file && bytes.length !== stable.length;
  const warnFonts = useFontWarning(identity);
  const owner = useRef<{ identity: object; release(): void } | null>(null);
  const session = useRef({ defaultFonts, initialSlide });
  session.current = { defaultFonts, initialSlide };
  const [resolved, setResolved] = useState<{ identity: object; fonts: readonly PptxFontFace[] } | null>(null);
  const release = useCallback(() => {
    if (owner.current?.identity === identity) owner.current.release();
  }, [identity]);

  useEffect(() => {
    if (!file || !pending) return;
    let active = true;
    let prepared: { release(): void } | undefined;
    const current = () => active;
    const releasePrepared = () => { prepared?.release(); prepared = undefined; };
    owner.current = { identity, release: releasePrepared };
    void (async () => {
      try {
        const { loadBundledFonts } = await import('./bundledFontsLoad');
        if (!current()) return;
        const { defaultFonts, initialSlide } = session.current;
        const loading = loadBundledFonts(file, stable, () => { if (current()) warnFonts(); }, { defaultFonts,
          activeSlideIndex: Number.isInteger(initialSlide) ? Math.max(0, initialSlide! - 1) : 0 });
        prepared = loading;
        const loaded = await loading.fonts;
        if (current()) setResolved({ identity, fonts: loaded });
      } catch {
        releasePrepared();
        if (current()) { warnFonts(); setResolved({ identity, fonts: bytes }); }
      }
    })();
    return () => { active = false; releasePrepared(); };
  }, [file, pending, stable, identity, bytes, warnFonts]);

  const ready = useCallback((api: never) => { release(); return onReady?.(api); }, [onReady, release]);
  const failed = useCallback((error: Error) => { release(); onError?.(error); }, [onError, release]);
  const loaded = pending ? resolved?.identity === identity ? resolved.fonts : undefined : bytes;
  return loaded === undefined ? fallback : children(loaded, ready, failed);
}
