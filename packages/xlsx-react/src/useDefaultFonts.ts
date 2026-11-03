import { useCallback, useLayoutEffect, useRef } from 'react';
import type { DisplayList, WorkbookSession } from '@betteroffice/xlsx';
import type { RemoteFontProvider } from '@betteroffice/fonts/remote';
import type { createDefaultFontPainter } from './defaultFonts';

export const defaultFontPainterModule = { load: () => import('./defaultFonts') };

export function useDefaultFonts(
  owner: unknown, decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined,
  defaultFonts: boolean | undefined, repaint: () => void, settled?: () => void,
  remoteFonts?: RemoteFontProvider,
) {
  const remote = JSON.stringify(remoteFonts ?? 'bunny');
  const repaintRef = useRef(repaint);
  repaintRef.current = repaint;
  const settledRef = useRef(settled);
  settledRef.current = settled;
  const state = useRef<{ painter?: ReturnType<typeof createDefaultFontPainter>; last?: DisplayList }>(undefined);
  useLayoutEffect(() => {
    if (defaultFonts === false || decoder === undefined) return;
    const current: NonNullable<typeof state.current> = {};
    state.current = current;
    void defaultFontPainterModule.load().then(({ createDefaultFontPainter }) => {
      if (state.current !== current) return;
      current.painter = createDefaultFontPainter(decoder, () => repaintRef.current(),
        { onSettled: () => settledRef.current?.(), remoteFonts: JSON.parse(remote) as RemoteFontProvider });
      if (current.last) current.painter.adopt(current.last);
    }, () => {
      if (state.current !== current) return;
      state.current = undefined;
      repaintRef.current();
      settledRef.current?.();
    });
    return () => {
      current.painter?.dispose();
      if (state.current === current) state.current = undefined;
    };
  }, [owner, decoder, defaultFonts, remote]);
  return useCallback((displayList: DisplayList, paint: () => void): boolean => {
    const current = state.current;
    if (current?.painter) return current.painter.paint(displayList, paint);
    paint();
    if (!current) return true;
    current.last = displayList;
    return false;
  }, []);
}
