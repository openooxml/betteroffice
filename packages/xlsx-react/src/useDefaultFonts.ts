import { useCallback, useLayoutEffect, useRef } from 'react';
import type { DisplayList, WorkbookSession } from '@betteroffice/xlsx';
import type { createDefaultFontPainter } from './defaultFonts';

export const defaultFontPainterModule = { load: () => import('./defaultFonts') };

export function useDefaultFonts(
  owner: unknown, decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined,
  defaultFonts: boolean | undefined, repaint: () => void,
) {
  const repaintRef = useRef(repaint);
  repaintRef.current = repaint;
  const state = useRef<{ painter?: ReturnType<typeof createDefaultFontPainter>; last?: DisplayList }>(undefined);
  useLayoutEffect(() => {
    if (defaultFonts === false || decoder === undefined) return;
    const current: NonNullable<typeof state.current> = {};
    state.current = current;
    void defaultFontPainterModule.load().then(({ createDefaultFontPainter }) => {
      if (state.current !== current) return;
      current.painter = createDefaultFontPainter(decoder, () => repaintRef.current());
      if (current.last) void current.painter.load(current.last);
    }, () => {});
    return () => {
      current.painter?.dispose();
      if (state.current === current) state.current = undefined;
    };
  }, [owner, decoder, defaultFonts]);
  return useCallback((displayList: DisplayList, paint: () => void) => {
    const current = state.current;
    if (current?.painter) current.painter.paint(displayList, paint);
    else {
      paint();
      if (current) current.last = displayList;
    }
  }, []);
}
