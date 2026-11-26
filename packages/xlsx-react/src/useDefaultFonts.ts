import { useCallback, useLayoutEffect, useMemo, useRef } from 'react';
import type { DisplayList, WorkbookSession, XlsxFontFace, XlsxPngFont, XlsxPngTarget } from '@betteroffice/xlsx';
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
  const state = useRef<{ owner: unknown; painter?: ReturnType<typeof createDefaultFontPainter>; last?: DisplayList }>(undefined);
  const openingFrame = useRef<DisplayList>(undefined);
  useLayoutEffect(() => {
    if (defaultFonts === false || decoder === undefined) return;
    const current: NonNullable<typeof state.current> = { owner };
    if (owner != null) {
      current.last = openingFrame.current;
      openingFrame.current = undefined;
    }
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
  const paintWithFonts = useCallback((displayList: DisplayList, paint: (list: DisplayList) => void): boolean => {
    const current = state.current;
    if (current && current.owner == null) openingFrame.current = displayList;
    if (current?.painter) return current.painter.paint(displayList, paint);
    paint(displayList);
    if (!current) return true;
    current.last = displayList;
    return false;
  }, []);
  return useMemo(() => Object.assign(paintWithFonts, {
    preload: (displayList: DisplayList, idle: () => boolean) => { state.current?.painter?.preload(displayList, idle); },
    withScriptFonts: (displayList: DisplayList) => state.current?.painter?.withScriptFonts?.(displayList) ?? displayList,
  }), [paintWithFonts]);
}

/** The faces a PNG export should paint `fonts` with, as the editor paints them; none when default fonts are off. */
export async function editorExportFaces(
  fonts: readonly XlsxPngFont[], decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined, defaultFonts: boolean | undefined,
  target?: XlsxPngTarget,
): Promise<XlsxFontFace[] | undefined> {
  if (defaultFonts === false || fonts.length === 0) return undefined;
  const { exportFaces } = await defaultFontPainterModule.load();
  const reported = new Set<Error>();
  const faces = await exportFaces(fonts, decoder, reported);
  if (faces.length) return faces;
  if (target && reported.size) {
    Object.defineProperty(target, Symbol.for('@betteroffice/xlsx/reportedFontCdnErrors'), { value: reported });
  }
  return undefined;
}
