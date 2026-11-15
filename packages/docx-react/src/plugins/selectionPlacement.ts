import { useRef } from 'react';
import {
  placeFloating,
  type FloatingPlacement,
  type FloatingPlacementOptions,
  type FloatingRect,
  type FloatingSize,
} from '../../../../shared/plugin-host/floatingPlacement';
import { pluginSelectionKeys } from './selectionIdentity';
import type { DocxPluginContext, DocxPluginGeometry } from './types';

/** @experimental The plugin API may change in minor releases. */
export interface DocxSelectionPlacementRequest extends FloatingSize, FloatingPlacementOptions {
  /** Geometry to read, such as an overlay's `geometry`; `context.geometry` by default. */
  geometry?: DocxPluginGeometry | null;
}

/** @experimental The plugin API may change in minor releases. */
export interface DocxSelectionPlacement extends FloatingPlacement {
  /** The selection's box on its anchor page, in overlay-layer pixels. */
  anchor: FloatingRect;
  /** The box is the last one placed: the selection is unchanged but the current layout has not placed it yet. */
  kept: boolean;
}

/** @experimental The plugin API may change in minor releases. */
export type DocxSelectionPlacer = (
  context: Pick<DocxPluginContext<unknown>, 'snapshot' | 'geometry'>,
  request: DocxSelectionPlacementRequest
) => DocxSelectionPlacement | null;

/**
 * Places a floating box by the current selection. While a new layout has not placed an unchanged
 * selection yet, the last box is kept; null without a selection box.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function createDocxSelectionPlacement(): DocxSelectionPlacer {
  let last: { key: string; anchor: FloatingRect; page: FloatingRect; viewport: FloatingRect } | null = null;
  return (context, { geometry: requested, width, height, ...options }) => {
    const geometry = requested === undefined ? context.geometry : requested;
    const selection = pluginSelectionKeys.get(context.snapshot);
    if (selection === undefined) return null;
    const key = `${context.snapshot.generation}\n${selection}`;
    const result = geometry?.getAnchorGeometry({ kind: 'selection' });
    let kept = false;
    if (geometry && result?.ok) {
      const onPage = result.rects.filter((rect) => rect.pageIndex === result.anchor.pageIndex);
      const rects = onPage.length ? onPage : [result.anchor];
      const x = Math.min(...rects.map((rect) => rect.x));
      const y = Math.min(...rects.map((rect) => rect.y));
      const anchor = {
        x, y,
        width: Math.max(...rects.map((rect) => rect.x + rect.width)) - x,
        height: Math.max(...rects.map((rect) => rect.y + rect.height)) - y,
      };
      last = { key, anchor, page: result.pageRect, viewport: geometry.viewport };
    } else if (last?.key === key) {
      kept = true;
      if (geometry) last.viewport = geometry.viewport;
    } else {
      last = null;
      return null;
    }
    const { anchor, page, viewport } = last;
    return { ...placeFloating({ ...options, anchor, page, viewport, size: { width, height } }), anchor, kept };
  };
}

/**
 * {@link createDocxSelectionPlacement} for an overlay: it places again on every render, so on
 * selection, viewport and layout changes.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function useDocxSelectionPlacement(
  context: Pick<DocxPluginContext<unknown>, 'snapshot' | 'geometry'>,
  request: DocxSelectionPlacementRequest
): DocxSelectionPlacement | null {
  const placer = useRef<DocxSelectionPlacer | null>(null);
  placer.current ??= createDocxSelectionPlacement();
  return placer.current(context, request);
}
