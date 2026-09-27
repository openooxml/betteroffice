import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import { sourceVersionOf } from '../components/DocxEditor/internals/layoutProvenance';
import type { DocxPluginGeometry, DocxPluginLayout, DocxPluginRect } from './types';

const layoutIds = new WeakMap<DisplayListQueries, string>();
let nextLayoutId = 0;

function layoutIdOf(queries: DisplayListQueries): string {
  let id = layoutIds.get(queries);
  if (!id) {
    nextLayoutId += 1;
    id = `layout-${nextLayoutId}`;
    layoutIds.set(queries, id);
  }
  return id;
}

/** The layout `queries` render, or null unless its pixels show `version`. */
export function pluginLayout(
  queries: DisplayListQueries | null,
  version: string | null,
  zoom: number
): DocxPluginLayout | null {
  if (!queries || version === null || sourceVersionOf(queries) !== version) return null;
  return { id: layoutIdOf(queries), version, zoom, pageCount: queries.pageCount() };
}

/**
 * Converts a rectangle in `RenderedDomContext` units (pages-container pixels divided by zoom)
 * into pixels of the unscaled `layer`, measuring both origins, the layer's border and its
 * scroll offset now.
 */
export function toOverlayRect(
  pages: HTMLElement,
  layer: HTMLElement,
  zoom: number,
  rect: DocxPluginRect
): DocxPluginRect {
  const pagesRect = pages.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  const originX = pagesRect.left - layerRect.left - layer.clientLeft + layer.scrollLeft;
  const originY = pagesRect.top - layerRect.top - layer.clientTop + layer.scrollTop;
  return {
    x: originX + rect.x * zoom,
    y: originY + rect.y * zoom,
    width: rect.width * zoom,
    height: rect.height * zoom,
  };
}

export function createPluginGeometry(
  layout: DocxPluginLayout,
  dom: RenderedDomContext,
  layer: HTMLElement,
  current: () => boolean
): DocxPluginGeometry {
  return {
    layout,
    dom,
    toOverlayRect: (rect) =>
      current() ? toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect) : null,
  };
}
