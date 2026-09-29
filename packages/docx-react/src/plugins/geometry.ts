import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import type { PointPosition, RenderedDomContext } from '@betteroffice/docx/plugin-api';
import { createCanvasHostProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { sourceVersionOf } from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { DocxPointPosition } from '../components/DocxEditor/types';
import { anchorFailure, resolveAnchorTarget } from './anchorGeometry';
import { currentPreviewKey, proposalSnapshot, renderedPreviewKey } from './proposalPreview';
import type { DocxAnchorRect, DocxPluginGeometry, DocxPluginLayout, DocxPluginRect } from './types';

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

/** The layout whose pixels show this document and proposal preview. */
export function pluginLayout(
  queries: DisplayListQueries | null,
  version: string | null,
  zoom: number,
  preview: { key: string; previewVersion: number }
): DocxPluginLayout | null {
  if (
    !queries ||
    version === null ||
    sourceVersionOf(queries) !== version ||
    renderedPreviewKey(queries) !== preview.key
  )
    return null;
  return {
    id: layoutIdOf(queries),
    version,
    previewVersion: preview.previewVersion,
    zoom,
    pageCount: queries.pageCount(),
  };
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

export interface AnchorGeometryAccess {
  session: YrsSession;
  editor: Pick<PagedEditorRef, 'yrsLocToDisplayPosition' | 'hasPendingInput'>;
  /** Whether the pages show this layout's pixels. */
  presented: boolean;
}

const EDGE = 0.5;

function sameLine(a: DisplayListRect, b: DisplayListRect): boolean {
  return a.pageIndex === b.pageIndex && a.y < b.y + b.height && b.y < a.y + a.height;
}

/** The last rectangle in reading order: page, then line, then the right edge. */
function lastInReadingOrder(rects: readonly DisplayListRect[]): DisplayListRect | null {
  let last: DisplayListRect | null = null;
  for (const rect of rects) {
    if (
      !last ||
      rect.pageIndex > last.pageIndex ||
      (rect.pageIndex === last.pageIndex &&
        (sameLine(rect, last) ? rect.x + rect.width > last.x + last.width : rect.y > last.y))
    ) {
      last = rect;
    }
  }
  return last && { ...last, x: last.x + last.width, width: 0 };
}

/** Geometry of the current frame, resolving targets against the live editor. */
export function createPluginGeometry(
  layout: DocxPluginLayout,
  dom: RenderedDomContext,
  layer: HTMLElement,
  current: () => boolean,
  resolve: (hit: PointPosition | null) => DocxPointPosition | null,
  queries: DisplayListQueries,
  access: () => AnchorGeometryAccess | null
): DocxPluginGeometry {
  const projector = createCanvasHostProjector(dom.pagesContainer, queries, dom.zoom);
  const project = (rect: DisplayListRect): DocxAnchorRect | null => {
    const projected = projector.projectRect(rect);
    return projected
      ? {
          ...toOverlayRect(dom.pagesContainer, layer, dom.zoom, projected),
          pageIndex: rect.pageIndex,
        }
      : null;
  };
  const unavailable = () =>
    anchorFailure('layout-unavailable', 'No rendered layout shows this target yet');
  /** The trailing edge of the unit before `to`, on the side its line of `rects` runs toward. */
  const endEdge = (to: number, rects: readonly DisplayListRect[]): DisplayListRect | null => {
    const unit = queries
      .rangeRects(to - 1, to)
      .filter((rect) => rect.width > 0)
      .at(-1);
    if (!unit) return null;
    const line = rects.find(
      (rect) =>
        sameLine(rect, unit) &&
        rect.x <= unit.x + EDGE &&
        unit.x + unit.width <= rect.x + rect.width + EDGE
    );
    const rtl =
      !!line &&
      Math.abs(unit.x - line.x) <= EDGE &&
      line.x + line.width - (unit.x + unit.width) > EDGE;
    return { ...unit, x: rtl ? unit.x : unit.x + unit.width, width: 0 };
  };
  return {
    layout,
    dom,
    toOverlayRect: (rect) =>
      current() ? toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect) : null,
    getPositionAtPoint(clientX, clientY) {
      if (!current()) return null;
      const position = resolve(dom.getPositionAtPoint?.(clientX, clientY) ?? null);
      return position ? { ...position, layoutId: layout.id } : null;
    },
    getAnchorGeometry(target) {
      if (!current()) return unavailable();
      const live = access();
      if (!live || !live.presented || live.editor.hasPendingInput()) return unavailable();
      const { session, editor } = live;
      if (session.version() !== layout.version) {
        return anchorFailure('stale-version', 'The document changed after that version');
      }
      if (
        (proposalSnapshot(session)?.previewVersion ?? 0) !== layout.previewVersion ||
        currentPreviewKey(session) !== renderedPreviewKey(queries)
      )
        return unavailable();
      const resolved = resolveAnchorTarget(session, target, layout.version);
      if (!resolved.ok) return resolved;
      const ranges: { from: number; to: number; hidden?: boolean }[] = [];
      for (const range of resolved.ranges) {
        const from = editor.yrsLocToDisplayPosition(range.start);
        const to = editor.yrsLocToDisplayPosition(range.end);
        if (from === null || to === null) {
          return anchorFailure('unsupported', 'The target has no body display position');
        }
        ranges.push({ from, to, hidden: range.hidden });
      }
      ranges.sort((a, b) => a.from - b.from || a.to - b.to);
      const union: { from: number; to: number }[] = [];
      for (const range of ranges) {
        if (range.hidden) continue;
        const previous = union.at(-1);
        if (previous && range.from < previous.to) previous.to = Math.max(previous.to, range.to);
        else union.push({ from: range.from, to: range.to });
      }
      const rects: DocxAnchorRect[] = [];
      const drawn: DisplayListRect[] = [];
      let tail: { to: number; rects: DisplayListRect[] } | null = null;
      for (const { from, to } of union) {
        if (from >= to) continue;
        const visible = queries.rangeRects(from, to).filter((rect) => rect.width > 0);
        for (const rect of visible) {
          const projected = project(rect);
          if (!projected) return unavailable();
          rects.push(projected);
          drawn.push(rect);
        }
        if (visible.length > 0) tail = { to, rects: visible };
      }
      const end = tail ? (endEdge(tail.to, tail.rects) ?? lastInReadingOrder(drawn)) : null;
      let anchor = end ? project(end) : null;
      const lastRange = ranges.at(-1);
      if (!anchor && lastRange) {
        const caret = queries.caretRect(lastRange.from);
        if (caret) anchor = project({ ...caret, width: 0 });
      }
      if (!anchor) {
        const position = editor.yrsLocToDisplayPosition(resolved.paragraph);
        const paragraph = position === null ? null : queries.anchorRect(position);
        if (paragraph) anchor = project({ ...paragraph, width: 0 });
      }
      if (!anchor) return unavailable();
      const page = projector.getPageBounds(anchor.pageIndex);
      if (!page) return unavailable();
      return {
        ok: true,
        version: layout.version,
        previewVersion: layout.previewVersion,
        layoutId: layout.id,
        rects,
        anchor,
        pageRect: toOverlayRect(dom.pagesContainer, layer, dom.zoom, page),
      };
    },
  };
}
