import {
  effectiveZoom,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { PointPosition, RenderedDomContext } from '@betteroffice/docx/plugin-api';
import { createCanvasHostProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { sourceVersionOf } from '../components/DocxEditor/internals/layoutProvenance';
import { displayWindowOf } from '../components/DocxEditor/internals/displayWindow';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { DocxPointPosition } from '../components/DocxEditor/types';
import {
  anchorFailure,
  hiddenRanges,
  resolveAnchorTarget,
  type RawAnchorRange,
} from './anchorGeometry';
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
 * into the `layer`'s own CSS pixels, measuring both origins, the layer's border and its
 * scroll offset now. Client offsets carry any ancestor CSS `zoom`; the result does not.
 */
export function toOverlayRect(
  pages: HTMLElement,
  layer: HTMLElement,
  zoom: number,
  rect: DocxPluginRect
): DocxPluginRect {
  const pagesRect = pages.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  const layerZoom = effectiveZoom(layer);
  const scale = zoom * (effectiveZoom(pages) / layerZoom);
  const originX =
    (pagesRect.left - layerRect.left) / layerZoom - layer.clientLeft + layer.scrollLeft;
  const originY = (pagesRect.top - layerRect.top) / layerZoom - layer.clientTop + layer.scrollTop;
  return {
    x: originX + rect.x * scale,
    y: originY + rect.y * scale,
    width: rect.width * scale,
    height: rect.height * scale,
  };
}

export interface AnchorGeometryAccess {
  session: YrsSession;
  editor: Pick<PagedEditorRef, 'yrsLocToDisplayPosition' | 'hasPendingInput'>;
  /** Whether the pages show this layout's pixels. */
  presented: boolean;
}

interface Interval {
  from: number;
  to: number;
}

/** `range` grown over every hidden interval it touches, directly or through another. */
function widen(range: Interval, hidden: readonly Interval[]): Interval {
  const span = { ...range };
  for (let grown = true; grown; ) {
    grown = false;
    for (const hole of hidden) {
      if (hole.from < span.from && hole.to >= span.from) {
        span.from = hole.from;
        grown = true;
      }
      if (hole.to > span.to && hole.from <= span.to) {
        span.to = hole.to;
        grown = true;
      }
    }
  }
  return span;
}

/** `ranges` without the parts `holes` cover, in document order. */
function subtract(ranges: readonly Interval[], holes: readonly Interval[]): Interval[] {
  let pieces = [...ranges];
  for (const hole of holes) {
    pieces = pieces.flatMap((piece) =>
      hole.to <= piece.from || hole.from >= piece.to
        ? [piece]
        : [
            { from: piece.from, to: hole.from },
            { from: hole.to, to: piece.to },
          ].filter((part) => part.from < part.to)
    );
  }
  return pieces.sort((a, b) => a.from - b.from || a.to - b.to);
}

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

/** Geometry of the current frame; visible unbuilt pages wait for exact content. */
export function createPluginGeometry(
  layout: DocxPluginLayout,
  dom: RenderedDomContext,
  layer: HTMLElement,
  current: () => boolean,
  resolve: (hit: PointPosition | null) => DocxPointPosition | null,
  queries: DisplayListQueries,
  access: () => AnchorGeometryAccess | null
): DocxPluginGeometry {
  const shown = () => dom.zoom === layout.zoom && current();
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
  /** Just past the last unit of `[from, to)` that draws, found by bisecting its visible suffix. */
  const lastUnitEnd = (from: number, to: number): number | null => {
    const draws = (start: number) => queries.rangeRects(start, to).some((rect) => rect.width > 0);
    if (!draws(from)) return null;
    let low = from;
    let high = to - 1;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (draws(middle)) low = middle;
      else high = middle - 1;
    }
    return low + 1;
  };
  const unitAt = (from: number): DisplayListRect | null =>
    queries
      .rangeRects(from, from + 1)
      .filter((rect) => rect.width > 0)
      .at(-1) ?? null;
  /** The text caret stop nearest `x` on `unit`'s line; null over an image or shape atom. */
  const stopAt = (unit: DisplayListRect, x: number): number | null => {
    const hit = queries.hitTestRegions(unit.pageIndex, x, unit.y + unit.height / 2);
    return hit?.region === 'body' && hit.target !== 'image' ? hit.pos : null;
  };
  /**
   * The collapsed caret at `pos`, on the edge of a neighbouring unit whose caret stop is nearest
   * `pos`: the unit before it first when `before`, else the unit after it first. An atom, which
   * has no text stops, keeps the caret on its near side.
   */
  const caretAt = (pos: number, before = false): DisplayListRect | null => {
    for (const from of before ? [pos - 1, pos] : [pos, pos - 1]) {
      const unit = unitAt(from);
      if (!unit) continue;
      const inset = Math.min(1, unit.width / 4);
      const left = stopAt(unit, unit.x + inset);
      const right = stopAt(unit, unit.x + unit.width - inset);
      const atLeft =
        left === null || right === null
          ? from === pos
          : Math.abs(left - pos) <= Math.abs(right - pos);
      return { ...unit, x: atLeft ? unit.x : unit.x + unit.width, width: 0 };
    }
    return null;
  };
  const endOf = ({ from, to }: Interval): DisplayListRect | null => {
    const end = lastUnitEnd(from, to);
    return end === null ? null : caretAt(end, true);
  };
  return {
    layout,
    dom,
    toOverlayRect: (rect) =>
      shown() ? toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect) : null,
    getPositionAtPoint(clientX, clientY) {
      if (!shown()) return null;
      const position = resolve(dom.getPositionAtPoint?.(clientX, clientY) ?? null);
      return position ? { ...position, layoutId: layout.id } : null;
    },
    getAnchorGeometry(target) {
      if (!shown()) return unavailable();
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
      const window = displayWindowOf(queries)?.read();
      const pendingPage = (pageIndex: number): boolean =>
        !!window &&
        pageIndex >= window[0] &&
        pageIndex < window[1] &&
        queries.displayList?.pages[pageIndex]?.unbuilt === true;
      const display = (range: RawAnchorRange): Interval | null => {
        const from = editor.yrsLocToDisplayPosition(range.start);
        const to = editor.yrsLocToDisplayPosition(range.end);
        return from === null || to === null ? null : { from, to };
      };
      const ranges: Interval[] = [];
      for (const range of resolved.ranges) {
        const mapped = display(range);
        if (!mapped) return anchorFailure('unsupported', 'The target has no body display position');
        ranges.push(mapped);
      }
      ranges.sort((a, b) => a.from - b.from || a.to - b.to);
      if (
        window &&
        queries.displayList?.pages.slice(window[0], window[1]).some(
          ({ unbuilt, positionSpan }) =>
            unbuilt &&
            positionSpan &&
            ranges.some(({ from, to }) => from <= positionSpan[1] && to >= positionSpan[0])
        )
      )
        return unavailable();
      const hidden = hiddenRanges(session, layout.version)
        .map(display)
        .filter((range): range is Interval => range !== null);
      const union: Interval[] = [];
      for (const range of subtract(ranges, hidden)) {
        const previous = union.at(-1);
        if (previous && range.from < previous.to) previous.to = Math.max(previous.to, range.to);
        else union.push({ ...range });
      }
      const rects: DocxAnchorRect[] = [];
      const drawn: DisplayListRect[] = [];
      let tail: Interval | null = null;
      for (const { from, to } of union) {
        if (from >= to) continue;
        const visible = queries.rangeRects(from, to).filter((rect) => rect.width > 0);
        for (const rect of visible) {
          if (pendingPage(rect.pageIndex)) return unavailable();
          const projected = project(rect);
          if (!projected) return unavailable();
          rects.push(projected);
          drawn.push(rect);
        }
        if (visible.length > 0) tail = { from, to };
      }
      const last = ranges.at(-1);
      const gap = last && widen(last, hidden);
      const paragraph = editor.yrsLocToDisplayPosition(resolved.paragraph);
      const end =
        tail !== null
          ? (endOf(tail) ?? lastInReadingOrder(drawn))
          : ((gap && (caretAt(gap.from, true) ?? caretAt(gap.to))) ??
            (paragraph === null ? null : caretAt(paragraph)));
      const fallback = end || paragraph === null ? null : queries.anchorRect(paragraph);
      if ((end && pendingPage(end.pageIndex)) || (fallback && pendingPage(fallback.pageIndex))) {
        return unavailable();
      }
      const anchor = end ? project(end) : fallback ? project({ ...fallback, width: 0 }) : null;
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
