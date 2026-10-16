import {
  effectiveZoom,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { PointPosition, RenderedDomContext } from '@betteroffice/docx/plugin-api';
import { createCanvasHostProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  anchorDisplayRoot,
  headerFooterDisplayRIds,
  paragraphInlineStart,
  proposalSetIdentity,
  resolvedTargetGeometry,
  type AnchorGeometryTarget,
  type DocxReservationsPayload,
  type DocxResolvedReservations,
  type ProposalGeometryMirror,
  type ProposalGeometryTarget,
  type YrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  reservationsKeyOf,
  reservationsVersionOfKey,
  sourceVersionOf,
} from '../components/DocxEditor/internals/layoutProvenance';
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
import type {
  DocxAnchorGeometryResult,
  DocxAnchorRect,
  DocxGeometryTarget,
  DocxPluginGeometry,
  DocxPluginLayout,
  DocxPluginPageRect,
  DocxPluginRect,
} from './types';

import { flushEditorInput } from '../components/DocxEditor/editorBatches';
import { isWorkerViewer } from '../components/DocxEditor/internals/workerViewer';
import { hasEditorWorkerProposalRounds } from '../components/DocxEditor/internals/workerProposalAuthority';
import { workerOpenReplicaReady } from '../components/DocxEditor/internals/workerOpenReplica';

export async function readPluginPositionAtPoint(
  editorRef: React.RefObject<PagedEditorRef | null>,
  clientX: number,
  clientY: number,
  experimentalWorkerOpen = false
): Promise<DocxPointPosition | null> {
  if (isWorkerViewer(editorRef.current)) return editorRef.current?.readPositionAtPoint(clientX, clientY) ?? null;
  const flushed = await flushEditorInput(editorRef, experimentalWorkerOpen);
  if (!flushed.ok && flushed.code !== 'editor-unavailable') throw flushed.error;
  return editorRef.current?.getPositionAtPoint(clientX, clientY) ?? null;
}

const layoutIds = new WeakMap<DisplayListQueries, string>();
let nextLayoutId = 0;

export function layoutIdOf(queries: DisplayListQueries): string {
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
    reservationsVersion: reservationsVersionOfKey(reservationsKeyOf(queries)),
    zoom,
    pageCount: queries.pageCount(),
  };
}

/** @internal Answers a plugin's `reservation` targets on one geometry. */
export interface ReservationReader {
  get(owner: string, id: string): DocxAnchorGeometryResult;
  /** Answers `targets` in order once their reads land, `reservation` ones as `owner`'s. */
  read?(owner: string, targets: readonly DocxGeometryTarget[]): Promise<DocxAnchorGeometryResult[]>;
}

const reservationReaders = new WeakMap<DocxPluginGeometry, ReservationReader>();

/** @internal Lets plugin-scoped views of `geometry` answer their `reservation` targets with `reader`. */
export function registerReservationReader(geometry: DocxPluginGeometry, reader: ReservationReader): void {
  reservationReaders.set(geometry, reader);
}
const scopedGeometries = new WeakMap<DocxPluginGeometry, Map<string, DocxPluginGeometry>>();

/** `geometry` as plugin `pluginId` reads it: its `reservation` targets name its own reservations. */
export function scopePluginGeometry(
  geometry: DocxPluginGeometry,
  pluginId: string
): DocxPluginGeometry {
  let byPlugin = scopedGeometries.get(geometry);
  if (!byPlugin) {
    byPlugin = new Map();
    scopedGeometries.set(geometry, byPlugin);
  }
  let scoped = byPlugin.get(pluginId);
  if (!scoped) {
    const reservation = (target: DocxGeometryTarget) => {
      if (target.kind !== 'reservation') return geometry.getAnchorGeometry(target);
      const reader = reservationReaders.get(geometry);
      return reader
        ? reader.get(pluginId, target.id)
        : anchorFailure('unsupported', 'This layout reports no reservations');
    };
    scoped = {
      get layout() {
        return geometry.layout;
      },
      get dom() {
        return geometry.dom;
      },
      toOverlayRect: (rect) => geometry.toOverlayRect(rect),
      getPositionAtPoint: (clientX, clientY) => geometry.getPositionAtPoint(clientX, clientY),
      readPositionAtPoint: (clientX, clientY) => geometry.readPositionAtPoint(clientX, clientY),
      readAnchorGeometries: async (targets) => {
        if (!targets.some((target) => target.kind === 'reservation')) return geometry.readAnchorGeometries(targets);
        const reader = reservationReaders.get(geometry);
        return reader?.read ? reader.read(pluginId, targets) : targets.map(reservation);
      },
      getPageRect: (pageIndex) => geometry.getPageRect(pageIndex),
      isPageBuilt: (pageIndex) => geometry.isPageBuilt(pageIndex),
      getAnchorGeometry: reservation,
    };
    byPlugin.set(pluginId, scoped);
  }
  return scoped;
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
  return overlayProjector(pages, layer, zoom)(rect);
}

/** @internal {@link toOverlayRect} with the origins and zoom measured once, now. */
export function overlayProjector(
  pages: HTMLElement,
  layer: HTMLElement,
  zoom: number
): (rect: DocxPluginRect) => DocxPluginRect {
  const pagesRect = pages.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  const layerZoom = effectiveZoom(layer);
  const scale = zoom * (effectiveZoom(pages) / layerZoom);
  const originX =
    (pagesRect.left - layerRect.left) / layerZoom - layer.clientLeft + layer.scrollLeft;
  const originY = (pagesRect.top - layerRect.top) / layerZoom - layer.clientTop + layer.scrollTop;
  return (rect) => ({
    x: originX + rect.x * scale,
    y: originY + rect.y * scale,
    width: rect.width * scale,
    height: rect.height * scale,
  });
}

export interface AnchorGeometryAccess {
  session: YrsSession;
  /** @internal */
  proposalGeometry?: ProposalGeometryMirror;
  /** @internal Display geometry of other targets at `proposalGeometry`; undefined until known. */
  anchorTarget?(
    target: Exclude<AnchorGeometryTarget, { kind: 'proposal' }>
  ): ProposalGeometryTarget | undefined;
  editor: Pick<PagedEditorRef, 'yrsLocToDisplayPosition' | 'hasPendingInput'>;
  /** Whether the pages show this layout's pixels. */
  presented: boolean;
  /** The reservations last resolved for layout and their key; '' keys none. */
  reservations?(): {
    key: string;
    resolved: DocxResolvedReservations | null;
    payload?: DocxReservationsPayload | null;
  };
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

/** Overlapping `ranges`, in document order, merged. */
function unionOf(ranges: readonly Interval[]): Interval[] {
  const union: Interval[] = [];
  for (const range of ranges) {
    const previous = union.at(-1);
    if (previous && range.from < previous.to) previous.to = Math.max(previous.to, range.to);
    else union.push({ ...range });
  }
  return union;
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

/** Whether `[from, to)` reaches the inclusive `span`; an empty range reaches the span it sits in. */
const reaches = ({ from, to }: Interval, [low, high]: readonly [number, number]): boolean =>
  from <= high && (from === to ? to >= low : to > low);

/** @internal Unbuilt pages whose display positions overlap `ranges`. */
export function unbuiltPagesOf(queries: DisplayListQueries, ranges: readonly Interval[]): number[] {
  return (queries.displayList?.pages ?? []).flatMap(({ unbuilt, positionSpan }, pageIndex) =>
    unbuilt && positionSpan && ranges.some((range) => reaches(range, positionSpan))
      ? [pageIndex]
      : []
  );
}

/** @internal Whether `pageIndex` is a page of `queries` whose content is laid out. */
export function pageBuilt(queries: DisplayListQueries, pageIndex: number): boolean {
  const pages = queries.displayList?.pages;
  if (pages) return pages[pageIndex] !== undefined && pages[pageIndex]!.unbuilt !== true;
  return Number.isInteger(pageIndex) && pageIndex >= 0 && pageIndex < queries.pageCount();
}

/** @internal The page box of `pageIndex`, or null while it has none. */
export function pageRectOf(
  queries: DisplayListQueries,
  pageIndex: number,
  bounds: DocxPluginRect | null,
  overlay: (rect: DocxPluginRect) => DocxPluginRect
): DocxPluginPageRect | null {
  return bounds
    ? { ...overlay(bounds), pageIndex, built: pageBuilt(queries, pageIndex) }
    : null;
}

/** @internal What {@link placeRepeated} reads of one presented layout. */
export interface RepeatedPlacement {
  queries: DisplayListQueries;
  /** Whether the layout is whole and final. */
  complete(): boolean;
  /** Whether a page shown in the display window is not built yet. */
  pending(pageIndex: number): boolean;
  /** Projects `rects` and `anchor` onto the pages as they stand now. */
  respond(
    rects: readonly DisplayListRect[],
    anchor: DisplayListRect | null,
    unbuiltPages?: readonly number[]
  ): DocxAnchorGeometryResult;
  unavailable(): DocxAnchorGeometryResult;
}

/** Just past the last unit of `[from, to)` that draws, over spans doubling back from `to`. */
function lastUnitEnd(queries: DisplayListQueries, from: number, to: number): number | null {
  const draws = (start: number) => queries.rangeRects(start, to).some((rect) => rect.width > 0);
  let high = to;
  let low = Math.max(to - 1, from);
  for (let span = 2; !draws(low); span *= 2) {
    if (low <= from) return null;
    high = low;
    low = Math.max(to - span, from);
  }
  while (high - low > 1) {
    const middle = Math.floor((low + high) / 2);
    if (draws(middle)) low = middle;
    else high = middle;
  }
  return low + 1;
}

/** Where the first unit of `[from, to)` that draws starts, over spans doubling from `from`. */
function firstUnitStart(queries: DisplayListQueries, from: number, to: number): number | null {
  const draws = (end: number) => queries.rangeRects(from, end).some((rect) => rect.width > 0);
  let low = from;
  let high = Math.min(from + 1, to);
  for (let span = 2; !draws(high); span *= 2) {
    if (high >= to) return null;
    low = high;
    high = Math.min(from + span, to);
  }
  while (high - low > 1) {
    const middle = Math.floor((low + high) / 2);
    if (draws(middle)) high = middle;
    else low = middle;
  }
  return high - 1;
}

/** The last unit drawn at `from`, which has width. */
const unitAt = (queries: DisplayListQueries, from: number): DisplayListRect | null =>
  queries
    .rangeRects(from, from + 1)
    .filter((rect) => rect.width > 0)
    .at(-1) ?? null;

/** The text caret stop nearest `x` on `unit`'s line; null over an image or shape atom. */
function stopAt(queries: DisplayListQueries, unit: DisplayListRect, x: number): number | null {
  const hit = queries.hitTestRegions(unit.pageIndex, x, unit.y + unit.height / 2);
  return hit?.region === 'body' && hit.target !== 'image' ? hit.pos : null;
}

/**
 * The collapsed caret at `pos`, on the edge of a neighbouring unit whose caret stop is nearest
 * `pos`: the unit before it first when `before`, else the unit after it first. An atom, which
 * has no text stops, keeps the caret on its near side.
 */
function caretAt(queries: DisplayListQueries, pos: number, before = false): DisplayListRect | null {
  for (const from of before ? [pos - 1, pos] : [pos, pos - 1]) {
    const unit = unitAt(queries, from);
    if (!unit) continue;
    const inset = Math.min(1, unit.width / 4);
    const left = stopAt(queries, unit, unit.x + inset);
    const right = stopAt(queries, unit, unit.x + unit.width - inset);
    const atLeft =
      left === null || right === null
        ? from === pos
        : Math.abs(left - pos) <= Math.abs(right - pos);
    return { ...unit, x: atLeft ? unit.x : unit.x + unit.width, width: 0 };
  }
  return null;
}

/**
 * @internal The caret `ranges` leave when none of them draws: before the text that does not draw
 * around the last one in `paragraph`, else after it, else at `paragraph`, else at the nearest unit
 * that draws on built pages, the next one first.
 */
export function hiddenBoundary(
  queries: DisplayListQueries,
  ranges: readonly Interval[],
  hidden: readonly Interval[],
  paragraph: number | null
): DisplayListRect | null {
  const at = (): DisplayListRect | null => {
    const rect =
      paragraph === null ? null : (caretAt(queries, paragraph) ?? queries.anchorRect(paragraph));
    return rect && { ...rect, width: 0 };
  };
  const last = ranges.at(-1);
  if (!last) return at();
  const gap = widen(last, hidden);
  const spans = (queries.displayList?.pages ?? []).flatMap(({ unbuilt, positionSpan }) =>
    unbuilt && positionSpan ? [positionSpan] : []
  );
  const start = Math.min(paragraph ?? gap.from, gap.from);
  const built = !spans.some(([low, high]) => start <= high && gap.to >= low);
  const before = (): DisplayListRect | null => {
    const end = built && paragraph !== null ? lastUnitEnd(queries, paragraph, gap.from) : null;
    return end === null ? null : caretAt(queries, end, true);
  };
  const nearest = (): DisplayListRect | null => {
    if (!built) return null;
    const ceiling = Math.min(2 ** 32, ...spans.flatMap(([low]) => (low >= gap.to ? [low] : [])));
    const next = firstUnitStart(queries, gap.to, ceiling);
    if (next !== null) return caretAt(queries, next);
    const floor = Math.max(0, ...spans.flatMap(([, high]) => (high < gap.from ? [high + 1] : [])));
    const end = lastUnitEnd(queries, floor, gap.from);
    return end === null ? null : caretAt(queries, end, true);
  };
  return (
    caretAt(queries, gap.from, true) ?? before() ?? caretAt(queries, gap.to) ?? at() ?? nearest()
  );
}

/** @internal `owner`'s reservation boxes `id` on the built pages; unbuilt pages wait. */
export function placeReservation(
  { queries, complete, pending, respond, unavailable }: RepeatedPlacement,
  owner: string,
  id: string
): DocxAnchorGeometryResult {
  const pages = queries.displayList?.pages ?? [];
  const rects: DisplayListRect[] = [];
  let unbuilt = false;
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    const page = pages[pageIndex]!;
    const found = page.reservations?.filter((box) => box.owner === owner && box.id === id) ?? [];
    if (found.length && pending(pageIndex)) return unavailable();
    if (page.unbuilt) {
      unbuilt = true;
      continue;
    }
    for (const box of found) {
      rects.push({ pageIndex, x: box.x, y: box.y, width: box.width, height: box.height });
    }
  }
  if (rects.length) return respond(rects, rects[0]!);
  return unbuilt || !complete()
    ? unavailable()
    : anchorFailure('missing-target', 'No page shows the reservation');
}

/**
 * @internal Places a target in a header or footer story, which repeats on every page whose band
 * shows it under one of `routed` relationship ids. Unbuilt pages outside the display window are
 * left out; the anchor sits on the first built page that paints it, else on the first unbuilt one.
 */
export function placeRepeated(
  { queries, complete, pending, respond, unavailable }: RepeatedPlacement,
  routed: readonly string[],
  ranges: Interval[],
  hidden: readonly Interval[],
  paragraph: number | null
): DocxAnchorGeometryResult {
  const rIds = new Set(routed);
  const pages = queries.displayList?.pages ?? [];
  const bands = pages.flatMap(({ header, footer, hfParts, unbuilt }, pageIndex) => {
    for (const [region, rId] of [
      ['header', header?.rId ?? hfParts?.header],
      ['footer', footer?.rId ?? hfParts?.footer],
    ] as const) {
      if (rId !== undefined && rIds.has(rId)) return [{ region, rId, pageIndex, unbuilt: !!unbuilt }];
    }
    return [];
  });
  if (!bands.length) {
    return complete()
      ? anchorFailure('missing-target', 'No page paints this header or footer')
      : anchorFailure('layout-unavailable', 'No laid-out page paints this header or footer');
  }
  if (bands.some(({ pageIndex }) => pending(pageIndex))) return unavailable();
  const built = new Map(
    bands.filter(({ unbuilt }) => !unbuilt).map((band) => [`${band.region}:${band.rId}`, band])
  );
  ranges.sort((a, b) => a.from - b.from || a.to - b.to);
  const shown = unionOf(subtract(ranges, hidden)).filter(({ from, to }) => from < to);
  const drawn = [...built.values()]
    .flatMap(({ region, rId }) =>
      shown.flatMap(({ from, to }) =>
        queries.hfRangeRects(region, rId, from, to).filter((rect) => rect.width > 0)
      )
    )
    .sort((a, b) => a.pageIndex - b.pageIndex);
  const firstOf = (candidates: readonly DisplayListRect[]) =>
    candidates.reduce<DisplayListRect | null>(
      (first, rect) => (!first || rect.pageIndex < first.pageIndex ? rect : first),
      null
    );
  const first = firstOf(drawn);
  const last = ranges.at(-1);
  const gap = last && widen(last, hidden);
  const end = first
    ? lastInReadingOrder(drawn.filter(({ pageIndex }) => pageIndex === first.pageIndex))
    : firstOf(
        [...built.values()].flatMap(({ region, rId }) => [
          ...(gap ? queries.hfCaretRects(region, rId, gap.from) : []),
          ...(gap ? queries.hfCaretRects(region, rId, gap.to) : []),
          ...(paragraph === null ? [] : queries.hfAnchorRects(region, rId, paragraph)),
        ])
      );
  const unbuiltPages = [...new Set(bands.filter((band) => band.unbuilt).map((band) => band.pageIndex))];
  if (end) return respond(drawn, { ...end, width: 0 }, unbuiltPages);
  const unbuilt = bands.find((band) => band.unbuilt);
  if (!unbuilt) return respond(drawn, null);
  const content = pages[unbuilt.pageIndex]?.contentBounds;
  return respond(drawn, {
    pageIndex: unbuilt.pageIndex,
    x: content?.x ?? 0,
    y: (content?.y ?? 0) + (unbuilt.region === 'footer' ? (content?.height ?? 0) : 0),
    width: 0,
    height: 0,
  }, unbuiltPages);
}

/**
 * Geometry of the current frame; visible unbuilt pages wait for exact content. While `held`,
 * overlays still draw this layout until geometry for the next one exists, so `toOverlayRect`
 * answers; hit tests and anchors wait for the new layout.
 */
export function createPluginGeometry(
  layout: DocxPluginLayout,
  dom: RenderedDomContext,
  layer: HTMLElement,
  current: () => boolean,
  resolve: (hit: PointPosition | null) => DocxPointPosition | null,
  queries: DisplayListQueries,
  access: (reservation?: boolean) => AnchorGeometryAccess | null,
  held: () => boolean = () => false,
  readPoint?: (clientX: number, clientY: number) => Promise<DocxPointPosition | null>,
  complete: () => boolean = () => false
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
  const endOf = ({ from, to }: Interval): DisplayListRect | null => {
    const end = lastUnitEnd(queries, from, to);
    return end === null ? null : caretAt(queries, end, true);
  };
  const pendingPage = (pageIndex: number): boolean => {
    const window = displayWindowOf(queries)?.read();
    return (
      !!window &&
      pageIndex >= window[0] &&
      pageIndex < window[1] &&
      queries.displayList?.pages[pageIndex]?.unbuilt === true
    );
  };
  /** Projects `rects` and `anchor` onto the pages as they stand now. */
  const respond = (
    rects: readonly DisplayListRect[],
    anchor: DisplayListRect | null,
    unbuiltPages: readonly number[] = []
  ): DocxAnchorGeometryResult => {
    const projected = rects.map(project);
    const anchorRect = anchor && project(anchor);
    const page = anchorRect && projector.getPageBounds(anchorRect.pageIndex);
    if (!anchorRect || !page || projected.some((rect) => rect === null)) return unavailable();
    return {
      ok: true,
      version: layout.version,
      previewVersion: layout.previewVersion,
      layoutId: layout.id,
      rects: projected as DocxAnchorRect[],
      anchor: anchorRect,
      pageRect: toOverlayRect(dom.pagesContainer, layer, dom.zoom, page),
      ...(unbuiltPages.length ? { unbuiltPages } : {}),
    };
  };

  /** Places `target`, or the raw ranges a caller resolved itself. */
  const anchorGeometry = (
    target: AnchorGeometryTarget | null,
    raw?: { ok: true; ranges: RawAnchorRange[]; paragraph: YrsLoc },
    reservationTarget = false,
    hiddenTarget = false
  ): DocxAnchorGeometryResult => {
    if (!shown()) return unavailable();
    const live = access(reservationTarget);
    if (!live || !live.presented || live.editor.hasPendingInput()) return unavailable();
    const { session, editor } = live;
    if (session.version() !== layout.version) {
      return anchorFailure('stale-version', 'The document changed after that version');
    }
    const snapshot = proposalSnapshot(session);
    if (
      (snapshot?.previewVersion ?? 0) !== layout.previewVersion ||
      currentPreviewKey(session) !== renderedPreviewKey(queries)
    )
      return unavailable();
    const mirror = live.proposalGeometry;
    if (
      mirror &&
      (mirror.version !== layout.version ||
        mirror.previewVersion !== layout.previewVersion ||
        !snapshot ||
        mirror.proposals !== proposalSetIdentity(snapshot))
    )
      return unavailable();
    if (mirror && !target) return unavailable();
    const mirrored = !mirror || !target
      ? null
      : target.kind === 'proposal'
        ? Object.hasOwn(mirror.targets, target.id)
          ? mirror.targets[target.id]
          : anchorFailure('unknown-proposal', 'The proposal is not registered in this document')
        : live.anchorTarget?.(target);
    if (mirror && !mirrored) return unavailable();
    if (mirrored && !mirrored.ok) return mirrored;
    const resolved = raw ?? (mirror || !target ? null : resolveAnchorTarget(session, target, layout.version));
    if (resolved && !resolved.ok) return resolved;
    const root = resolved ? anchorDisplayRoot(resolved.paragraph.story) : 'body';
    if (root === null) return anchorFailure('unsupported', 'The target has no display position');
    const repeated =
      resolved && root !== 'body'
        ? resolvedTargetGeometry(session, layout.version, () =>
            hiddenRanges(session, layout.version)
          )(resolved)
        : mirrored;
    if (repeated && !repeated.ok) return repeated;
    if (repeated?.root) {
      const { root: story, rIds, ranges, hidden = [], paragraph } = repeated;
      const routed = rIds ?? headerFooterDisplayRIds(session, story.slice('hf:'.length));
      return placeRepeated(
        { queries, complete, pending: pendingPage, respond, unavailable },
        routed,
        [...ranges],
        hidden,
        paragraph
      );
    }
    const window = displayWindowOf(queries)?.read();
    const display = (range: RawAnchorRange): Interval | null => {
      const from = editor.yrsLocToDisplayPosition(range.start);
      const to = editor.yrsLocToDisplayPosition(range.end);
      return from === null || to === null ? null : { from, to };
    };
    const ranges: Interval[] = mirrored?.ok ? [...mirrored.ranges] : [];
    for (const range of resolved?.ok ? resolved.ranges : []) {
      const mapped = display(range);
      if (!mapped) return anchorFailure('unsupported', 'The target has no display position');
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
    const hidden = [
      ...((mirror && (!hasEditorWorkerProposalRounds(session) || !workerOpenReplicaReady(session))
        ? mirror.hidden
        : undefined) ??
      hiddenRanges(session, layout.version)
        .map(display)
        .filter((range): range is Interval => range !== null)),
      ...(hiddenTarget ? ranges : []),
    ];
    const drawn: DisplayListRect[] = [];
    let tail: Interval | null = null;
    for (const { from, to } of unionOf(subtract(ranges, hidden))) {
      if (from >= to) continue;
      const visible = queries.rangeRects(from, to).filter((rect) => rect.width > 0);
      for (const rect of visible) {
        if (pendingPage(rect.pageIndex)) return unavailable();
        drawn.push(rect);
      }
      if (visible.length > 0) tail = { from, to };
    }
    const paragraph = mirrored?.ok
      ? mirrored.paragraph
      : resolved?.ok
        ? editor.yrsLocToDisplayPosition(resolved.paragraph)
        : null;
    const end =
      tail !== null
        ? (endOf(tail) ?? lastInReadingOrder(drawn))
        : hiddenBoundary(queries, ranges, hidden, paragraph);
    if (end && pendingPage(end.pageIndex)) return unavailable();
    return respond(drawn, end, unbuiltPagesOf(queries, ranges));
  };
  /** This layout's box for `owner`'s reservation `id`, or a hidden range's boundary. */
  const reservation = (owner: string, id: string): DocxAnchorGeometryResult => {
    if (!shown()) return unavailable();
    const live = access(true);
    if (!live || !live.presented || live.editor.hasPendingInput()) return unavailable();
    if (live.session.version() !== layout.version) {
      return anchorFailure('stale-version', 'The document changed after that version');
    }
    const current = live.reservations?.();
    if (!current || reservationsKeyOf(queries) !== current.key) return unavailable();
    const entry = current.resolved?.[JSON.stringify([owner, id])];
    const item = current.payload?.sets.find((set) => set.owner === owner)?.items.find((item) => item.id === id);
    if (!entry && !item) return anchorFailure('missing-target', 'The reservation did not resolve');
    if (item?.kind === 'hide') return anchorGeometry(item.anchor, undefined, true, true);
    if (entry?.kind === 'hide') {
      const lead = paragraphInlineStart(live.session, layout.version, entry.story, entry.paraId);
      const at = (offset: number): YrsLoc => ({ story: entry.story, paraId: entry.paraId, offset: lead + offset });
      const span = live.session.paragraphSpans(entry.story).find((span) => span.paraId === entry.paraId);
      if (!span || lead + entry.start > span.length || lead + entry.end > span.length + 1) {
        return anchorFailure('missing-target', 'The reservation did not resolve');
      }
      const length = span.length - lead;
      return anchorGeometry(null, {
        ok: true,
        ranges: [{ start: at(entry.start), end: at(Math.min(entry.end, length)) }],
        paragraph: at(0),
      }, false, true);
    }
    return placeReservation({ queries, complete, pending: pendingPage, respond, unavailable }, owner, id);
  };
  const drawable = () => shown() || (dom.zoom === layout.zoom && held());
  const created: DocxPluginGeometry = {
    layout,
    dom,
    toOverlayRect: (rect) =>
      drawable() ? toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect) : null,
    getPageRect: (pageIndex) =>
      drawable()
        ? pageRectOf(queries, pageIndex, projector.getPageBounds(pageIndex), (rect) =>
            toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect)
          )
        : null,
    isPageBuilt: (pageIndex) => drawable() && pageBuilt(queries, pageIndex),
    readAnchorGeometries: async (targets) => targets.map((target) => created.getAnchorGeometry(target)),
    getPositionAtPoint(clientX, clientY) {
      if (!shown()) return null;
      const position = resolve(dom.getPositionAtPoint?.(clientX, clientY) ?? null);
      return position ? { ...position, layoutId: layout.id } : null;
    },
    async readPositionAtPoint(clientX, clientY) {
      if (!shown()) return null;
      const position = readPoint
        ? await readPoint(clientX, clientY)
        : resolve(dom.getPositionAtPoint?.(clientX, clientY) ?? null);
      return shown() && position && position.version === layout.version
        ? { ...position, layoutId: layout.id } : null;
    },
    getAnchorGeometry: (target) =>
      target.kind === 'reservation'
        ? anchorFailure('unsupported', 'Reservations are read through a plugin context')
        : anchorGeometry(target),
  };
  registerReservationReader(created, { get: reservation });
  return created;
}
