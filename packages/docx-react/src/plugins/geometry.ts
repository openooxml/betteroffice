import {
  effectiveZoom,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { PointPosition, RenderedDomContext } from '@betteroffice/docx/plugin-api';
import { createCanvasHostProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  proposalSetIdentity,
  type AnchorDisplayTarget,
  type DocxParagraphAnchor,
  type ProposalGeometryMirror,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  isPresented,
  sourceVersionOf,
  workerFrameVersionOf,
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
  /** @internal */
  proposalGeometry?: ProposalGeometryMirror;
  editor: Pick<PagedEditorRef, 'yrsLocToDisplayPosition' | 'hasPendingInput'>;
  /** Whether the pages show this layout's pixels. */
  presented: boolean;
}

type WorkerTarget = Exclude<DocxGeometryTarget, { kind: 'proposal' }>;

/**
 * Resolves targets in the worker at the worker `version` under the rendered `previewKey`, in input
 * order; null once the worker moved past it.
 */
export type ReadAnchorTargets = (
  targets: readonly WorkerTarget[],
  version: string,
  previewVersion: number,
  previewKey: string
) => Promise<readonly AnchorDisplayTarget[] | null>;

interface AnchorPlacement {
  rects: readonly DisplayListRect[];
  anchor: DisplayListRect;
}

/**
 * Resolved worker replies of one worker version and rendered preview, placed again on every
 * layout; at most `limit` are kept. `pending` holds reads in flight.
 */
export interface AnchorReadCache {
  key: string;
  limit: number;
  replies: Map<string, Extract<AnchorDisplayTarget, { ok: true }>>;
  pending: Map<string, Promise<AnchorDisplayTarget | null | undefined>>;
}

export function createAnchorReadCache(limit = 4096): AnchorReadCache {
  return { key: '', limit, replies: new Map(), pending: new Map() };
}

function paragraphAnchor(anchor: DocxParagraphAnchor): DocxParagraphAnchor {
  switch (anchor.kind) {
    case 'session':
      return { kind: 'session', sessionId: anchor.sessionId, story: anchor.story, paraId: anchor.paraId };
    case 'source':
      return {
        kind: 'source',
        packageSha256: anchor.packageSha256,
        partUri: anchor.partUri,
        paragraphOrdinal: anchor.paragraphOrdinal,
      };
    case 'persisted': {
      const { story } = anchor;
      return {
        kind: 'persisted',
        story: 'itemId' in story
          ? { partUri: story.partUri, kind: story.kind, itemId: story.itemId }
          : { partUri: story.partUri, kind: story.kind },
        paraId: anchor.paraId,
      };
    }
  }
}

/** `target` with only the fields that resolve it. */
function normalTarget(target: WorkerTarget): WorkerTarget {
  switch (target.kind) {
    case 'revision':
      return { kind: 'revision', revisionId: target.revisionId };
    case 'paragraph':
      return { kind: 'paragraph', paragraph: paragraphAnchor(target.paragraph) };
    case 'search':
      return {
        kind: 'search',
        paragraph: paragraphAnchor(target.paragraph),
        text: target.text,
        ...(target.occurrence === undefined ? {} : { occurrence: target.occurrence }),
      };
    case 'range': {
      const { story, start, end, view } = target.range;
      return {
        kind: 'range',
        version: target.version,
        range: {
          story,
          start: { paraId: start.paraId, offset: start.offset },
          end: { paraId: end.paraId, offset: end.offset },
          view,
        },
      };
    }
  }
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
  access: () => AnchorGeometryAccess | null,
  held: () => boolean = () => false,
  readPoint?: (clientX: number, clientY: number) => Promise<DocxPointPosition | null>,
  workerReads?: { read: ReadAnchorTargets; cache: AnchorReadCache }
): DocxPluginGeometry {
  const shown = () => dom.zoom === layout.zoom && current();
  const painted = () => shown() && isPresented(dom.pagesContainer, queries.displayList);
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
  const stale = () => anchorFailure('stale-version', 'The document changed after that version');
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
  const pageOrigin = (pageIndex: number): DisplayListRect | null => {
    const bounds = queries.pageBounds(pageIndex);
    return bounds && { ...bounds, pageIndex, width: 0, height: 0 };
  };
  /**
   * Places display `ranges` less `hidden`. Unless `deferUnbuilt`, a fragment on a visible
   * unbuilt page refuses; with it, such fragments are left out and their pages listed.
   */
  const place = (
    ranges: Interval[],
    hidden: readonly Interval[],
    paragraph: number | null,
    deferUnbuilt: boolean
  ):
    | (AnchorPlacement & { ok: true; unbuiltPages: number[] })
    | Extract<DocxAnchorGeometryResult, { ok: false }> => {
    ranges.sort((a, b) => a.from - b.from || a.to - b.to);
    const window = displayWindowOf(queries)?.read();
    const pages = queries.displayList?.pages ?? [];
    const pendingPage = (pageIndex: number): boolean =>
      pages[pageIndex]?.unbuilt === true &&
      (deferUnbuilt || (!!window && pageIndex >= window[0] && pageIndex < window[1]));
    const reaches = (
      { unbuilt, positionSpan }: (typeof pages)[number],
      spans: readonly Interval[] = ranges
    ) =>
      !!unbuilt &&
      !!positionSpan &&
      spans.some(({ from, to }) =>
        from < to
          ? from < positionSpan[1] && to > positionSpan[0]
          : from >= positionSpan[0] && from <= positionSpan[1]
      );
    const unbuiltAt = (unit: number): number | undefined => {
      const index = pages.findIndex(
        ({ unbuilt, positionSpan }) =>
          unbuilt && positionSpan && unit >= positionSpan[0] && unit < positionSpan[1]
      );
      return index < 0 ? undefined : index;
    };
    const shownRanges = subtract(ranges, hidden);
    const drawable = shownRanges.filter(({ from, to }) => from < to);
    const unbuiltPages = deferUnbuilt
      ? pages.flatMap((page, pageIndex) => (reaches(page, drawable) ? [pageIndex] : []))
      : [];
    if (!deferUnbuilt && window && pages.slice(window[0], window[1]).some((page) => reaches(page))) return unavailable();
    const union: Interval[] = [];
    for (const range of shownRanges) {
      const previous = union.at(-1);
      if (previous && range.from < previous.to) previous.to = Math.max(previous.to, range.to);
      else union.push({ ...range });
    }
    const drawn: DisplayListRect[] = [];
    let tail: Interval | null = null;
    for (const { from, to } of union) {
      if (from >= to) continue;
      let placed = false;
      for (const rect of queries.rangeRects(from, to)) {
        if (rect.width <= 0) continue;
        if (pendingPage(rect.pageIndex)) {
          if (deferUnbuilt) continue;
          return unavailable();
        }
        drawn.push(rect);
        placed = true;
      }
      if (placed) tail = { from, to };
    }
    let anchor: DisplayListRect | null;
    const lastUnbuilt = unbuiltPages.at(-1);
    if (lastUnbuilt !== undefined && drawn.every(({ pageIndex }) => pageIndex < lastUnbuilt)) {
      anchor = pageOrigin(lastUnbuilt);
    } else {
      const last = ranges.at(-1);
      const gap = last && widen(last, hidden);
      const caret = tail === null && gap ? (caretAt(gap.from, true) ?? caretAt(gap.to)) : null;
      const caretPage =
        deferUnbuilt && tail === null && gap && !caret
          ? [gap.from - 1, gap.from, gap.to, gap.to - 1].map(unbuiltAt).find((index) => index !== undefined)
          : undefined;
      const end =
        tail !== null
          ? (endOf(tail) ?? lastInReadingOrder(drawn))
          : caretPage !== undefined
            ? null
            : (caret ?? (paragraph === null ? null : caretAt(paragraph)));
      const fallback =
        end || paragraph === null || caretPage !== undefined ? null : queries.anchorRect(paragraph);
      const pending =
        caretPage ?? [end, fallback].find((rect) => rect && pendingPage(rect.pageIndex))?.pageIndex;
      if (pending !== undefined && !deferUnbuilt) return unavailable();
      if (pending !== undefined && !unbuiltPages.includes(pending)) {
        unbuiltPages.push(pending);
        unbuiltPages.sort((a, b) => a - b);
      }
      anchor = pending !== undefined
        ? pageOrigin(pending)
        : (end ?? (fallback && { ...fallback, width: 0 }));
    }
    return anchor ? { ok: true, rects: drawn, anchor, unbuiltPages } : unavailable();
  };
  /** Projects a placement onto the pages as they stand now. */
  const answer = (
    { rects, anchor }: AnchorPlacement,
    unbuiltPages?: readonly number[]
  ): DocxAnchorGeometryResult => {
    const projected = rects.map(project);
    const anchorRect = project(anchor);
    const page = projector.getPageBounds(anchor.pageIndex);
    if (!anchorRect || !page || projected.some((rect) => rect === null)) return unavailable();
    return {
      ok: true,
      version: layout.version,
      previewVersion: layout.previewVersion,
      layoutId: layout.id,
      rects: projected as DocxAnchorRect[],
      anchor: anchorRect,
      pageRect: toOverlayRect(dom.pagesContainer, layer, dom.zoom, page),
      ...(unbuiltPages ? { unbuiltPages } : {}),
    };
  };
  const geometry: DocxPluginGeometry = {
    layout,
    dom,
    toOverlayRect: (rect) =>
      shown() || (dom.zoom === layout.zoom && held())
        ? toOverlayRect(dom.pagesContainer, layer, dom.zoom, rect)
        : null,
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
    getAnchorGeometry(target) {
      if (!shown()) return unavailable();
      const live = access();
      if (!live || !live.presented || live.editor.hasPendingInput()) return unavailable();
      const { session, editor } = live;
      if (session.version() !== layout.version) return stale();
      const snapshot = proposalSnapshot(session);
      if (
        (snapshot?.previewVersion ?? 0) !== layout.previewVersion ||
        currentPreviewKey(session) !== renderedPreviewKey(queries)
      )
        return unavailable();
      const mirror = live.proposalGeometry;
      if (
        mirror &&
        (target.kind !== 'proposal' ||
          mirror.version !== layout.version ||
          mirror.previewVersion !== layout.previewVersion ||
          !snapshot ||
          mirror.proposals !== proposalSetIdentity(snapshot))
      )
        return unavailable();
      const mirrored =
        mirror && target.kind === 'proposal'
          ? Object.hasOwn(mirror.targets, target.id)
            ? mirror.targets[target.id]
            : anchorFailure('unknown-proposal', 'The proposal is not registered in this document')
          : null;
      if (mirrored && !mirrored.ok) return mirrored;
      const resolved = mirror ? null : resolveAnchorTarget(session, target, layout.version);
      if (resolved && !resolved.ok) return resolved;
      const display = (range: RawAnchorRange): Interval | null => {
        const from = editor.yrsLocToDisplayPosition(range.start);
        const to = editor.yrsLocToDisplayPosition(range.end);
        return from === null || to === null ? null : { from, to };
      };
      const ranges: Interval[] = mirrored?.ok ? [...mirrored.ranges] : [];
      for (const range of resolved?.ok ? resolved.ranges : []) {
        const mapped = display(range);
        if (!mapped) return anchorFailure('unsupported', 'The target has no body display position');
        ranges.push(mapped);
      }
      const hidden =
        (mirror && (!hasEditorWorkerProposalRounds(session) || !workerOpenReplicaReady(session))
          ? mirror.hidden
          : undefined) ??
        hiddenRanges(session, layout.version)
          .map(display)
          .filter((range): range is Interval => range !== null);
      const paragraph = mirrored?.ok
        ? mirrored.paragraph
        : resolved?.ok
          ? editor.yrsLocToDisplayPosition(resolved.paragraph)
          : null;
      const placed = place(ranges, hidden, paragraph, false);
      return placed.ok ? answer(placed) : placed;
    },
    async readAnchorGeometry(target) {
      return (await geometry.readAnchorGeometries([target]))[0]!;
    },
    async readAnchorGeometries(targets) {
      if (!workerReads) return targets.map((target) => geometry.getAnchorGeometry(target));
      const version = painted() ? workerFrameVersionOf(queries) : null;
      const { read, cache } = workerReads;
      const previewKey = renderedPreviewKey(queries);
      const key = JSON.stringify([version, layout.previewVersion, previewKey]);
      if (version !== null && cache.key !== key) {
        Object.assign(cache, createAnchorReadCache(cache.limit), { key });
      }
      const { replies, pending } = cache;
      const entries = targets.map((target) => {
        if (version === null || target.kind === 'proposal') return null;
        if (target.kind === 'range' && target.version !== layout.version) return null;
        const sent = normalTarget(target.kind === 'range' ? { ...target, version } : target);
        return { id: JSON.stringify(sent), sent };
      });
      const missing = new Map<string, WorkerTarget>();
      for (const entry of entries) {
        if (entry && !replies.has(entry.id) && !pending.has(entry.id)) missing.set(entry.id, entry.sent);
      }
      if (version !== null && missing.size > 0) {
        const request = new Promise<readonly AnchorDisplayTarget[] | null>((resolve) =>
          resolve(read([...missing.values()], version, layout.previewVersion, previewKey))
        );
        [...missing.keys()].forEach((id, index) => {
          const reply = request.then((answers) => (answers ? answers[index]! : null), () => undefined);
          pending.set(id, reply);
          void reply.then(() => {
            if (pending.get(id) === reply) pending.delete(id);
          });
        });
      }
      const resolved = await Promise.all(
        entries.map((entry) => entry && (replies.get(entry.id) ?? pending.get(entry.id)))
      );
      const shown = painted();
      return targets.map((target, index) => {
        if (target.kind === 'proposal') return geometry.getAnchorGeometry(target);
        const entry = entries[index];
        if (!entry) return version === null ? unavailable() : stale();
        if (!shown) return unavailable();
        const reply = resolved[index];
        if (reply === undefined) return unavailable();
        if (reply === null) return stale();
        if (!reply.ok) return reply;
        if (cache.key === key && !cache.replies.has(entry.id)) {
          if (cache.replies.size >= cache.limit) cache.replies.clear();
          cache.replies.set(entry.id, reply);
        }
        const placement = place([...reply.ranges], reply.hidden, reply.paragraph, true);
        return placement.ok ? answer(placement, placement.unbuiltPages) : placement;
      });
    },
  };
  return geometry;
}
