/**
 * Comment / tracked-change anchor derivation over the display list, shared
 * by both adapters' sidebar plumbing on the canvas-renderer path.
 *
 * `visitAnchorKeys` walks the PM doc and reports every comment /
 * tracked-change anchor as `(key, pos)` — shared between each adapter's
 * layout-math derivation (DOM painter path) and the display-list derivation
 * here, so both register identical keys in identical order.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import type { YrsRevisionInfo, YrsSession } from '../../yrs';
import type { DisplayListQueries, DisplayListRect } from './displayListQueries';
import { canvasPageTops } from './canvasPageMetrics';
import {
  type YrsSidebarDisplayPoint,
  type YrsSidebarProjection,
  yrsIdToNumericId,
} from './yrsSidebarProjection';

/** Structural view shape used only by the legacy display-list anchor walker. */
export interface AnchorEditorView {
  readonly state: {
    readonly doc: {
      descendants(callback: (node: any, pos: number) => boolean | void): void;
    };
    readonly schema: { readonly marks: Record<string, any> };
  };
}

/**
 * Walk the PM doc and report every comment / tracked-change anchor as
 * `(key, pos)`. `register` is called once per occurrence; callers dedupe.
 */
export function visitAnchorKeys(
  pmView: AnchorEditorView,
  register: (key: string, pos: number) => void
): void {
  const { doc: pmDoc, schema } = pmView.state;
  const commentType = schema.marks.comment;
  const insertionType = schema.marks.insertion;
  const deletionType = schema.marks.deletion;
  if (!commentType && !insertionType && !deletionType) return;

  pmDoc.descendants((node, pos) => {
    // Structural tracked-change attrs on non-text nodes (whole-table insert,
    // row insert/delete, cell insert, paragraph-break tracked, etc). Without
    // these, an empty inserted table has no anchor — the sidebar's
    // hasPositions check stays false and the whole rail renders at opacity 0.
    //
    // The attrs use three different shapes for the revisionId:
    //   • flat       — trIns / trDel / pPrIns / pPrDel: `{ revisionId, ... }`
    //   • nested     — cellMarker: `{ kind, info: { revisionId, ... } }`
    //   • array+info — *PrChange (paragraph/row/cell/table): `[{ info: { id } }, ...]`
    // Pre-fix all three by extracting the revisionId at registration time.
    const attrs = node.attrs as Record<string, unknown> | undefined;
    if (attrs) {
      const flat = [attrs.trIns, attrs.trDel, attrs.pPrIns, attrs.pPrDel];
      for (const entry of flat) {
        const revId = (entry as { revisionId?: unknown } | null | undefined)?.revisionId;
        if (typeof revId === 'number') register(`revision-${revId}`, pos);
      }
      const cellMarker = attrs.cellMarker as { info?: { revisionId?: unknown } } | null;
      const cellRev = cellMarker?.info?.revisionId;
      if (typeof cellRev === 'number') register(`revision-${cellRev}`, pos);
      const propChangeArrays = [
        attrs.pPrChange,
        attrs.trPrChange,
        attrs.tcPrChange,
        attrs.tblPrChange,
      ];
      for (const arr of propChangeArrays) {
        if (!Array.isArray(arr)) continue;
        for (const entry of arr as Array<{ info?: { id?: unknown } }>) {
          const id = entry?.info?.id;
          if (typeof id === 'number') register(`revision-${id}`, pos);
        }
      }
    }

    // Text AND inline atoms (image, shape) can carry comment / tracked-change
    // marks, so an inserted picture's card gets a sidebar anchor like inserted
    // text. Without this an image-only change has no positioned card.
    if (!node.isInline) return;
    for (const mark of node.marks) {
      let key: string | null = null;
      if (commentType && mark.type === commentType) {
        key = `comment-${mark.attrs.commentId}`;
      } else if (
        (insertionType && mark.type === insertionType) ||
        (deletionType && mark.type === deletionType)
      ) {
        key = `revision-${mark.attrs.revisionId}`;
      }
      if (!key) continue;
      register(key, pos);
    }
  });
}

/**
 * Display-list variant of the adapters' layout-based anchor derivation for
 * the experimental canvas renderer: anchor Ys come from Rust `range_rects`
 * queries over the display list (page-local px) plus the canvas page stack
 * math, instead of layout-fragment scans in the DOM viewport's coordinate
 * space. Returns `.canvas-pages`-host Y offsets keyed by
 * `comment-{id}` / `revision-{revisionId}`.
 */
export function computeAnchorPositionsFromDisplayList(
  pmView: AnchorEditorView | null,
  queries: DisplayListQueries,
  projectY?: (rect: DisplayListRect) => number | null
): Map<string, number> {
  const positions = new Map<string, number>();
  if (!pmView?.state) return positions;

  const pageTops = canvasPageTops(queries);
  const seen = new Set<string>();
  visitAnchorKeys(pmView, (key, pos) => {
    if (seen.has(key)) return;
    seen.add(key);
    const rect = queries.anchorRect(pos);
    if (!rect) return;
    const y = projectY ? projectY(rect) : (pageTops[rect.pageIndex] ?? 0) + rect.y;
    if (y != null) positions.set(key, y);
  });

  return positions;
}

export interface YrsHeaderFooterRegions {
  get(rId: string): 'header' | 'footer' | undefined;
}

function rectForYrsPoint(
  point: YrsSidebarDisplayPoint,
  queries: DisplayListQueries,
  hfRegions?: YrsHeaderFooterRegions
): DisplayListRect | null {
  if (!point.hfRid) return queries.anchorRect(point.position);
  const region = hfRegions?.get(point.hfRid);
  return region ? (queries.hfAnchorRects(region, point.hfRid, point.position)[0] ?? null) : null;
}

/**
 * yrs-authoritative sidebar anchors. Comment coverage comes from sticky yrs
 * side-map positions and tracked-change anchors come from `listRevisions()`;
 * neither path reads an EditorView or ProseMirror document.
 */
export function computeAnchorPositionsFromYrs(
  session: YrsSession,
  commentIds: Iterable<string | number>,
  revisions: readonly YrsRevisionInfo[],
  projection: YrsSidebarProjection,
  queries: DisplayListQueries,
  hfRegions?: YrsHeaderFooterRegions,
  projectY?: (rect: DisplayListRect) => number | null
): Map<string, number> {
  const positions = new Map<string, number>();
  const pageTops = canvasPageTops(queries);

  const register = (key: string, point: YrsSidebarDisplayPoint | null): boolean => {
    if (!point || positions.has(key)) return positions.has(key);
    const rect = rectForYrsPoint(point, queries, hfRegions);
    if (!rect) return false;
    const y = projectY ? projectY(rect) : (pageTops[rect.pageIndex] ?? 0) + rect.y;
    if (y == null) return false;
    positions.set(key, y);
    return true;
  };

  for (const commentId of commentIds) {
    const key = `comment-${commentId}`;
    try {
      for (const anchor of session.resolveComment(String(commentId))) {
        if (register(key, projection.storyOffsetToDisplayPoint(anchor.story, anchor.start))) break;
      }
    } catch {
      // A deleted/orphaned comment has no live yrs anchor and remains unplaced.
    }
  }

  for (const revision of revisions) {
    register(
      `revision-${yrsIdToNumericId(revision.revisionId)}`,
      projection.locToDisplayPoint({ story: revision.story, ...revision.range.start })
    );
  }

  return positions;
}

export interface DisplayListHfAnchorSource {
  region: 'header' | 'footer';
  rId: string;
  view: AnchorEditorView;
}

/** Merge header/footer comment and revision anchors into an existing map. */
export function mergeHfAnchorPositionsFromDisplayList(
  sources: Iterable<DisplayListHfAnchorSource>,
  queries: DisplayListQueries,
  positions: Map<string, number>,
  projectY?: (rect: DisplayListRect) => number | null
): Map<string, number> {
  const pageTops = canvasPageTops(queries);
  for (const { region, rId, view } of sources) {
    visitAnchorKeys(view, (key, pos) => {
      if (positions.has(key)) return;
      const rect = queries.hfAnchorRects(region, rId, pos)[0];
      if (!rect) return;
      const y = projectY ? projectY(rect) : (pageTops[rect.pageIndex] ?? 0) + rect.y;
      if (y != null) positions.set(key, y);
    });
  }
  return positions;
}
