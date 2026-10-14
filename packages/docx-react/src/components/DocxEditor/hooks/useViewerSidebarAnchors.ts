import { useEffect, useMemo, useRef } from 'react';
import type { Document } from '@betteroffice/docx/types/document';
import type { DisplayListQueries, TrackedChangesResult } from '@betteroffice/docx/layout/render';
import type { DocxSidebarRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { presentedWorkerVersion } from '../internals/layoutProvenance';
import {
  sidebarAnchorProjectY,
  SIDEBAR_ANCHOR_EMIT_MS,
  SIDEBAR_ANCHOR_STALE_MS,
} from '../internals/sidebarAnchorProjection';
import {
  ViewerSidebarReads,
  viewerCommentRanges,
  viewerSidebarPositions,
  type ViewerCommentRanges,
} from '../internals/viewerSidebarReads';
import { EMPTY_ANCHOR_POSITIONS } from '../commentFactories';

export function useViewerSidebarAnchors(options: {
  read?: ResidentEngineWorkerClient['documentRead'];
  queries: DisplayListQueries | null | undefined;
  commentIds: readonly (string | number)[];
  zoom: number;
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  overlayTarget?: HTMLElement | null;
  document?: Document | null;
  onPositions?: (positions: Map<string, number>) => void;
  onTracked?: (result: TrackedChangesResult) => void;
  onRanges?: (ranges: ViewerCommentRanges) => void;
}) {
  const latest = useRef(options);
  latest.current = options;
  const reads = useMemo(() => options.read ? new ViewerSidebarReads(options.read) : null, [options.read]);
  const lastPositions = useRef<Map<string, number> | null>(null);
  const lastRanges = useRef<{ value: DocxSidebarRead; ranges: ViewerCommentRanges } | null>(null);
  const commentIdsKey = JSON.stringify(options.commentIds.map(String));
  const lastEmitAt = useRef(0);
  const rangesVersion = useRef<string | null>(null);
  useEffect(() => {
    const { queries, canvasHostRef, pagesContainerRef, overlayTarget, zoom, document } = options;
    const commentIds = JSON.parse(commentIdsKey) as string[];
    const version = presentedWorkerVersion(queries);
    if (rangesVersion.current !== null && rangesVersion.current !== version) {
      rangesVersion.current = null;
      latest.current.onRanges?.(new Map());
    }
    if (!reads || !queries || version === null) return;
    let cancelled = false;
    let frame: number | null = null;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const current = () => !cancelled && latest.current.read === options.read &&
      presentedWorkerVersion(latest.current.queries) === version;
    const emit = async () => {
      if (!current()) return;
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      if (!host?.classList.contains('canvas-pages')) {
        frame = requestAnimationFrame(() => void emit());
        return;
      }
      const target = overlayTarget ?? host.parentElement;
      if (!target) return;
      const value = await reads.sidebar(version, commentIds, () =>
        presentedWorkerVersion(latest.current.queries));
      if (!value || !current()) return;
      reads.deliver(version, value, latest.current.onTracked);
      if (lastRanges.current?.value !== value) {
        lastRanges.current = { value, ranges: viewerCommentRanges(value) };
      }
      rangesVersion.current = version;
      latest.current.onRanges?.(lastRanges.current.ranges);
      const hfRegions = new Map<string, 'header' | 'footer'>();
      for (const rId of document?.package?.headers?.keys() ?? []) hfRegions.set(rId, 'header');
      for (const rId of document?.package?.footers?.keys() ?? []) {
        if (!hfRegions.has(rId)) hfRegions.set(rId, 'footer');
      }
      const positions = value.comments.length === 0 && value.revisions.length === 0
        ? EMPTY_ANCHOR_POSITIONS
        : viewerSidebarPositions(value, queries, hfRegions, sidebarAnchorProjectY(host, target, queries, zoom));
      const previous = lastPositions.current;
      if (previous && previous.size === positions.size &&
        [...positions].every(([key, y]) => previous.get(key) === y)) return;
      lastPositions.current = positions;
      latest.current.onPositions?.(positions);
    };
    const schedule = () => {
      if (!current() || timer !== null) return;
      if (performance.now() - lastEmitAt.current >= SIDEBAR_ANCHOR_STALE_MS) {
        lastEmitAt.current = performance.now();
        void emit();
        return;
      }
      timer = setTimeout(() => {
        timer = null;
        lastEmitAt.current = performance.now();
        void emit();
      }, SIDEBAR_ANCHOR_EMIT_MS);
    };
    schedule();
    void queries.whenReady().then(schedule, () => undefined);
    return () => {
      cancelled = true;
      if (frame !== null) cancelAnimationFrame(frame);
      if (timer !== null) clearTimeout(timer);
    };
  }, [reads, options.read, options.queries, commentIdsKey, options.zoom,
    options.canvasHostRef, options.pagesContainerRef, options.overlayTarget, options.document,
    options.onPositions, options.onTracked, options.onRanges]);
}
