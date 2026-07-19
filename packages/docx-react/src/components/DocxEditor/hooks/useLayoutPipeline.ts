/** Resident Rust layout scheduling and React paint-state publication. */

import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

import type { LayoutBlock, Layout, BlockExtent } from '@betteroffice/docx/layout/pagination';
import {
  buildResidentRegionLayoutRequest,
  computeLayout,
  type LayoutComputation,
} from '@betteroffice/docx/editor';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  ResidentFontRequirement,
  ResidentMeasurementConfig,
} from '@betteroffice/docx/layout';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';

import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { viewportMinHeightPx } from '../internals/scrollUtils';
import {
  captureDisplayListScrollAnchor,
  restoreDisplayListScrollAnchor,
  restoreScrollSnapshot,
  type DisplayListScrollAnchor,
} from '../internals/scrollRestore';

export interface UseLayoutPipelineOptions {
  document: Document | null;
  session: YrsSession | null;
  renderEnv: YrsRenderEnv;
  pageGap: number;
  zoom: number;
  residentMeasurementConfig: (
    requirements: ResidentFontRequirement[]
  ) => ResidentMeasurementConfig | null;
  /**
   * Rust measurement readiness gate (`useRustMeasurement.deferLayoutPass`).
   * Checked before computing (engine may still be loading) and before
   * committing (a pass may have discovered unresolved font chains); a
   * deferred pass is skipped/discarded and the measurement hook re-runs the
   * pipeline once the engine/fonts settle.
   */
  deferLayoutPass: (blocks?: LayoutBlock[]) => boolean;
  /**
   * True when the experimental canvas renderer is painting (the DOM painter is
   * parked in a 0×0 stage). Gates the painter-DOM–coupled steps of the paint
   * pass off: the `RenderedDomContext` is re-backed by the a11y mirror and the
   * HF sidebar-anchor scan runs against the mirror — both owned by PagedEditor's
   * canvas effect. Default (false) leaves the painter path byte-for-byte
   * unchanged.
   */
  displayListQueries?: DisplayListQueries | null;
  interactionPageHostRef?: React.RefObject<HTMLDivElement | null>;
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  viewportLayoutRef: React.RefObject<HTMLDivElement | null>;
  /** Current display-list position used to preserve the scroll anchor. */
  getSelectionHead?: () => number;
  syncCoordinator: LayoutSelectionGate;
  getScrollContainer: () => HTMLDivElement | null;
  onTotalPagesChange?: (totalPages: number) => void;
  /** Layout of each pass (null on reset) — experimental canvas renderer plumbing. */
  onLayoutComputed?: (layout: Layout | null) => void;
  onAnchorPositionsChange?: (positions: Map<string, number>) => void;
}

export interface UseLayoutPipelineReturn {
  layout: Layout | null;
  blocks: LayoutBlock[];
  measures: BlockExtent[];
  runLayoutPipeline: () => void;
  scheduleLayout: () => void;
}

export function useLayoutPipeline(opts: UseLayoutPipelineOptions): UseLayoutPipelineReturn {
  const {
    document,
    session,
    renderEnv,
    pageGap,
    zoom,
    residentMeasurementConfig,
    deferLayoutPass,
    displayListQueries,
    interactionPageHostRef,
    pagesContainerRef,
    viewportLayoutRef,
    getSelectionHead,
    syncCoordinator,
    getScrollContainer,
    onTotalPagesChange,
    onLayoutComputed,
    onAnchorPositionsChange,
  } = opts;

  const [layout, setLayout] = useState<Layout | null>(null);
  const [blocks, setBlocks] = useState<LayoutBlock[]>([]);
  const [measures, setMeasures] = useState<BlockExtent[]>([]);

  // Callback refs — parent may hand in a fresh closure every render. Mirroring
  // these in refs keeps `runLayoutPipeline`'s dep array stable; otherwise
  // every parent re-render would invalidate the rAF-coalesced scheduler.
  const onTotalPagesChangeRef = useRef(onTotalPagesChange);
  const onLayoutComputedRef = useRef(onLayoutComputed);
  const onAnchorPositionsChangeRef = useRef(onAnchorPositionsChange);
  // Query facades are immutable per display-list build. Reading the current
  // facade through a ref keeps runLayoutPipeline identity-stable when a build
  // lands; otherwise useLayoutTriggers sees a new callback, starts another
  // layout, and creates a display-list -> layout feedback loop.
  const displayListQueriesRef = useRef(displayListQueries);
  const deferLayoutPassRef = useRef(deferLayoutPass);
  onTotalPagesChangeRef.current = onTotalPagesChange;
  onLayoutComputedRef.current = onLayoutComputed;
  onAnchorPositionsChangeRef.current = onAnchorPositionsChange;
  displayListQueriesRef.current = displayListQueries;
  deferLayoutPassRef.current = deferLayoutPass;

  // Total-pages notifier — fires only when count changes (including N → 0).
  const lastTotalPagesRef = useRef<number>(0);
  useEffect(() => {
    onLayoutComputedRef.current?.(layout);
    const total = layout?.pages.length ?? 0;
    if (total === lastTotalPagesRef.current) return;
    lastTotalPagesRef.current = total;
    onTotalPagesChangeRef.current?.(total);
  }, [layout]);

  // Scroll-restore plumbing. `pendingScrollRestoreRef` is read by both the
  // pipeline and the post-commit useLayoutEffect below.
  const pendingScrollRestoreRef = useRef<DisplayListScrollAnchor | null>(null);

  // =========================================================================
  // Layout Pipeline
  // =========================================================================

  const runLayoutPipeline = useCallback(
    () => {
      const pipelineStart = performance.now();

      const currentEpoch = syncCoordinator.getStateSeq();
      syncCoordinator.onLayoutStart();

      if (deferLayoutPassRef.current() || !session) {
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      let measurement: ResidentMeasurementConfig | null = null;
      try {
        const request = buildResidentRegionLayoutRequest(document, pageGap, renderEnv);
        const requirements = JSON.parse(
          session.layoutFontRequirementsJson(JSON.stringify(request))
        ) as ResidentFontRequirement[];
        measurement = residentMeasurementConfig(requirements);
      } catch (error) {
        console.error('[PagedEditor] Resident font preflight error:', error);
      }
      if (!measurement) {
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      const computeInputs = { document, pageGap, session, renderEnv, measurement };

      // Step 4+: paint + scroll/events with the computed values.
      const applyComputation = (computation: LayoutComputation) => {
        const { blocks: newBlocks, measures: newMeasures, layout: newLayout } = computation;

        const pagesEl = pagesContainerRef.current;
        const scrollParent =
          getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
        const interactionHost = interactionPageHostRef?.current ?? pagesEl;
        const queries = displayListQueriesRef.current;
        const anchor =
          scrollParent?.isConnected && interactionHost && queries
            ? captureDisplayListScrollAnchor(
                queries,
                interactionHost,
                scrollParent,
                getSelectionHead?.() ?? 0
              )
            : null;

        setBlocks(newBlocks);
        setMeasures(newMeasures);
        setLayout(newLayout);

        const vp = viewportLayoutRef.current;
        if (vp) {
          const mh = viewportMinHeightPx(newLayout, pageGap);
          vp.style.minHeight = `${mh}px`;
          vp.style.marginBottom = zoom !== 1 ? `${mh * (zoom - 1)}px` : '';
        }
        pendingScrollRestoreRef.current = scrollParent?.isConnected ? anchor : null;

        const totalTime = performance.now() - pipelineStart;
        if (totalTime > 2000) {
          console.warn(
            `[PagedEditor] Layout pipeline took ${Math.round(totalTime)}ms total ` +
              `(${newBlocks.length} blocks, ${newMeasures.length} measures)`
          );
        }
      };

      // The wasm engine paginates synchronously in native code — every pass
      // is a full relayout (the Rust kernel keeps no resume checkpoints; the
      // TS incremental/off-thread kernels died with TS pagination).
      try {
        applyComputation(computeLayout(computeInputs));
      } catch (error) {
        console.error('[PagedEditor] Layout pipeline error:', error);
      }
      syncCoordinator.onLayoutComplete(currentEpoch);
    },
    [
      pageGap,
      zoom,
      syncCoordinator,
      document,
      session,
      renderEnv,
      residentMeasurementConfig,
      getScrollContainer,
      getSelectionHead,
      interactionPageHostRef,
      pagesContainerRef,
      viewportLayoutRef,
    ]
  );

  // Hold the exact scrollTop while the next display-list commit is built.
  useLayoutEffect(() => {
    const pending = pendingScrollRestoreRef.current;
    if (!pending) return;
    const pagesEl = pagesContainerRef.current;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (scrollParent?.isConnected) restoreScrollSnapshot(pending, scrollParent);
  }, [layout, getScrollContainer, pagesContainerRef]);

  // A new immutable display-list/query facade is the geometry commit signal.
  useLayoutEffect(() => {
    const pending = pendingScrollRestoreRef.current;
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!pending || !displayListQueries || !host || !scrollParent?.isConnected) return;
    pendingScrollRestoreRef.current = null;
    restoreDisplayListScrollAnchor(pending, displayListQueries, host, scrollParent);
    const rafId = requestAnimationFrame(() => {
      if (scrollParent.isConnected) {
        restoreDisplayListScrollAnchor(pending, displayListQueries, host, scrollParent);
      }
    });
    return () => cancelAnimationFrame(rafId);
  }, [displayListQueries, getScrollContainer, interactionPageHostRef, pagesContainerRef]);

  // =========================================================================
  // Coalesced Layout (rAF throttle)
  // =========================================================================

  /**
   * Multiple rapid transactions (e.g. typing "hello") within the same frame
   * are coalesced so only the final state triggers a full layout pass. The
   * coalescer lives in core (`createLayoutScheduler`) so React and Vue share
   * it; the `runRef` indirection lets the stable scheduler always call the
   * latest `runLayoutPipeline` without recreating itself.
   */
  const runRef = useRef(runLayoutPipeline);
  runRef.current = runLayoutPipeline;
  const schedulerRef = useRef<number | null>(null);
  const scheduleLayout = useCallback(() => {
    if (schedulerRef.current != null) return;
    schedulerRef.current = requestAnimationFrame(() => {
      schedulerRef.current = null;
      runRef.current();
    });
  }, []);

  // Clean up pending rAF on unmount
  useEffect(() => {
    return () => {
      if (schedulerRef.current != null) cancelAnimationFrame(schedulerRef.current);
    };
  }, []);

  return {
    layout,
    blocks,
    measures,
    runLayoutPipeline,
    scheduleLayout,
  };
}
