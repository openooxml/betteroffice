/**
 * Layout pipeline hook for PagedEditor.
 *
 * Owns the 4-step layout pass (PM doc → flow blocks → measure → layout →
 * paint), its rAF-coalesced scheduler, and the scroll-restore state that
 * keeps the user's scroll position locked across re-paints.
 *
 * Extraction note: every line of `runLayoutPipeline` moves in here
 * verbatim. The LayoutBlock invariant (`assertExhaustiveLayoutBlock` in the
 * `toLayoutBlocks` chain via `measureBlock.ts`) depends on this site staying
 * stable — if a new LayoutBlock variant is added, the three measureBlock
 * switches still need updates per the LayoutBlock invariant (see the repo guidelines).
 */

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';

import type {
  LayoutBlock,
  Layout,
  BlockExtent,
  LayoutPaginationSource,
} from '@betteroffice/docx/layout/pagination';
import { getMargins, getPageSize, getColumns } from '@betteroffice/docx/layout';
import {
  computeLayout,
  type LayoutComputation,
} from '@betteroffice/docx/editor';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type {
  Document,
  HeaderFooter,
  SectionProperties,
  StyleDefinitions,
  Theme,
} from '@betteroffice/docx/types/document';

import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { FloatPageGeometry } from '@betteroffice/docx/layout';
import { viewportMinHeightPx } from '../internals/scrollUtils';
import {
  captureDisplayListScrollAnchor,
  restoreDisplayListScrollAnchor,
  restoreScrollSnapshot,
  type DisplayListScrollAnchor,
} from '../internals/scrollRestore';
import { tracePerfSync } from '../internals/perfTrace';

export interface UseLayoutPipelineOptions {
  document: Document | null;
  styles?: StyleDefinitions | null;
  theme?: Theme | null;
  sectionProperties?: SectionProperties | null;
  finalSectionProperties?: SectionProperties | null;
  headerContent?: HeaderFooter | null;
  footerContent?: HeaderFooter | null;
  firstPageHeaderContent?: HeaderFooter | null;
  firstPageFooterContent?: HeaderFooter | null;
  /** Body-block source for the yrs-owned React editing session. */
  yrsBodyBlocks?: (env: { pageContentHeight: number }) => LayoutBlock[] | null;
  /** Non-body story source for the yrs-owned React editing session. */
  yrsStoryBlocks?: (storyId: string) => LayoutBlock[] | null;
  pageGap: number;
  zoom: number;
  /**
   * The Rust measurement implementation (`useRustMeasurement.measureBlocksImpl`)
   * — the sole measurement path. Must be identity-stable.
   */
  measureBlocksImpl: (
    blocks: LayoutBlock[],
    contentWidth: number | number[],
    pageGeometry?: FloatPageGeometry
  ) => BlockExtent[];
  /**
   * Rust measurement readiness gate (`useRustMeasurement.deferLayoutPass`).
   * Checked before computing (engine may still be loading) and before
   * committing (a pass may have discovered unresolved font chains); a
   * deferred pass is skipped/discarded and the measurement hook re-runs the
   * pipeline once the engine/fonts settle.
   */
  deferLayoutPass: (blocks?: LayoutBlock[]) => boolean;
  /**
   * The mandatory Rust pagination engine (see `useRustPagination`). Must be
   * identity-stable. Passes are deferred while the wasm engine loads
   * (`isReady()`), and the pipeline re-runs once `whenReady()` settles —
   * the same defer-and-rerun pattern as the measurement readiness gate.
   */
  paginationSource: LayoutPaginationSource;
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
  contentWidth: number;
  runLayoutPipeline: () => void;
  scheduleLayout: () => void;
}

export function useLayoutPipeline(opts: UseLayoutPipelineOptions): UseLayoutPipelineReturn {
  const {
    document,
    styles,
    theme,
    sectionProperties,
    finalSectionProperties,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    yrsBodyBlocks,
    yrsStoryBlocks,
    pageGap,
    zoom,
    measureBlocksImpl,
    deferLayoutPass,
    paginationSource,
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
  const yrsBodyBlocksRef = useRef(yrsBodyBlocks);
  const yrsStoryBlocksRef = useRef(yrsStoryBlocks);
  // Query facades are immutable per display-list build. Reading the current
  // facade through a ref keeps runLayoutPipeline identity-stable when a build
  // lands; otherwise useLayoutTriggers sees a new callback, starts another
  // layout, and creates a display-list -> layout feedback loop.
  const displayListQueriesRef = useRef(displayListQueries);
  const deferLayoutPassRef = useRef(deferLayoutPass);
  onTotalPagesChangeRef.current = onTotalPagesChange;
  onLayoutComputedRef.current = onLayoutComputed;
  onAnchorPositionsChangeRef.current = onAnchorPositionsChange;
  yrsBodyBlocksRef.current = yrsBodyBlocks;
  yrsStoryBlocksRef.current = yrsStoryBlocks;
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

  // Page geometry derived from section properties.
  const pageSize = useMemo(() => getPageSize(sectionProperties), [sectionProperties]);
  const margins = useMemo(() => getMargins(sectionProperties), [sectionProperties]);
  const columns = useMemo(() => getColumns(sectionProperties), [sectionProperties]);
  const { finalPageSize, finalMargins, finalColumns } = useMemo(() => {
    const props = finalSectionProperties ?? sectionProperties;
    return {
      finalPageSize: getPageSize(props),
      finalMargins: getMargins(props),
      finalColumns: getColumns(props),
    };
  }, [finalSectionProperties, sectionProperties]);
  const contentWidth = pageSize.w - margins.left - margins.right;

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

      // Rust engine readiness pre-gate: while the measurement engine is
      // still loading (or fonts a previous pass needed are still resolving),
      // or the pagination engine's wasm hasn't settled, skip the pass
      // entirely — the measurement hook / the whenReady effect below re-run
      // the pipeline once the engines settle.
      if (deferLayoutPassRef.current() || !paginationSource.isReady()) {
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      // Steps 1-3 (PM doc → blocks → measure → HF resolve → margin extend →
      // layout → footnote items) are the shared compute pass, lifted to
      // `@betteroffice/docx/editor`. Paint + scroll/events stay here.
      const tracedPaginationSource: LayoutPaginationSource = {
        isReady: () => paginationSource.isReady(),
        whenReady: () => paginationSource.whenReady(),
        paginate: (measured, options) =>
          tracePerfSync('paginate', () => paginationSource.paginate(measured, options), {
            calls: 1,
            detail: `${measured.length} blocks`,
          }),
      };
      const computeInputs = {
        state: null,
        document,
        pageSize,
        margins,
        columns,
        finalPageSize,
        finalMargins,
        finalColumns,
        pageGap,
        contentWidth,
        theme,
        styles,
        sectionProperties,
        finalSectionProperties,
        headerContent,
        footerContent,
        firstPageHeaderContent,
        firstPageFooterContent,
        yrsBodyBlocks: yrsBodyBlocksRef.current,
        yrsStoryBlocks: yrsStoryBlocksRef.current,
        measureBlocks: (
          inputBlocks: LayoutBlock[],
          inputWidth: number | number[],
          pageGeometry?: FloatPageGeometry
        ) =>
          tracePerfSync(
            'measure',
            () => measureBlocksImpl(inputBlocks, inputWidth, pageGeometry),
            { calls: 1, detail: `${inputBlocks.length} blocks` }
          ),
        paginationSource: tracedPaginationSource,
      };

      // Step 4+: paint + scroll/events with the computed values.
      const applyComputation = (computation: LayoutComputation) => {
        // Readiness post-gate: the measure pass may have discovered font
        // chains that are still resolving (HF/footnote fonts especially).
        // Such a pass is provisional — some paragraphs carry synthetic
        // extents — so discard it; the measurement hook re-runs the pipeline
        // when the chains settle.
        if (deferLayoutPassRef.current(computation.blocks)) return;

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
      contentWidth,
      columns,
      pageSize,
      margins,
      finalPageSize,
      finalMargins,
      finalColumns,
      pageGap,
      zoom,
      syncCoordinator,
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
      sectionProperties,
      finalSectionProperties,
      document,
      measureBlocksImpl,
      paginationSource,
      getScrollContainer,
      getSelectionHead,
      interactionPageHostRef,
      pagesContainerRef,
      styles,
      theme,
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

  // Pagination readiness — un-defer the passes that ran before the wasm
  // engine loaded. The module is shared with the measurement engine, so this
  // settles at (or before) the moment measurement un-defers; the re-run is
  // correctness insurance for the boot race, mirroring useRustMeasurement.
  useEffect(() => {
    if (paginationSource.isReady()) return;
    let cancelled = false;
    paginationSource
      .whenReady()
      .then(() => {
        if (cancelled) return;
        runRef.current();
      })
      .catch((error) => {
        console.error('[PagedEditor] Rust pagination engine failed to load — no layout', error);
      });
    return () => {
      cancelled = true;
    };
  }, [paginationSource]);

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
    contentWidth,
    runLayoutPipeline,
    scheduleLayout,
  };
}
