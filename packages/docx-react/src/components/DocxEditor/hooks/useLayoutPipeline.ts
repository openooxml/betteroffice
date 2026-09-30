/** Resident Rust layout scheduling and React paint-state publication. */

import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

import type { LayoutBlock, Layout } from '@betteroffice/docx/layout/pagination';
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
import {
  ResidentWorkerOutOfMemoryError,
  type YrsLoc,
  type YrsRenderEnv,
  type YrsSession,
  type YrsStickyPosition,
} from '@betteroffice/docx/yrs';

import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import { documentPageCount } from './documentPageCount';
import type { LayoutInWorker } from './useDisplayList';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { viewportMinHeightPx } from '../internals/scrollUtils';
import {
  readSessionVersion,
  revisionPreviewKey,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../internals/layoutProvenance';
import {
  captureDisplayListScrollAnchor,
  captureDisplayListViewportAnchor,
  restoreDisplayListScrollAnchor,
  restoreDisplayListViewportAnchor,
  restoreScrollSnapshot,
  type DisplayListScrollAnchor,
  type DisplayListViewportAnchor,
} from '../internals/scrollRestore';
import {
  mergeLayoutUpdateOrigin,
  PendingScrollRestoreController,
  type LayoutUpdateOrigin,
} from '../internals/viewportAnchoring';

interface SelectionScrollRestore {
  kind: 'selection';
  anchor: DisplayListScrollAnchor;
}

interface ViewportScrollRestore {
  kind: 'viewport';
  anchor: DisplayListViewportAnchor;
}

type PendingScrollRestore = SelectionScrollRestore | ViewportScrollRestore;

interface CurrentViewportAnchor {
  anchor: DisplayListViewportAnchor;
  navigationEpoch: number;
}

export interface UseLayoutPipelineOptions {
  onError?: (error: Error) => void;
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
  /** Queries available while the canvas renderer is active. */
  displayListQueries?: DisplayListQueries | null;
  interactionPageHostRef?: React.RefObject<HTMLDivElement | null>;
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  viewportLayoutRef: React.RefObject<HTMLDivElement | null>;
  /** Current display-list position used to preserve the scroll anchor. */
  getSelectionHead?: () => number;
  displayPositionToYrsLoc?: (position: number) => YrsLoc | null;
  yrsLocToDisplayPosition?: (loc: YrsLoc) => number | null;
  syncCoordinator: LayoutSelectionGate;
  getScrollContainer: () => HTMLDivElement | null;
  onTotalPagesChange?: (totalPages: number) => void;
  /** Receives each computed layout and resets with null. */
  onLayoutComputed?: (layout: Layout | null) => void;
  /** Hands passes to the resident worker, which then runs the only layout. */
  layoutInWorker?: LayoutInWorker;
  onAnchorPositionsChange?: (positions: Map<string, number>) => void;
}

export interface UseLayoutPipelineReturn {
  layout: Layout | null;
  layoutUpdateOrigin: LayoutUpdateOrigin;
  /** `onHost` lays out on this thread even when a worker could. */
  runLayoutPipeline: (options?: { onHost?: boolean }) => void;
  scheduleLayout: (origin?: LayoutUpdateOrigin) => void;
  cancelPendingScrollRestore: () => void;
  /** Counts navigation intents, the user's and programmatic scrolls alike. */
  navigationEpoch: () => number;
  /**
   * The region layout request the pipeline would lay the current document out with now, or
   * `null` while it has no session or the fonts the document needs are not ready.
   */
  getLayoutRequest: () => string | null;
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
    displayPositionToYrsLoc,
    yrsLocToDisplayPosition,
    syncCoordinator,
    getScrollContainer,
    onError,
    onTotalPagesChange,
    onLayoutComputed,
    layoutInWorker,
    onAnchorPositionsChange,
  } = opts;

  const [layout, setLayout] = useState<Layout | null>(null);

  // Callback refs — parent may hand in a fresh closure every render. Mirroring
  // these in refs keeps `runLayoutPipeline`'s dep array stable; otherwise
  // every parent re-render would invalidate the rAF-coalesced scheduler.
  const onErrorRef = useRef(onError);
  onErrorRef.current = onError;
  const onTotalPagesChangeRef = useRef(onTotalPagesChange);
  const onLayoutComputedRef = useRef(onLayoutComputed);
  const onAnchorPositionsChangeRef = useRef(onAnchorPositionsChange);
  // Query facades are immutable per display-list build. Reading the current
  // facade through a ref keeps runLayoutPipeline identity-stable when a build
  // lands; otherwise useLayoutTriggers sees a new callback, starts another
  // layout, and creates a display-list -> layout feedback loop.
  const displayListQueriesRef = useRef(displayListQueries);
  const deferLayoutPassRef = useRef(deferLayoutPass);
  const displayPositionToYrsLocRef = useRef(displayPositionToYrsLoc);
  const yrsLocToDisplayPositionRef = useRef(yrsLocToDisplayPosition);
  const layoutInWorkerRef = useRef(layoutInWorker);
  layoutInWorkerRef.current = layoutInWorker;
  // Bumped by every pass, so a worker pass answering late never overwrites a
  // newer layout.
  const passRef = useRef(0);
  const sessionRef = useRef(session);
  sessionRef.current = session;
  // The document version the first pass of this session laid out.
  const openedVersionRef = useRef<{ session: YrsSession; version: string | null } | null>(null);
  // A deferred pass that had to run on this thread keeps that requirement.
  const pendingOnHostRef = useRef(false);
  onTotalPagesChangeRef.current = onTotalPagesChange;
  onLayoutComputedRef.current = onLayoutComputed;
  onAnchorPositionsChangeRef.current = onAnchorPositionsChange;
  displayListQueriesRef.current = displayListQueries;
  deferLayoutPassRef.current = deferLayoutPass;
  displayPositionToYrsLocRef.current = displayPositionToYrsLoc;
  yrsLocToDisplayPositionRef.current = yrsLocToDisplayPosition;

  // Total-pages notifier — fires only when count changes (including N → 0).
  const lastTotalPagesRef = useRef<number>(0);
  useEffect(() => {
    onLayoutComputedRef.current?.(layout);
    const total = documentPageCount(layout);
    if (total === lastTotalPagesRef.current) return;
    lastTotalPagesRef.current = total;
    onTotalPagesChangeRef.current?.(total);
  }, [layout]);

  const scrollRestoreControllerRef =
    useRef<PendingScrollRestoreController<PendingScrollRestore> | null>(null);
  if (!scrollRestoreControllerRef.current) {
    scrollRestoreControllerRef.current =
      new PendingScrollRestoreController<PendingScrollRestore>();
  }
  const scrollRestoreController = scrollRestoreControllerRef.current;
  const navigationEpochRef = useRef(0);
  const cancelPendingScrollRestore = useCallback(() => {
    navigationEpochRef.current += 1;
    scrollRestoreController.cancel();
  }, [scrollRestoreController]);
  const currentViewportAnchorRef = useRef<CurrentViewportAnchor | null>(null);
  const viewportAnchorSessionRef = useRef(session);
  const viewportAnchorCaptureReadyRef = useRef(true);
  const expectedScrollTopRef = useRef<number | null>(null);
  if (viewportAnchorSessionRef.current !== session) {
    viewportAnchorSessionRef.current = session;
    currentViewportAnchorRef.current = null;
  }
  const pendingLayoutOriginRef = useRef<LayoutUpdateOrigin | null>(null);
  const layoutUpdateOriginRef = useRef<LayoutUpdateOrigin>('local');

  const captureViewportPosition = useCallback(
    (position: number) => {
      const loc = displayPositionToYrsLocRef.current?.(position);
      // Round-trip guard: a projection that cannot map the loc back to the same
      // display position would resolve the anchor into another story on restore.
      if (!session || !loc || yrsLocToDisplayPositionRef.current?.(loc) !== position) return null;
      try {
        return session.encodeStickyPosition(loc);
      } catch {
        return null;
      }
    },
    [session]
  );
  const resolveViewportPosition = useCallback(
    (position: YrsStickyPosition) => {
      const loc = session?.resolveStickyPosition(position);
      if (!loc) return null;
      return yrsLocToDisplayPositionRef.current?.(loc) ?? null;
    },
    [session]
  );
  const captureCurrentViewportAnchor = useCallback(() => {
    const queries = displayListQueriesRef.current;
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    const anchor =
      queries && host && scrollParent?.isConnected
        ? captureDisplayListViewportAnchor(queries, host, scrollParent, captureViewportPosition)
        : null;
    currentViewportAnchorRef.current = anchor
      ? { anchor, navigationEpoch: navigationEpochRef.current }
      : null;
  }, [
    captureViewportPosition,
    getScrollContainer,
    interactionPageHostRef,
    pagesContainerRef,
  ]);

  // =========================================================================
  // Layout Pipeline
  // =========================================================================

  const runLayoutPipeline = useCallback(
    (options?: { onHost?: boolean }) => {
      const pass = ++passRef.current;
      const onHost = options?.onHost === true || pendingOnHostRef.current;
      const layoutUpdateOrigin = pendingLayoutOriginRef.current ?? 'local';
      pendingLayoutOriginRef.current = null;
      if (layoutUpdateOrigin === 'local') scrollRestoreController.cancel();
      const pipelineStart = performance.now();

      const currentEpoch = syncCoordinator.getStateSeq();
      syncCoordinator.onLayoutStart();

      if (deferLayoutPassRef.current() || !session) {
        pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
          pendingLayoutOriginRef.current,
          layoutUpdateOrigin
        );
        pendingOnHostRef.current = onHost;
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
        onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }
      if (!measurement) {
        pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
          pendingLayoutOriginRef.current,
          layoutUpdateOrigin
        );
        pendingOnHostRef.current = onHost;
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }
      pendingOnHostRef.current = false;

      const computeInputs = { document, pageGap, session, renderEnv, measurement };
      const sourceVersion = readSessionVersion(session);
      const previewKey = revisionPreviewKey(renderEnv.revisionPreview);

      // Step 4+: paint + scroll/events with the computed values.
      const applyComputation = (
        computation: LayoutComputation,
        origin: LayoutUpdateOrigin = layoutUpdateOrigin,
        version = sourceVersion
      ) => {
        const { layout: newLayout } = computation;
        stampSourceVersion(newLayout, version);
        stampRevisionPreviewKey(newLayout, previewKey);

        const pagesEl = pagesContainerRef.current;
        const scrollParent =
          getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
        const interactionHost = interactionPageHostRef?.current ?? pagesEl;
        const queries = displayListQueriesRef.current;
        const currentViewportAnchor = currentViewportAnchorRef.current;
        const anchor =
          scrollParent?.isConnected && interactionHost && queries
            ? origin === 'remote'
              ? currentViewportAnchor?.navigationEpoch === navigationEpochRef.current
                ? {
                    kind: 'viewport' as const,
                    anchor: currentViewportAnchor.anchor,
                  }
                : null
              : {
                  kind: 'selection' as const,
                  anchor: captureDisplayListScrollAnchor(
                    queries,
                    interactionHost,
                    scrollParent,
                    getSelectionHead?.() ?? 0
                  ),
                }
            : null;

        viewportAnchorCaptureReadyRef.current = false;
        layoutUpdateOriginRef.current = origin;
        setLayout(newLayout);

        const vp = viewportLayoutRef.current;
        if (vp) {
          const mh = viewportMinHeightPx(newLayout, pageGap);
          vp.style.minHeight = `${mh}px`;
          vp.style.marginBottom = zoom !== 1 ? `${mh * (zoom - 1)}px` : '';
        }
        if (scrollParent?.isConnected && anchor) {
          scrollRestoreController.capture(anchor);
        } else {
          scrollRestoreController.cancel();
        }
      };

      const layOutHere = (): void => {
        try {
          // An edit may have landed since the pass began.
          const version = readSessionVersion(session);
          const computation = computeLayout(computeInputs);
          applyComputation(computation, layoutUpdateOrigin, version);
          const totalTime = performance.now() - pipelineStart;
          if (totalTime > 2000) {
            console.warn(
              `[PagedEditor] Layout pipeline took ${Math.round(totalTime)}ms total ` +
                `(${computation.layout.pages.length} pages)`
            );
          }
        } catch (error) {
          console.error('[PagedEditor] Layout pipeline error:', error);
          onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
        }
      };

      // A resident worker, when one can take the pass, runs the only layout
      // and answers with it and its first frame.
      // The document as opened is laid out by the resident worker alone: that
      // pass also builds its first frame, so the main thread runs no layout
      // before the first paint. Once the document changes, passes run here.
      if (openedVersionRef.current?.session !== session) {
        openedVersionRef.current = { session, version: sourceVersion };
      }
      let workerPass: ReturnType<LayoutInWorker> = null;
      if (
        !onHost &&
        sourceVersion !== null &&
        sourceVersion === openedVersionRef.current.version
      ) {
        try {
          workerPass =
            layoutInWorkerRef.current?.(
              session,
              JSON.stringify({
                ...buildResidentRegionLayoutRequest(document, pageGap, renderEnv),
                measurement,
              })
            ) ?? null;
        } catch (error) {
          console.error('[PagedEditor] Resident worker layout could not start:', error);
        }
      }
      if (!workerPass) {
        layOutHere();
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }
      void workerPass
        .then(
          (computation) => {
            if (pass !== passRef.current || sessionRef.current !== session) return;
            // A change that landed meanwhile makes the worker's layout stale.
            if (!computation || readSessionVersion(session) !== sourceVersion) {
              layOutHere();
              return;
            }
            applyComputation(computation);
            // The first pages paint now; the full layout replaces them.
            void computation.complete?.then(
              (complete) => {
                if (pass !== passRef.current || sessionRef.current !== session) return;
                if (complete && readSessionVersion(session) === sourceVersion) {
                  // Nothing the user did changed: keep their viewport.
                  applyComputation(complete, 'remote');
                } else {
                  layOutHere();
                }
              },
              () => {}
            );
          },
          (error: unknown) => {
            if (pass !== passRef.current) return;
            // The display reports a worker out of memory; nothing lays out here.
            if (error instanceof ResidentWorkerOutOfMemoryError) return;
            console.error('[PagedEditor] Layout pipeline error:', error);
            onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)));
          }
        )
        .finally(() => {
          if (pass === passRef.current) syncCoordinator.onLayoutComplete(currentEpoch);
        });
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
      scrollRestoreController,
    ]
  );

  // Hold the exact scrollTop while the next display-list commit is built.
  useLayoutEffect(() => {
    const ticket = scrollRestoreController.peek();
    if (!ticket) return;
    const pagesEl = pagesContainerRef.current;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (scrollParent?.isConnected) {
      scrollRestoreController.run(ticket, () => {
        restoreScrollSnapshot(ticket.value.anchor, scrollParent);
        expectedScrollTopRef.current = scrollParent.scrollTop;
      });
    }
  }, [layout, getScrollContainer, pagesContainerRef, scrollRestoreController]);

  // A new immutable display-list/query facade is the geometry commit signal.
  useLayoutEffect(() => {
    const ticket = scrollRestoreController.peek();
    const pagesEl = pagesContainerRef.current;
    const host = interactionPageHostRef?.current ?? pagesEl;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!ticket || !displayListQueries || !host || !scrollParent?.isConnected) return;
    const pending = scrollRestoreController.take();
    if (!pending) return;
    const restore = (): void => {
      if (pending.value.kind === 'viewport') {
        restoreDisplayListViewportAnchor(
          pending.value.anchor,
          displayListQueries,
          host,
          scrollParent,
          resolveViewportPosition
        );
      } else {
        restoreDisplayListScrollAnchor(
          pending.value.anchor,
          displayListQueries,
          host,
          scrollParent
        );
      }
      expectedScrollTopRef.current = scrollParent.scrollTop;
    };
    if (!scrollRestoreController.run(pending, restore)) return;
    const rafId = requestAnimationFrame(() => {
      if (scrollParent.isConnected) scrollRestoreController.run(pending, restore);
    });
    return () => cancelAnimationFrame(rafId);
  }, [
    displayListQueries,
    getScrollContainer,
    interactionPageHostRef,
    pagesContainerRef,
    resolveViewportPosition,
    scrollRestoreController,
  ]);

  useLayoutEffect(() => {
    if (!displayListQueries) return;
    viewportAnchorCaptureReadyRef.current = true;
    captureCurrentViewportAnchor();
    const rafId = requestAnimationFrame(() => {
      captureCurrentViewportAnchor();
    });
    return () => cancelAnimationFrame(rafId);
  }, [captureCurrentViewportAnchor, displayListQueries]);

  useEffect(() => {
    const pagesEl = pagesContainerRef.current;
    const scrollParent =
      getScrollContainer() ?? (pagesEl ? findVerticalScrollParentOrRoot(pagesEl) : null);
    if (!scrollParent) return;
    scrollParent.addEventListener('wheel', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    scrollParent.addEventListener('touchstart', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    scrollParent.addEventListener('pointerdown', cancelPendingScrollRestore, {
      capture: true,
      passive: true,
    });
    let captureTimer: ReturnType<typeof setTimeout> | null = null;
    const captureAfterScroll = (): void => {
      const expectedScrollTop = expectedScrollTopRef.current;
      expectedScrollTopRef.current = null;
      if (
        expectedScrollTop == null ||
        Math.abs(scrollParent.scrollTop - expectedScrollTop) > 0.5
      ) {
        cancelPendingScrollRestore();
      }
      if (captureTimer !== null) clearTimeout(captureTimer);
      captureTimer = setTimeout(() => {
        captureTimer = null;
        if (viewportAnchorCaptureReadyRef.current) captureCurrentViewportAnchor();
      }, 80);
    };
    scrollParent.addEventListener('scroll', captureAfterScroll, { passive: true });
    return () => {
      scrollParent.removeEventListener('wheel', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('touchstart', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('pointerdown', cancelPendingScrollRestore, true);
      scrollParent.removeEventListener('scroll', captureAfterScroll);
      if (captureTimer !== null) clearTimeout(captureTimer);
    };
  }, [
    cancelPendingScrollRestore,
    captureCurrentViewportAnchor,
    getScrollContainer,
    pagesContainerRef,
  ]);

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
  const scheduleLayout = useCallback((origin: LayoutUpdateOrigin = 'local') => {
    if (origin === 'local') scrollRestoreController.cancel();
    pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
      pendingLayoutOriginRef.current,
      origin
    );
    if (schedulerRef.current != null) return;
    schedulerRef.current = requestAnimationFrame(() => {
      schedulerRef.current = null;
      if (pendingLayoutOriginRef.current) runRef.current();
    });
  }, [scrollRestoreController]);

  // Clean up pending rAF on unmount. A worker pass answering later must not
  // touch the session, which its owner frees on unmount.
  useEffect(() => {
    return () => {
      passRef.current += 1;
      if (schedulerRef.current != null) cancelAnimationFrame(schedulerRef.current);
    };
  }, []);

  const getLayoutRequest = useCallback((): string | null => {
    if (!session) return null;
    const request = buildResidentRegionLayoutRequest(document, pageGap, renderEnv);
    const requirements = JSON.parse(
      session.layoutFontRequirementsJson(JSON.stringify(request))
    ) as ResidentFontRequirement[];
    const measurement = residentMeasurementConfig(requirements);
    if (!measurement) return null;
    request.measurement = measurement;
    return JSON.stringify(request);
  }, [document, pageGap, renderEnv, residentMeasurementConfig, session]);

  const navigationEpoch = useCallback(() => navigationEpochRef.current, []);

  return {
    layout,
    layoutUpdateOrigin: layoutUpdateOriginRef.current,
    runLayoutPipeline,
    scheduleLayout,
    cancelPendingScrollRestore,
    navigationEpoch,
    getLayoutRequest,
  };
}
