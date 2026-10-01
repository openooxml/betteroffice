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
  proposalRevisionPreview,
  type YrsLoc,
  type YrsRenderEnv,
  type YrsSession,
  type YrsStickyPosition,
} from '@betteroffice/docx/yrs';

import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import { documentPageCount } from './documentPageCount';
import type { FontRequirementsInWorker, LayoutInWorker } from './useDisplayList';
import {
  ensureWorkerOpenReplica,
  workerOpenReplicaPending,
  workerOpenSourceVersion,
} from '../internals/workerOpenReplica';
import {
  registeredWorkerProposalAuthority,
  workerProposalAuthority,
  workerProposalFailure,
} from '../internals/workerProposalAuthority';
import { SupersededPreviewError } from '../internals/supersededPreview';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { viewportMinHeightPx } from '../internals/scrollUtils';
import {
  isSupersededLayout,
  markLayoutQueued,
  markSupersededLayout,
  readSessionVersion,
  revisionPreviewKey,
  revisionPreviewKeyOf,
  sourceVersionOf,
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
  /** `session`: the session whose pass failed. */
  onError?: (error: Error, session: YrsSession) => void;
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
  fontRequirementsInWorker?: FontRequirementsInWorker;
  experimentalWorkerOpen?: boolean;
  onAnchorPositionsChange?: (positions: Map<string, number>) => void;
}

export interface UseLayoutPipelineReturn {
  layout: Layout | null;
  layoutUpdateOrigin: LayoutUpdateOrigin;
  /** `onHost` lays out on this thread even when a worker could. */
  runLayoutPipeline: (options?: { onHost?: boolean }) => void;
  /**
   * `inWorker` lets the pass run in the resident worker unless a change that
   * asked for no such pass lands before it runs; remote updates ask for it.
   */
  scheduleLayout: (origin?: LayoutUpdateOrigin, inWorker?: boolean) => void;
  cancelPendingScrollRestore: () => void;
  /** Counts navigation intents, the user's and programmatic scrolls alike. */
  navigationEpoch: () => number;
  /**
   * The region layout request the pipeline would lay the current document out with now, or
   * `null` while it has no session or the fonts the document needs are not ready.
   */
  getLayoutRequest: () => string | null;
}

/** Whether `next` measures as `last` does and only adds font chains. */
function addsFontChainsOnly(
  last: ResidentMeasurementConfig,
  next: ResidentMeasurementConfig
): boolean {
  if (
    JSON.stringify({ ...last, fontChains: null }) !== JSON.stringify({ ...next, fontChains: null })
  ) {
    return false;
  }
  return Object.entries(last.fontChains).every(([key, chain]) => {
    const kept = next.fontChains[key];
    return kept?.length === chain.length && kept.every((id, index) => id === chain[index]);
  });
}

function workerProposalRenderEnv(session: YrsSession, renderEnv: YrsRenderEnv): YrsRenderEnv {
  return workerProposalAuthority(session)?.initialized
    ? { ...renderEnv, revisionPreview: proposalRevisionPreview(session.getProposals()) }
    : renderEnv;
}

/** A pass may run in the worker only if every change it lays out asked for that. */
function mergeInWorker(current: boolean | null, next: boolean): boolean {
  return (current ?? true) && next;
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
    fontRequirementsInWorker,
    experimentalWorkerOpen = false,
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
  const fontRequirementsInWorkerRef = useRef(fontRequirementsInWorker);
  fontRequirementsInWorkerRef.current = fontRequirementsInWorker;
  const workerOpenEnabledRef = useRef(experimentalWorkerOpen);
  workerOpenEnabledRef.current = experimentalWorkerOpen;
  // Bumped by every pass, so a worker pass answering late never overwrites a
  // newer layout.
  const passRef = useRef(0);
  const sessionRef = useRef(session);
  sessionRef.current = session;
  // The document version the first pass of this session laid out.
  const openedVersionRef = useRef<{ session: YrsSession; version: string | null } | null>(null);
  // The last layout this pipeline applied: its session, document version,
  // request without the revision preview or measurement, preview and measurement.
  const laidOutRef = useRef<{
    session: YrsSession;
    version: string | null;
    request: string;
    previewKey: string;
    measurement: ResidentMeasurementConfig;
  } | null>(null);
  const workerPrewarmRef = useRef<{ session: YrsSession; release: () => void } | null>(null);
  const releaseWorkerPrewarm = useCallback((owner: YrsSession | null) => {
    const worker = workerPrewarmRef.current;
    if (!worker || worker.session !== owner) return;
    worker.release();
    workerPrewarmRef.current = null;
  }, []);
  useEffect(() => () => releaseWorkerPrewarm(session), [releaseWorkerPrewarm, session]);
  // A deferred pass that had to run on this thread keeps that requirement.
  const pendingOnHostRef = useRef(false);
  // Whether every change the next pass lays out asked for a worker pass.
  const pendingInWorkerRef = useRef<boolean | null>(null);
  // The worker pass in flight. A pass that may run in the worker waits for it,
  // so a burst of updates lays out their latest state once, not each in turn.
  const workerPassRef = useRef<{ pass: number; session: YrsSession; opening: boolean } | null>(
    null
  );
  const queuedBehindWorkerRef = useRef(false);
  const schedulerRef = useRef<number | null>(null);
  const runRef = useRef<() => void>(() => {});
  const unmountedRef = useRef(false);
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
    if (
      layout &&
      session &&
      workerProposalAuthority(session)?.initialized &&
      !isSupersededLayout(layout) &&
      sourceVersionOf(layout) === session.version() &&
      revisionPreviewKeyOf(layout) ===
        revisionPreviewKey(proposalRevisionPreview(session.getProposals()))
    ) markLayoutQueued(session, false);
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
  const requestPass = useCallback(() => {
    if (schedulerRef.current != null || unmountedRef.current) return;
    schedulerRef.current = requestAnimationFrame(() => {
      schedulerRef.current = null;
      if (pendingLayoutOriginRef.current) runRef.current();
    });
  }, []);

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

  const workerHeld = useCallback(
    (owner: YrsSession): boolean =>
      registeredWorkerProposalAuthority(owner)?.holdsWorkerState() === true ||
      (workerOpenEnabledRef.current && workerOpenReplicaPending(owner)),
    []
  );
  const queueWorkerPass = useCallback((owner: YrsSession): void => {
    queuedBehindWorkerRef.current = true;
    markLayoutQueued(owner, true);
    pendingInWorkerRef.current = mergeInWorker(pendingInWorkerRef.current, true);
    pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
      pendingLayoutOriginRef.current,
      'remote'
    );
  }, []);

  const runLayoutPipeline = useCallback(
    (options?: { onHost?: boolean }) => {
      const workerRequired = session !== null &&
        registeredWorkerProposalAuthority(session)?.holdsWorkerState() === true;
      const onHost = !workerRequired && (options?.onHost === true || pendingOnHostRef.current);
      const inWorker = workerRequired || (!onHost && pendingInWorkerRef.current === true);
      const inFlight = workerPassRef.current;
      // A host batch waits for the worker pass in flight, and so does a pass no
      // change asked to run here, such as a preview change, unless the pass in
      // flight lays out the document as opened.
      const waits =
        inWorker || (!onHost && pendingInWorkerRef.current === null && !inFlight?.opening);
      if (waits && session && inFlight?.session === session) {
        queuedBehindWorkerRef.current = true;
        markLayoutQueued(session, true);
        pendingLayoutOriginRef.current ??= 'local';
        return;
      }
      queuedBehindWorkerRef.current = false;
      pendingInWorkerRef.current = null;
      const pass = ++passRef.current;
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
        pendingInWorkerRef.current = mergeInWorker(pendingInWorkerRef.current, inWorker);
        syncCoordinator.onLayoutComplete(currentEpoch);
        return;
      }

      const passRenderEnv = workerProposalRenderEnv(session, renderEnv);
      const run = (workerRequirements?: string | null): void => {
        let measurement: ResidentMeasurementConfig | null = null;
        try {
          const request = buildResidentRegionLayoutRequest(document, pageGap, passRenderEnv);
          if (workerOpenEnabledRef.current) request.cachedPageTotals = true;
          const input = JSON.stringify(request);
          const pendingRequirements =
            (workerOpenEnabledRef.current ||
              registeredWorkerProposalAuthority(session)?.holdsWorkerState()) &&
              workerRequirements === undefined
              ? fontRequirementsInWorkerRef.current?.(session, input)
              : null;
          if (pendingRequirements) {
            void pendingRequirements.then(
              (requirements) => {
                if (pass === passRef.current && sessionRef.current === session) run(requirements);
              },
              (error: unknown) => {
                if (pass !== passRef.current || sessionRef.current !== session) return;
                markLayoutQueued(session, false);
                // A superseded preflight drops its pass; the pass that superseded it lays out.
                if (!(error instanceof SupersededPreviewError)) {
                  onErrorRef.current?.(
                    error instanceof Error ? error : new Error(String(error)),
                    session
                  );
                }
                syncCoordinator.onLayoutComplete(currentEpoch);
              }
            );
            return;
          }
          const requirements = JSON.parse(
            workerRequirements ?? session.layoutFontRequirementsJson(input)
          ) as ResidentFontRequirement[];
          measurement = residentMeasurementConfig(requirements);
        } catch (error) {
          console.error('[PagedEditor] Resident font preflight error:', error);
          markLayoutQueued(session, false);
          releaseWorkerPrewarm(session);
          onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)), session);
          syncCoordinator.onLayoutComplete(currentEpoch);
          return;
        }
        if (!measurement) {
          if (!onHost && workerPrewarmRef.current?.session !== session) {
            const version = readSessionVersion(session);
            if (
              version !== null &&
              (openedVersionRef.current?.session !== session ||
                version ===
                  (workerOpenEnabledRef.current
                    ? workerOpenSourceVersion(session, openedVersionRef.current.version)
                    : openedVersionRef.current.version))
            ) {
              try {
                const release = layoutInWorkerRef.current?.prewarm?.(session);
                if (release) {
                  workerPrewarmRef.current?.release();
                  workerPrewarmRef.current = { session, release };
                }
              } catch {}
            }
          }
          pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
            pendingLayoutOriginRef.current,
            layoutUpdateOrigin
          );
          pendingOnHostRef.current = onHost;
          pendingInWorkerRef.current = mergeInWorker(pendingInWorkerRef.current, inWorker);
          syncCoordinator.onLayoutComplete(currentEpoch);
          return;
        }
        pendingOnHostRef.current = false;
        // A queued pass deferred above still holds settles until it gets this far.
        if (!workerProposalAuthority(session)?.initialized) markLayoutQueued(session, false);

        const computeInputs = {
          document,
          pageGap,
          session,
          renderEnv: passRenderEnv,
          measurement,
          ...(workerOpenEnabledRef.current ? { cachedPageTotals: true } : {}),
        };
        const sourceVersion = readSessionVersion(session);
        const previewKey = revisionPreviewKey(passRenderEnv.revisionPreview);
        const request = {
          ...buildResidentRegionLayoutRequest(document, pageGap, passRenderEnv),
          measurement,
          ...(workerOpenEnabledRef.current ? { cachedPageTotals: true } : {}),
        };
        const requestWithoutPreview = JSON.stringify(
          { ...request, measurement: undefined },
          (key, value: unknown) => (key === 'revisionPreview' ? undefined : value)
        );

        // Step 4+: paint + scroll/events with the computed values.
        const applyComputation = (
          computation: LayoutComputation,
          origin: LayoutUpdateOrigin = layoutUpdateOrigin,
          version = sourceVersion
        ) => {
          const { layout: newLayout } = computation;
          stampSourceVersion(
            newLayout,
            workerOpenEnabledRef.current ? workerOpenSourceVersion(session, version) : version
          );
          stampRevisionPreviewKey(newLayout, previewKey);
          laidOutRef.current = {
            session,
            version,
            request: requestWithoutPreview,
            previewKey,
            measurement: computeInputs.measurement,
          };

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
          if (registeredWorkerProposalAuthority(session)?.holdsWorkerState()) {
            queueWorkerPass(session);
            if (!workerPassRef.current) requestPass();
            return;
          }
          try {
            // An edit may have landed since the pass began.
            if (workerOpenEnabledRef.current) ensureWorkerOpenReplica(session);
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
            onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)), session);
          }
        };

        // A resident worker, when one can take the pass, runs the only layout
        // and answers with it and its first frame.
        // The document as opened is laid out by the resident worker alone: that
        // pass also builds its first frame, so the main thread runs no layout
        // before the first paint. So is a pass that changes only the revision
        // preview of the layout last applied (adding at most the font chains
        // that preview needs), and a pass for host batches or
        // remote updates alone, which no caret waits on. Passes for local edits
        // run here.
        if (openedVersionRef.current?.session !== session) {
          openedVersionRef.current = { session, version: sourceVersion };
        }
        const laidOut = laidOutRef.current;
        const previewOnly =
          laidOut?.session === session &&
          laidOut.version === sourceVersion &&
          laidOut.request === requestWithoutPreview &&
          (JSON.stringify(laidOut.measurement) === JSON.stringify(measurement) ||
            (laidOut.previewKey !== previewKey &&
              addsFontChainsOnly(laidOut.measurement, computeInputs.measurement)));
        let workerPass: ReturnType<LayoutInWorker> = null;
        if (
          registeredWorkerProposalAuthority(session)?.holdsWorkerState() ||
          (!onHost &&
            sourceVersion !== null &&
            (previewOnly ||
              inWorker ||
              sourceVersion ===
                (workerOpenEnabledRef.current
                  ? workerOpenSourceVersion(session, openedVersionRef.current.version)
                  : openedVersionRef.current.version)))
        ) {
          try {
            workerPass = layoutInWorkerRef.current?.(session, JSON.stringify(request)) ?? null;
          } catch (error) {
            if (workerProposalFailure(session) === error) {
              syncCoordinator.onLayoutComplete(currentEpoch);
              return;
            }
            console.error('[PagedEditor] Resident worker layout could not start:', error);
            if (registeredWorkerProposalAuthority(session)?.holdsWorkerState()) {
              onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)), session);
              syncCoordinator.onLayoutComplete(currentEpoch);
              return;
            }
          }
        }
        // The spare warmed while fonts loaded has been adopted by now, or is not needed.
        releaseWorkerPrewarm(session);
        if (!workerPass) {
          layOutHere();
          syncCoordinator.onLayoutComplete(currentEpoch);
          return;
        }
        workerPassRef.current = { pass, session, opening: !inWorker && !previewOnly };
        void workerPass
          .then(
            (computation) => {
              if (pass !== passRef.current || sessionRef.current !== session) return;
              // A change that landed meanwhile makes the worker's layout stale.
              // Only a queued worker pass follows it at once; until then it is
              // the newest layout there is, so it paints but settles no wait.
              // A queued pass supersedes it even at the same version, such as a
              // revision preview change, and lays out in its place when it failed.
              const stale =
                readSessionVersion(session) !==
                (workerOpenEnabledRef.current
                  ? workerOpenSourceVersion(session, sourceVersion)
                  : sourceVersion);
              // A document only the worker holds lays out its newer state there, not here.
              if (stale && !queuedBehindWorkerRef.current && workerHeld(session)) {
                queueWorkerPass(session);
              }
              const queued = queuedBehindWorkerRef.current;
              if (!computation || (stale && !queued)) {
                if (!queued) layOutHere();
                return;
              }
              if (stale || queued) markSupersededLayout(computation.layout);
              applyComputation(computation);
              // The first pages paint now; the full layout replaces them.
              void computation.complete?.then(
                (complete) => {
                  if (pass !== passRef.current || sessionRef.current !== session) return;
                  if (
                    complete &&
                    readSessionVersion(session) ===
                      (workerOpenEnabledRef.current
                        ? workerOpenSourceVersion(session, sourceVersion)
                        : sourceVersion)
                  ) {
                    if (queuedBehindWorkerRef.current) markSupersededLayout(complete.layout);
                    // Nothing the user did changed: keep their viewport.
                    applyComputation(complete, 'remote');
                  } else if (queuedBehindWorkerRef.current) {
                    return;
                  } else if (
                    (complete && workerHeld(session)) ||
                    (!complete && registeredWorkerProposalAuthority(session)?.holdsWorkerState() === true)
                  ) {
                    queueWorkerPass(session);
                    requestPass();
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
              if (workerProposalFailure(session) === error) return;
              if (error instanceof SupersededPreviewError) return;
              console.error('[PagedEditor] Layout pipeline error:', error);
              onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)), session);
            }
          )
          .finally(() => {
            if (pass === passRef.current) syncCoordinator.onLayoutComplete(currentEpoch);
            if (workerPassRef.current?.pass !== pass) return;
            workerPassRef.current = null;
            if (queuedBehindWorkerRef.current && sessionRef.current === session) {
              requestPass();
            }
          });
      };
      run();
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
      requestPass,
      releaseWorkerPrewarm,
      workerHeld,
      queueWorkerPass,
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
  runRef.current = runLayoutPipeline;
  const scheduleLayout = useCallback(
    (origin: LayoutUpdateOrigin = 'local', inWorker = origin === 'remote') => {
      if (origin === 'local') scrollRestoreController.cancel();
      pendingLayoutOriginRef.current = mergeLayoutUpdateOrigin(
        pendingLayoutOriginRef.current,
        origin
      );
      pendingInWorkerRef.current = mergeInWorker(pendingInWorkerRef.current, inWorker);
      requestPass();
    },
    [requestPass, scrollRestoreController]
  );

  // Clean up pending rAF on unmount. A worker pass answering later must not
  // touch the session, which its owner frees on unmount.
  useEffect(() => {
    unmountedRef.current = false;
    return () => {
      unmountedRef.current = true;
      if (sessionRef.current) markLayoutQueued(sessionRef.current, false);
      passRef.current += 1;
      if (schedulerRef.current != null) cancelAnimationFrame(schedulerRef.current);
      schedulerRef.current = null;
    };
  }, []);

  const getLayoutRequest = useCallback((): string | null => {
    if (!session) return null;
    const request = buildResidentRegionLayoutRequest(
      document,
      pageGap,
      workerProposalRenderEnv(session, renderEnv)
    );
    if (workerOpenEnabledRef.current) request.cachedPageTotals = true;
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
