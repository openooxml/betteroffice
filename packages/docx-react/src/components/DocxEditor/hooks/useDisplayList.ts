import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildRustDisplayList,
  buildRustDisplayFrame,
  applyFrameDelta,
  applyFrameDeltaOwned,
  createCanvasImageResolver,
  createDisplayListQueries,
  decodeFrameDelta,
  demoDisplayList,
  encodeDisplayListFrameExtras,
  isDisplayListQuerySourceDead,
  type DisplayList,
  type DisplayListQueries,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RustDisplayListEngine,
  type RetainedFrame,
  type ResidentDisplayListQueryEngine,
} from '@betteroffice/docx/layout/render';
import {
  getLayoutKernelInputs,
  workerLayoutComputation,
  type LayoutComputation,
} from '@betteroffice/docx/editor';
import {
  canUseResidentEngineWorker,
  residentCaretSnapshotForFrame,
  ResidentEngineWorkerClient,
  ResidentWorkerFailureError,
  takePreloadedResidentEngineWorker,
  ResidentWorkerOutOfMemoryError,
  sameYrsSelection,
  type ResidentCaretPaintStyle,
  type ResidentEngineOffscreenPage,
  type ResidentEngineWorkerFrame,
  type YrsResidentCaretSnapshot,
  type YrsRenderEnv,
  type YrsResidentWorkerSnapshot,
  type YrsSelection,
  type WasmModuleMemory,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { RustFontChainsProvider } from './useRustMeasurement';
import { displayListNeedsHostImages } from '../canvasPresentation';
import { CARET_PAINT_IDLE_MS, PaintedCaretMachine } from '../paintedCaret';
import {
  readSessionVersion,
  revisionPreviewKey,
  revisionPreviewKeyOf,
  sourceVersionOf,
  stampRevisionPreviewKey,
  stampSourceVersion,
  UNKNOWN_REVISION_PREVIEW_KEY,
} from '../internals/layoutProvenance';
import {
  DisplayListQueryEpochGate,
  type ResolveDisplayListQueries,
} from './displayListQueryEpochGate';

// provider for the canvas renderer's display list: returns the injected value
// when the host supplies one, otherwise the demo fixture. consumers only ever
// see a DisplayList, never where it came from.
export function useDisplayList(injected?: DisplayList | null): DisplayList {
  return injected ?? demoDisplayList;
}

export interface UseRustDisplayListResult {
  /** the latest successfully built display list (kept across rebuilds so the canvas never blanks mid-compute) */
  displayList: DisplayList | null;
  /** fatal error from resolving inputs or building the display list */
  error: Error | null;
  /** The engine whose layout, build or input failed with `error`. */
  errorEngine: unknown;
  /** true until the first display list for the current document is ready */
  loading: boolean;
  /** Binary retained-frame state; null on the compatibility JSON path. */
  frame: RetainedFrame | null;
  /** Query facade built from the same display list as `frame`. */
  queries: DisplayListQueries | null;
  /** Resolve the newest query facade after pending document/frame changes. */
  resolveQueries: ResolveDisplayListQueries;
  /**
   * The display list once it shows every document change so far, laid out in
   * full. `relayout` runs a layout pass when none is on its way; rejects when
   * rendering fails, or after `timeoutMs` (15 s by default, none when null).
   */
  settledDisplayList(
    relayout: (() => void) | null,
    timeoutMs?: number | null
  ): Promise<DisplayList>;
  /**
   * Drops the settled display list. Without `failure` a new document is on its way, so waiters
   * wait for its display list; with one they reject.
   */
  resetSettled(failure?: Error | null): void;
  /** True from a document load until the loaded document's first layout arrives. */
  awaitingDocument(): boolean;
  /** Ties a session, as it is created, to the document load now under way. */
  recordSession(session: YrsSession | null): void;
  /** Worker-computed caret tagged to `frame`. */
  caret: YrsResidentCaretSnapshot | null;
  /** Apply a plain-text edit through the resident engine and publish its frame. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** Apply up to `count` collapsed deletions/paragraph merges through the resident engine. */
  applyDelete(
    direction: 'backward' | 'forward',
    count?: number
  ): Promise<ResidentFrameApplyResult | null>;
  /** See {@link LayoutInWorker}. */
  layoutInWorker: LayoutInWorker;
  /** The engine whose layout the latest published frame shows. */
  presentedEngine: unknown;
  /** {@link presentedEngine} as of now, before the render that shows the frame. */
  shownFrameEngine(): unknown;
  /** Lets go of every engine the pages showed: the resident worker and the presented engine. */
  release(): void;
  /**
   * The pages `[start, end)` near the viewport. Only these are built; every
   * other page arrives as geometry until it comes near.
   */
  setDisplayWindow(start: number, end: number): void;
  /** The resident worker's wasm memories as of its latest reply; null without a worker. */
  workerMemory(): WasmModuleMemory[] | null;
  /**
   * True while the worker owns the visible page surfaces. Sticky across
   * invalidation (remote/structural updates) so the canvas keeps its last
   * pixels instead of remounting; drops only on genuine fallback or reset.
   */
  workerSurfacesActive: boolean;
  workerPresentationActive: boolean;
  setWorkerPresentationActive(active: boolean): void;
  attachOffscreenCanvases(
    pages: ResidentEngineOffscreenPage[],
    activePageIds: string[],
    devicePixelRatio: number,
    zoom: number,
    caretStyle: ResidentCaretPaintStyle
  ): Promise<boolean>;
  /** True while the worker-painted caret line owns the caret (DOM caret hidden). */
  paintedCaretActive: boolean;
  /** Local text input: keeps painted-caret mode alive for follow-up frames. */
  notifyCaretInput(): void;
  /** Text input dispatched: hide the DOM caret before the worker round-trip. */
  notifyCaretInputDispatched(): void;
  /** Selection move / blur / IME start / mode change: immediate swap to the DOM caret. */
  notifyCaretInterrupt(): void;
}

/**
 * Lay the document out in the resident worker, which then owns that layout:
 * the reply carries the layout and its first frame, and the main thread runs
 * no layout of its own. Null when no worker can take it, and a null result
 * when the worker failed; either way the caller lays out on the main thread.
 */
export type LayoutInWorker = (
  session: YrsSession,
  request: string
) => Promise<WorkerLayoutComputation | null> | null;

/**
 * A worker layout. On open it may cover only the first pages: `complete`
 * then brings the full layout, or null when the worker could not finish it.
 */
export interface WorkerLayoutComputation extends LayoutComputation {
  complete?: Promise<LayoutComputation | null>;
}

/** Pages the first worker layout covers before the rest of the body. */
const PROVISIONAL_LAYOUT_PAGES = 3;
/** How long the rest of the layout waits for the first surfaces to attach. */
const PROVISIONAL_SURFACE_WAIT_MS = 250;
// Body blocks the first step of a provisional layout's completion measures;
// later steps are sized to a time slice. 0 completes it in one step.
const COMPLETION_SLICE_BLOCKS = 64;

interface WorkerLayoutFrame {
  result: ResidentEngineWorkerFrame;
  previousFrame: RetainedFrame | null;
  engine: YrsSession;
  /** Content epoch and display extras the worker built the frame for. */
  contentEpoch: number;
  layoutExtras: string;
  /** The frame shows a layout of the first pages only. */
  provisional: boolean;
}

/** The display fallback needs a main-thread layout of a worker-run one. */
class MainThreadLayoutPendingError extends Error {}

/**
 * A newer layout reached the session, with another revision preview or from a
 * worker pass run again; its own pass shows it.
 */
class SupersededPreviewError extends Error {}

export interface ResidentFrameApplyResult {
  frameEpoch: number | null;
  caretSynchronized: boolean;
  /** Characters a resident deletion removed; absent when unknown. */
  deletedUnits?: number;
}

/** Pages a worker's first frame builds before the viewport is known. */
const INITIAL_DISPLAY_WINDOW: [number, number] = [0, 5];
/** Unbuilt pages built per request while the complete list is awaited. */
const SETTLE_BUILD_BATCH_PAGES = 32;
/** Unbuilt pages built per idle period away from the viewport. */
const BACKGROUND_BUILD_BATCH_PAGES = 16;
const BACKGROUND_BUILD_DELAY_MS = 200;

type PageBuildTimer = ReturnType<typeof setTimeout> | { idle: number };

function cancelPageBuilds(timer: { current: PageBuildTimer | null }): void {
  const scheduled = timer.current;
  timer.current = null;
  if (scheduled === null) return;
  if (typeof scheduled === 'object' && 'idle' in scheduled) cancelIdleCallback(scheduled.idle);
  else clearTimeout(scheduled);
}

/** test seam: unit tests inject a fake engine/inputs-resolver instead of the wasm module */
export interface RustDisplayListHookOverrides {
  build?: typeof buildRustDisplayList;
  getInputs?: typeof getLayoutKernelInputs;
}

type ResidentInputOperation =
  | { kind: 'insert'; text: string }
  | { kind: 'delete'; direction: 'backward' | 'forward'; count: number };

interface RustDisplayListSnapshot {
  displayList: DisplayList | null;
  frame: RetainedFrame | null;
  queries: DisplayListQueries | null;
  caret: YrsResidentCaretSnapshot | null;
}

const EMPTY_DISPLAY_LIST_SNAPSHOT: RustDisplayListSnapshot = {
  displayList: null,
  frame: null,
  queries: null,
  caret: null,
};

interface BuiltDisplay {
  displayList: DisplayList;
  frame: RetainedFrame | null;
  caret: YrsResidentCaretSnapshot | null;
  queryEngine: RustDisplayListEngine | null | undefined;
  workerProduced: boolean;
  caretPainted: boolean;
  /** Shows a layout of the first pages only, so it does not settle. */
  provisional?: boolean;
  /** The preview a worker frame was built with; absent for a frame of `layout` itself. */
  previewKey?: string | null;
}

// A replacement worker's frames follow the frame on screen.
function followedFrameEpoch(frame: RetainedFrame | null): { frameEpoch?: number } {
  return frame?.frameEpoch ? { frameEpoch: frame.frameEpoch } : {};
}

// rebuilds the display list through the rust wasm engine after every layout
// pass. dumb replay glue: the `{ measured, options, layout }` triple (plus the
// kernel-recorded `headersFooters` payload when the document has HF parts) is
// serialized as-is (the same envelope the golden fixtures pin) and every
// geometry/paint decision happens in rust. Every adapter interaction is
// display-list-backed. A generation counter drops stale async results so only
// the newest layout wins.
export function useRustDisplayList(
  layout: Layout | null,
  overrides?: RustDisplayListHookOverrides,
  // Host slot the Rust measure source fills with the merged doc-wide font
  // chains. Read at build time and passed through to the builder — its
  // presence (a non-empty map) is what activates GlyphRun emission. Null /
  // undefined ⇒ the builder emits browser-shaped TextRunPrimitives (unchanged).
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>,
  // Resolved comment threads whose range wash the canvas must hide (the crate
  // drops their comment-range decorations and stamps status="resolved").
  // Changing the set rebuilds the display list so resolve/reopen — and the
  // "expanded resolved card re-tints its range" flow — repaint immediately.
  resolvedCommentIds?: ReadonlySet<number>,
  engine?: RustDisplayListEngine | null,
  /** Asks the host for a layout of the document as it is now. */
  requestLayout?: () => void,
  /** The most a resident worker's editing core may allocate at once. */
  workerHeapLimitBytes?: number,
  /**
   * The session whose pages are shown while another session replaces it
   * page for page: its worker and page surfaces carry over to the next.
   */
  handoffFromRef?: React.RefObject<YrsSession | null>
): UseRustDisplayListResult {
  const requestLayoutRef = useRef(requestLayout);
  requestLayoutRef.current = requestLayout;
  const workerHeapLimitRef = useRef(workerHeapLimitBytes);
  workerHeapLimitRef.current = workerHeapLimitBytes;
  // The engine whose layout the shown frame is of: another engine's frames
  // never apply to it as a base.
  const frameEngineRef = useRef<unknown>(null);
  const [presentedEngine, setPresentedEngine] = useState<unknown>(null);
  const [snapshot, setSnapshot] = useState<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const snapshotRef = useRef<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const queryEpochGateRef = useRef<DisplayListQueryEpochGate | null>(null);
  if (!queryEpochGateRef.current) queryEpochGateRef.current = new DisplayListQueryEpochGate();
  const queryEpochGate = queryEpochGateRef.current;
  const contentEpochRef = useRef(0);
  const layoutPreviewKeyRef = useRef<string | null>(null);
  const workerPreviewKeysRef = useRef(new Map<number, string>());
  const settledEpochRef = useRef<number | null>(null);
  const settleErrorRef = useRef<Error | null>(null);
  const settleWaitersRef = useRef(new Set<() => void>());
  const settleRelayoutRef = useRef<(() => void) | null>(null);
  const layoutRef = useRef(layout);
  layoutRef.current = layout;
  // The layout shown when another document started loading. What becomes of
  // it settles no wait: waits are for the document being loaded.
  const replacedLayoutRef = useRef<{ layout: Layout | null } | null>(null);
  const documentLoadsRef = useRef(0);
  const markSettled = useCallback(
    (epoch: number | null, failure: Error | null = null, authoritative = false): void => {
      if (replacedLayoutRef.current && !authoritative) return;
      settledEpochRef.current = epoch;
      settleErrorRef.current = failure;
      for (const waiter of [...settleWaitersRef.current]) waiter();
    },
    []
  );
  const requestSettleRelayout = useCallback((): void => {
    if (settleWaitersRef.current.size === 0) return;
    setTimeout(() => {
      if (settleWaitersRef.current.size > 0) settleRelayoutRef.current?.();
    }, 0);
  }, []);
  const [error, setError] = useState<Error | null>(null);
  const [errorEngine, setErrorEngine] = useState<unknown>(null);
  const [loading, setLoading] = useState(true);
  const generationRef = useRef(0);
  const workerRef = useRef<{
    engine: YrsSession;
    client: ResidentEngineWorkerClient;
    /** The document load its session belongs to. */
    load: number;
  } | null>(null);
  // The document load each session belongs to: the one under way when it was
  // created, as the editor records it, else when it was first laid out or shown.
  const sessionLoadsRef = useRef(new WeakMap<YrsSession, number>());
  const sessionLoad = useCallback((session: YrsSession): number => {
    let load = sessionLoadsRef.current.get(session);
    if (load === undefined) {
      load = documentLoadsRef.current;
      sessionLoadsRef.current.set(session, load);
    }
    return load;
  }, []);
  const spawnedWorkerEnginesRef = useRef(new WeakSet<YrsSession>());
  // Sessions whose worker and page surfaces a successor took over: they never take one back.
  const handedOverEnginesRef = useRef(new WeakSet<YrsSession>());
  // The worker for `hostEngine`: its own, the one of the session it
  // takes over from, a spare, or a new one.
  const workerFor = useCallback(
    (hostEngine: YrsSession): NonNullable<typeof workerRef.current> => {
      const current = workerRef.current;
      if (current?.engine === hostEngine) return current;
      const load = sessionLoad(hostEngine);
      const replacement = spawnedWorkerEnginesRef.current.has(hostEngine);
      spawnedWorkerEnginesRef.current.add(hostEngine);
      if (current && handoffFromRef?.current === current.engine) {
        handedOverEnginesRef.current.add(current.engine);
        current.client.rebootstrap();
        workerRef.current = { engine: hostEngine, client: current.client, load };
        return workerRef.current;
      }
      current?.client.destroy();
      // A successor that fails to construct leaves no destroyed client current.
      workerRef.current = null;
      const spare = replacement ? null : takePreloadedResidentEngineWorker();
      workerRef.current = {
        engine: hostEngine,
        client: spare ?? new ResidentEngineWorkerClient(),
        load,
      };
      return workerRef.current;
    },
    [handoffFromRef, sessionLoad]
  );
  const workerFallbackEngineRef = useRef<YrsSession | null>(null);
  // No worker starts once the hook is gone, whatever failure arrives late.
  const unmountedRef = useRef(false);
  // The engines whose worker ran out of memory, each with the failure once its
  // replacement did too. Weak, so a replaced document's session is not kept.
  const outOfMemoryRef = useRef(new WeakMap<YrsSession, ResidentWorkerOutOfMemoryError | null>());
  const displayWindowRef = useRef<[number, number]>(INITIAL_DISPLAY_WINDOW);
  const pageBuildInFlightRef = useRef(false);
  const pageBuildTimerRef = useRef<PageBuildTimer | null>(null);
  const schedulePageBuildsWhenIdleRef = useRef<() => void>(() => {});
  const workerLayoutFramesRef = useRef(new WeakMap<Layout, WorkerLayoutFrame>());
  const completionGateRef = useRef<(() => void) | null>(null);
  const resolvedCommentIdsRef = useRef(resolvedCommentIds);
  resolvedCommentIdsRef.current = resolvedCommentIds;
  const recoveryFrameEpochRef = useRef(0);
  const workerInputQueueRef = useRef<Promise<void>>(Promise.resolve());
  const suppressWorkerInvalidationRef = useRef(0);
  const [workerSurfacesActive, setWorkerSurfacesActive] = useState(false);
  const workerPresentationActiveRef = useRef(false);
  const [workerPresentationActive, setWorkerPresentationActiveState] = useState(false);

  const paintedCaretMachineRef = useRef<PaintedCaretMachine | null>(null);
  if (!paintedCaretMachineRef.current) paintedCaretMachineRef.current = new PaintedCaretMachine();
  const paintedCaretMachine = paintedCaretMachineRef.current;
  const [paintedCaretActive, setPaintedCaretActiveState] = useState(false);
  const paintedCaretIdleTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const dispatchHoldTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const residentPaintInflightRef = useRef(0);

  // Idle/interrupt swap: the DOM caret renders first (state flip), then the
  // erase posts two rAFs later — an overlap is possible, a caret gap is not.
  // The post re-checks the machine so a paint that landed meanwhile survives.
  const requestPaintedCaretErase = useCallback((): void => {
    setPaintedCaretActiveState(false);
    const post = (): void => {
      if (paintedCaretMachine.isActive() || residentPaintInflightRef.current > 0) return;
      workerRef.current?.client.eraseCaret();
    };
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => requestAnimationFrame(post));
    } else {
      setTimeout(post, 0);
    }
  }, [paintedCaretMachine]);

  const schedulePaintedCaretIdle = useCallback((): void => {
    if (paintedCaretIdleTimerRef.current !== null) clearTimeout(paintedCaretIdleTimerRef.current);
    const fire = (): void => {
      paintedCaretIdleTimerRef.current = null;
      const now = performance.now();
      if (paintedCaretMachine.idleTimeout(now)) {
        requestPaintedCaretErase();
        return;
      }
      if (paintedCaretMachine.isActive()) {
        paintedCaretIdleTimerRef.current = setTimeout(
          fire,
          Math.max(1, paintedCaretMachine.msUntilIdle(now))
        );
      }
    };
    paintedCaretIdleTimerRef.current = setTimeout(fire, CARET_PAINT_IDLE_MS);
  }, [paintedCaretMachine, requestPaintedCaretErase]);

  const applyPaintedCaretReply = useCallback(
    (painted: boolean, token: number): void => {
      if (painted && paintedCaretMachine.framePainted(token)) {
        setPaintedCaretActiveState(true);
        schedulePaintedCaretIdle();
        return;
      }
      // Painted but interrupted mid-flight: the DOM caret is already showing,
      // so the stale line can be erased immediately.
      if (painted) workerRef.current?.client.eraseCaret();
      // An in-flight resident input owns the next verdict; deciding here would
      // remount the DOM caret for a frame in the middle of a burst.
      if (residentPaintInflightRef.current > 0) return;
      paintedCaretMachine.frameUnpainted();
      setPaintedCaretActiveState(false);
    },
    [paintedCaretMachine, schedulePaintedCaretIdle]
  );

  const notifyCaretInput = useCallback((): void => {
    paintedCaretMachine.noteInput(performance.now());
  }, [paintedCaretMachine]);

  // Dispatch-time hide: called synchronously from the input event, BEFORE the
  // worker round-trip. The worker presents glyphs+painted caret atomically off
  // the main thread, so a caret left mounted until the reply commits shows the
  // new character with a stale caret one position behind. Hiding for the
  // in-flight window is invisible; the dispatch hold self-expires and every
  // resolution path (painted, unpainted, interrupt) reconciles it.
  const notifyCaretInputDispatched = useCallback((): void => {
    if (!workerPresentationActiveRef.current) return;
    paintedCaretMachine.noteDispatch(performance.now());
    setPaintedCaretActiveState(true);
    if (dispatchHoldTimerRef.current !== null) clearTimeout(dispatchHoldTimerRef.current);
    dispatchHoldTimerRef.current = setTimeout(() => {
      dispatchHoldTimerRef.current = null;
      const now = performance.now();
      if (!paintedCaretMachine.isActive() && !paintedCaretMachine.isHolding(now)) {
        setPaintedCaretActiveState(false);
      }
    }, CARET_PAINT_IDLE_MS + 16);
  }, [paintedCaretMachine]);

  const notifyCaretInterrupt = useCallback((): void => {
    if (paintedCaretMachine.interrupt()) {
      requestPaintedCaretErase();
      return;
    }
    // A pending dispatch hold must not survive a selection change.
    setPaintedCaretActiveState(false);
  }, [paintedCaretMachine, requestPaintedCaretErase]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    window.addEventListener('blur', notifyCaretInterrupt);
    return () => window.removeEventListener('blur', notifyCaretInterrupt);
  }, [notifyCaretInterrupt]);

  const setWorkerPresentationActive = useCallback(
    (active: boolean): void => {
      workerPresentationActiveRef.current = active;
      if (!active) notifyCaretInterrupt();
      setWorkerPresentationActiveState((current) => (current === active ? current : active));
    },
    [notifyCaretInterrupt]
  );

  const residentEngine = isWorkerHostEngine(engine) ? engine : null;
  const residentEngineRef = useRef(residentEngine);
  residentEngineRef.current = residentEngine;

  // Rendering moves to the host engine for good: its frames start from a fresh
  // base, numbered after the worker's last frame.
  const adoptHostEngine = useCallback(
    (hostEngine: YrsSession): void => {
      if (workerFallbackEngineRef.current === hostEngine) return;
      hostEngine.resetFrameBase();
      queryEpochGate.clear();
      recoveryFrameEpochRef.current = snapshotRef.current.frame?.frameEpoch ?? 0;
      const fallbackSnapshot = { ...snapshotRef.current, frame: null, queries: null, caret: null };
      snapshotRef.current = fallbackSnapshot;
      setSnapshot(fallbackSnapshot);
      workerFallbackEngineRef.current = hostEngine;
    },
    [queryEpochGate]
  );

  useEffect(() => {
    if (!residentEngine) return;
    return residentEngine.onUpdate((update) => {
      if (suppressWorkerInvalidationRef.current > 0) return;
      contentEpochRef.current += 1;
      queryEpochGate.invalidate();
      requestSettleRelayout();
      // Update observers fire from inside the wasm transaction. Calling any
      // other EditSession method here would re-enter the borrowed wasm object
      // (wasm-bindgen correctly rejects that unsafe alias). Selection is sent
      // with the next input request after the transaction has returned.
      if (canUseResidentEngineWorker()) workerRef.current?.client.invalidate(update, null);
      // Only frame freshness drops here. The worker keeps the page surfaces
      // (workerSurfacesActive) so the canvas retains its pixels until the
      // post-sync frame lands — flipping surfaces would remount every page.
    });
  }, [queryEpochGate, requestSettleRelayout, residentEngine]);

  useEffect(() => {
    unmountedRef.current = false;
    return () => {
      unmountedRef.current = true;
      if (paintedCaretIdleTimerRef.current !== null) {
        clearTimeout(paintedCaretIdleTimerRef.current);
      }
      if (dispatchHoldTimerRef.current !== null) {
        clearTimeout(dispatchHoldTimerRef.current);
      }
      // A build the worker leaves unanswered must not fall back on the freed session.
      generationRef.current += 1;
      workerRef.current?.client.destroy();
      workerRef.current = null;
      queryEpochGate.clear();
      markSettled(null, new Error('The editor was unmounted'), true);
    };
  }, [markSettled, queryEpochGate]);

  const publishQuerySnapshot = useCallback(
    (nextSnapshot: RustDisplayListSnapshot, contentEpoch: number): void => {
      const queries = nextSnapshot.queries;
      if (!queries) {
        if (contentEpoch === contentEpochRef.current) queryEpochGate.clear();
        return;
      }
      if (contentEpoch !== contentEpochRef.current) return;
      queryEpochGate.invalidate();
      const publish = (): void => {
        if (
          snapshotRef.current !== nextSnapshot ||
          contentEpoch !== contentEpochRef.current ||
          !queries.isReady()
        ) {
          return;
        }
        queryEpochGate.publish({
          queries,
          frameEpoch: nextSnapshot.frame?.frameEpoch ?? null,
        });
      };
      if (queries.isReady()) {
        publish();
        return;
      }
      void queries.whenReady().then(publish, () => {
        if (
          snapshotRef.current === nextSnapshot &&
          contentEpoch === contentEpochRef.current
        ) {
          queryEpochGate.clear();
        }
      });
    },
    [queryEpochGate]
  );

  const resolveQueries = useCallback<ResolveDisplayListQueries>(
    (minimumFrameEpoch = null) => queryEpochGate.resolve(minimumFrameEpoch),
    [queryEpochGate]
  );

  // Handle acquisition (page-delta serialization into the Rust query store) is
  // deferred out of the keystroke path: prime the newest facade from idle time
  // so the first interaction after a typing burst pays one accumulated delta
  // at most. Superseded facades no-op their pending prime.
  useEffect(() => {
    const queries = snapshot.queries;
    if (!queries) return;
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(() => queries.prime());
      return () => cancelIdleCallback(id);
    }
    const id = setTimeout(() => queries.prime(), 200);
    return () => clearTimeout(id);
  }, [snapshot.queries]);

  // A worker that ran out of memory is replaced by a fresh one once. The main
  // thread never takes over its work: its memory has the same limit and
  // already holds the document. `retry` asks the caller to use the current
  // worker, `stale` means the failed worker no longer serves this engine.
  // A worker of a document that another load replaced fails nothing of the new one.
  const replaceOutOfMemoryWorker = useCallback(
    (
      hostEngine: YrsSession,
      worker: { client: ResidentEngineWorkerClient; load: number } | null,
      failure: ResidentWorkerOutOfMemoryError
    ): 'retry' | 'stale' | 'failed' => {
      if (worker && worker.load !== documentLoadsRef.current) return 'stale';
      const client = worker?.client ?? null;
      const previous = outOfMemoryRef.current.has(hostEngine);
      if (outOfMemoryRef.current.get(hostEngine)) return 'failed';
      if (unmountedRef.current || workerFallbackEngineRef.current === hostEngine) return 'stale';
      const current = workerRef.current;
      if (current && current.engine !== hostEngine) return 'stale';
      // Another request of the failed worker already replaced it.
      if (current?.client !== client) return previous ? 'retry' : 'stale';
      current.client.destroy();
      workerRef.current = null;
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      if (!previous) {
        outOfMemoryRef.current.set(hostEngine, null);
        console.warn(
          '[CanvasRenderer] Resident engine worker ran out of memory; starting a fresh worker',
          failure
        );
        return 'retry';
      }
      outOfMemoryRef.current.set(hostEngine, failure);
      console.error('[CanvasRenderer] Resident engine worker ran out of memory again', failure);
      queryEpochGate.clear();
      setError(failure);
      setErrorEngine(hostEngine);
      setLoading(false);
      // A display-only preview's failure fails no wait: the full session replaces it.
      if (hostEngine.isDisplayOnly?.() !== true) markSettled(null, failure, true);
      return 'failed';
    },
    [markSettled, queryEpochGate, setWorkerPresentationActive]
  );

  const applyResidentInput = useCallback(
    (operation: ResidentInputOperation): Promise<ResidentFrameApplyResult | null> => {
      const documentLoad = documentLoadsRef.current;
      const replayInputOnMainThread = async (
        pending: ResidentInputOperation,
        hostEngine: YrsSession,
        frameEpoch: number,
        paintToken: number
      ): Promise<ResidentFrameApplyResult | null> => {
        adoptHostEngine(hostEngine);
        if (workerRef.current?.engine === hostEngine) {
          workerRef.current.client.destroy();
          workerRef.current = null;
        }
        setWorkerSurfacesActive(false);
        setWorkerPresentationActive(false);
        let encoded: Uint8Array;
        let deletedUnits: number | undefined;
        suppressWorkerInvalidationRef.current += 1;
        try {
          if (pending.kind === 'insert') {
            encoded = hostEngine.applyInput(pending.text, frameEpoch);
          } else {
            encoded = hostEngine.applyDelete(pending.direction, frameEpoch, pending.count);
            deletedUnits = hostEngine.residentDeletedUnits();
          }
        } catch (error) {
          suppressWorkerInvalidationRef.current -= 1;
          if (
            error instanceof Error &&
            error.message.includes('resident input state is not ready')
          ) {
            return null;
          }
          throw error;
        }
        suppressWorkerInvalidationRef.current -= 1;
        const delta = decodeFrameDelta(encoded);
        const previous = snapshotRef.current;
        const nextFrame = delta.full
          ? applyFrameDeltaOwned(null, delta)
          : applyFrameDeltaOwned(previous.frame, delta);
        const hostSelection = hostEngine.selection();
        const caret = residentCaretForSelection(
          hostEngine.residentCaretSnapshot(),
          hostSelection,
          hostSelection,
          nextFrame
        );
        const nextSnapshot = createRustDisplayListSnapshot(
          nextFrame.displayList,
          nextFrame,
          caret,
          null,
          { ...previous, queries: null },
          readSessionVersion(hostEngine),
          UNKNOWN_REVISION_PREVIEW_KEY
        );
        generationRef.current += 1;
        snapshotRef.current = nextSnapshot;
        publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
        markSettled(contentEpochRef.current);
        applyPaintedCaretReply(false, paintToken);
        // The host's retained layout may show another preview: lay out again for a known one.
        setTimeout(() => requestLayoutRef.current?.(), 0);
        return { frameEpoch: nextFrame.frameEpoch, caretSynchronized: false, deletedUnits };
      };
      // An input to a document another load replaced, or an unmount freed, applies to nothing.
      const dropped = (): boolean =>
        unmountedRef.current || documentLoadsRef.current !== documentLoad;
      const run = async (): Promise<ResidentFrameApplyResult | null> => {
        if (dropped()) return { frameEpoch: null, caretSynchronized: false };
        const worker = workerRef.current;
        const currentFrame = snapshotRef.current.frame;
        // A worker still on the session being replaced does not take the new one's input.
        if (
          !worker ||
          worker.engine !== residentEngineRef.current ||
          !worker.client.isReady() ||
          !currentFrame
        ) {
          return null;
        }
        const selection = worker.engine.selection();
        if (!selection) return null;
        const dispatchedEpoch = contentEpochRef.current;
        paintedCaretMachine.noteInput(performance.now());
        const paintCaret = workerPresentationActiveRef.current;
        const paintToken = paintedCaretMachine.token();
        residentPaintInflightRef.current += 1;
        let result;
        let workerDelta: ReturnType<typeof decodeFrameDelta> | undefined;
        try {
          result =
            operation.kind === 'insert'
              ? await worker.client.applyInput(
                  operation.text,
                  selection,
                  currentFrame.frameEpoch,
                  false,
                  paintCaret
                )
              : await worker.client.applyDelete(
                  operation.direction,
                  selection,
                  currentFrame.frameEpoch,
                  false,
                  paintCaret,
                  operation.count
                );
          if (result.applied) {
            try {
              workerDelta = decodeFrameDelta(result.frame);
            } catch (error) {
              throw new ResidentWorkerFailureError(
                `Resident engine worker returned an undecodable FrameDelta: ${error instanceof Error ? error.message : String(error)}`
              );
            }
          }
        } catch (error) {
          if (error instanceof ResidentWorkerOutOfMemoryError) {
            // The edit never reached the host: it takes the structural path there,
            // and the next frame comes from a fresh worker, or reports the failure.
            replaceOutOfMemoryWorker(worker.engine, worker, error);
            return null;
          }
          if (!(error instanceof ResidentWorkerFailureError)) throw error;
          if (dropped()) return { frameEpoch: null, caretSynchronized: false };
          console.error(
            '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
            error
          );
          return replayInputOnMainThread(
            operation,
            worker.engine,
            currentFrame.frameEpoch,
            paintToken
          );
        } finally {
          residentPaintInflightRef.current -= 1;
        }
        if (dropped()) return { frameEpoch: null, caretSynchronized: false };
        if (!result.applied) return null;
        const delta = workerDelta ?? decodeFrameDelta(result.frame);
        suppressWorkerInvalidationRef.current += 1;
        try {
          for (const update of result.updates) worker.engine.applyLocalUpdate(update);
        } finally {
          suppressWorkerInvalidationRef.current -= 1;
        }
        const previous = snapshotRef.current;
        if (previous.frame && delta.frameEpoch <= previous.frame.frameEpoch) {
          // Superseded: a newer frame's reply owns the painted-caret verdict.
          return { frameEpoch: null, caretSynchronized: false, deletedUnits: result.deletedUnits };
        }
        const nextFrame = applyFrameDeltaOwned(previous.frame, delta);
        const caret = residentCaretForSelection(
          result.caret,
          result.selection,
          worker.engine.selection(),
          nextFrame
        );
        if (contentEpochRef.current !== dispatchedEpoch) {
          // Another change reached the session while the worker computed this frame: show its
          // pixels, but publish no queries and leave rendering unsettled until a fresh layout.
          const overtaken: RustDisplayListSnapshot = {
            displayList: nextFrame.displayList,
            frame: nextFrame,
            queries: null,
            caret,
          };
          snapshotRef.current = overtaken;
          setSnapshot(overtaken);
          setError(null);
          setLoading(false);
          requestSettleRelayout();
          setTimeout(() => requestLayoutRef.current?.(), 0);
          applyPaintedCaretReply(false, paintToken);
          return {
            frameEpoch: nextFrame.frameEpoch,
            caretSynchronized: false,
            deletedUnits: result.deletedUnits,
          };
        }
        const nextSnapshot = createRustDisplayListSnapshot(
          nextFrame.displayList,
          nextFrame,
          caret,
          null,
          previous,
          readSessionVersion(worker.engine),
          workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision)
        );
        // Supersede an older async compatibility build before publishing the
        // frame produced by the edit transaction.
        generationRef.current += 1;
        snapshotRef.current = nextSnapshot;
        publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
        markSettled(contentEpochRef.current);
        applyPaintedCaretReply(Boolean(result.caretPainted && caret?.caretRect), paintToken);
        return {
          frameEpoch: nextFrame.frameEpoch,
          caretSynchronized: Boolean(
            caret?.caretRect &&
              workerPresentationActiveRef.current &&
              !displayListNeedsHostImages(nextFrame.displayList)
          ),
          deletedUnits: result.deletedUnits,
        };
      };
      const pending = workerInputQueueRef.current.then(run, run);
      workerInputQueueRef.current = pending.then(
        () => undefined,
        () => undefined
      );
      return pending.catch((error) => {
        const nextError =
          error instanceof Error ? error : new Error(`Resident input failed: ${String(error)}`);
        // This Rust rejection happens before the edit transaction. Let the
        // caller use the structural compatibility path for unsupported
        // paragraphs (inline media/float-dependent measurement, non-body
        // stories, or a frame that has not established resident state).
        if (nextError.message.includes('resident input state is not ready')) return null;
        console.error('[CanvasRenderer] Resident input failed', nextError);
        // An input to a document another load replaced fails nothing of the new one.
        if (!dropped()) {
          queryEpochGate.clear();
          setError(nextError);
          setErrorEngine(residentEngineRef.current);
          markSettled(null, nextError);
        }
        // Once invoked, never fall through to the legacy op: a worker failure
        // may have happened after committing the transaction.
        return { frameEpoch: null, caretSynchronized: false };
      });
    },
    [
      adoptHostEngine,
      applyPaintedCaretReply,
      markSettled,
      paintedCaretMachine,
      publishQuerySnapshot,
      queryEpochGate,
      replaceOutOfMemoryWorker,
      requestSettleRelayout,
    ]
  );

  const applyInput = useCallback(
    (text: string) => applyResidentInput({ kind: 'insert', text }),
    [applyResidentInput]
  );

  const applyDelete = useCallback(
    (direction: 'backward' | 'forward', count = 1) =>
      applyResidentInput({ kind: 'delete', direction, count }),
    [applyResidentInput]
  );

  // The display extras this hook builds frames with, minus the header/footer
  // payload a layout supplies.
  const frameExtrasInputs = useCallback((): Pick<
    Parameters<typeof encodeDisplayListFrameExtras>[0],
    'fontChains' | 'resolvedCommentIds'
  > => {
    const fontChains = fontChainsProviderRef?.current?.();
    const resolved = resolvedCommentIdsRef.current;
    return {
      ...(fontChains ? { fontChains } : {}),
      ...(resolved && resolved.size > 0
        ? { resolvedCommentIds: [...resolved].sort((a, b) => a - b) }
        : {}),
    };
  }, [fontChainsProviderRef]);

  // Stop using the worker for `hostEngine`: its frames and queries go, and
  // every later layout and frame runs on the main thread.
  const dropWorker = useCallback(
    (hostEngine: YrsSession): void => {
      adoptHostEngine(hostEngine);
      if (workerRef.current?.engine === hostEngine) {
        workerRef.current.client.destroy();
        workerRef.current = null;
      }
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
    },
    [adoptHostEngine, setWorkerPresentationActive]
  );

  const shownFrameEngine = useCallback((): unknown => frameEngineRef.current, []);
  const release = useCallback((): void => {
    // A failed load's failure holds for later waits until the next load.
    replacedLayoutRef.current = { layout: null };
    cancelPageBuilds(pageBuildTimerRef);
    workerRef.current?.client.destroy();
    workerRef.current = null;
    frameEngineRef.current = null;
    setPresentedEngine(null);
    setWorkerSurfacesActive(false);
    setWorkerPresentationActive(false);
  }, [setWorkerPresentationActive]);

  /** The frame `engine`'s next frame applies to. */
  const frameBase = useCallback(
    (engine: unknown): RetainedFrame | null =>
      frameEngineRef.current === engine ? snapshotRef.current.frame : null,
    []
  );

  // Build the unbuilt pages of the worker frame: those the viewport shows
  // first, then the rest while the main thread is idle, since accessibility
  // mirrors and printing read every page's content.
  const buildUnbuiltPages = useCallback(
    (idle = false): void => {
      pageBuildTimerRef.current = null;
      if (pageBuildInFlightRef.current) return;
      const worker = workerRef.current;
      const frame = snapshotRef.current.frame;
      if (!worker || !worker.client.isReady() || !frame) return;
      // A worker handed to another session builds its pages once it has
      // presented that session's frame.
      if (frameEngineRef.current !== worker.engine) return;
      const pages = frame.displayList.pages;
      const [start, end] = displayWindowRef.current;
      const unbuilt: number[] = [];
      for (let index = 0; index < pages.length; index += 1) {
        if (pages[index]?.unbuilt) unbuilt.push(index);
      }
      if (unbuilt.length === 0) return;
      let batch = unbuilt.filter((index) => index >= start && index < end);
      if (batch.length === 0) {
        const settling = settleWaitersRef.current.size > 0;
        if (!settling && !idle) {
          schedulePageBuildsWhenIdleRef.current();
          return;
        }
        const distance = (index: number) => (index < start ? start - index : index - end + 1);
        batch = unbuilt
          .sort((a, b) => distance(a) - distance(b))
          .slice(0, settling ? SETTLE_BUILD_BATCH_PAGES : BACKGROUND_BUILD_BATCH_PAGES)
          .sort((a, b) => a - b);
      }
      pageBuildInFlightRef.current = true;
      const dispatchedEpoch = contentEpochRef.current;
      const paintToken = paintedCaretMachine.token();
      const paintCaret =
        workerPresentationActiveRef.current && paintedCaretMachine.shouldPaint(performance.now());
      const failed = (cause: unknown): void => {
        if (workerRef.current !== worker) return;
        if (cause instanceof ResidentWorkerOutOfMemoryError) {
          if (replaceOutOfMemoryWorker(worker.engine, worker, cause) === 'retry') {
            requestLayoutRef.current?.();
          }
          return;
        }
        console.error(
          '[CanvasRenderer] Building display pages failed; falling back to the main-thread engine',
          cause
        );
        dropWorker(worker.engine);
        requestLayoutRef.current?.();
      };
      void worker.client.buildPages(batch, frame.frameEpoch, paintCaret).then(
        (result) => {
          pageBuildInFlightRef.current = false;
          if (workerRef.current !== worker) return;
          if (frameEngineRef.current !== worker.engine) {
            schedulePageBuildsWhenIdleRef.current();
            return;
          }
          try {
            const previous = snapshotRef.current;
            const delta = decodeFrameDelta(result.frame);
            if (previous.frame && delta.frameEpoch <= previous.frame.frameEpoch) return;
            const nextFrame = applyFrameDeltaOwned(previous.frame, delta);
            const caret = residentCaretForSelection(
              result.caret,
              result.selection,
              worker.engine.selection(),
              nextFrame
            );
            const nextSnapshot =
              contentEpochRef.current === dispatchedEpoch
                ? createRustDisplayListSnapshot(
                    nextFrame.displayList,
                    nextFrame,
                    caret,
                    null,
                    previous,
                    readSessionVersion(worker.engine),
                    workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision)
                  )
                : { displayList: nextFrame.displayList, frame: nextFrame, queries: null, caret };
            snapshotRef.current = nextSnapshot;
            publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
            setSnapshot(nextSnapshot);
            applyPaintedCaretReply(Boolean(result.caretPainted && caret?.caretRect), paintToken);
          } catch (error) {
            failed(error);
            return;
          }
          for (const waiter of [...settleWaitersRef.current]) waiter();
        },
        (error) => {
          pageBuildInFlightRef.current = false;
          failed(error);
        }
      );
    },
    [
      applyPaintedCaretReply,
      dropWorker,
      paintedCaretMachine,
      publishQuerySnapshot,
      replaceOutOfMemoryWorker,
    ]
  );

  const schedulePageBuilds = useCallback(
    (delay: number): void => {
      cancelPageBuilds(pageBuildTimerRef);
      pageBuildTimerRef.current = setTimeout(() => buildUnbuiltPages(), delay);
    },
    [buildUnbuiltPages]
  );
  schedulePageBuildsWhenIdleRef.current = () => {
    cancelPageBuilds(pageBuildTimerRef);
    pageBuildTimerRef.current =
      typeof requestIdleCallback === 'function'
        ? { idle: requestIdleCallback(() => buildUnbuiltPages(true)) }
        : setTimeout(() => buildUnbuiltPages(true), BACKGROUND_BUILD_DELAY_MS);
  };

  const setDisplayWindow = useCallback(
    (start: number, end: number): void => {
      const current = displayWindowRef.current;
      if (current[0] === start && current[1] === end) return;
      displayWindowRef.current = [start, end];
      schedulePageBuilds(0);
    },
    [schedulePageBuilds]
  );

  useEffect(() => {
    if (!snapshot.frame?.displayList.pages.some((page) => page.unbuilt)) return;
    schedulePageBuilds(pageBuildInFlightRef.current ? 50 : 16);
  }, [schedulePageBuilds, snapshot.frame]);

  useEffect(() => () => cancelPageBuilds(pageBuildTimerRef), []);

  const layoutInWorker = useCallback<LayoutInWorker>(
    (hostEngine, request) => {
      if (
        overrides?.build ||
        !canUseResidentEngineWorker() ||
        !isWorkerHostEngine(hostEngine) ||
        !hostEngine.adoptResidentWorkerLayout ||
        workerFallbackEngineRef.current === hostEngine ||
        handedOverEnginesRef.current.has(hostEngine)
      ) {
        return null;
      }
      const outOfMemory = outOfMemoryRef.current.get(hostEngine);
      if (outOfMemory) return Promise.reject(outOfMemory);
      const owner = workerFor(hostEngine);
      const worker = owner.client;
      const bootstrapping = !worker.bootstrapSent();
      const previousFrame = bootstrapping ? null : frameBase(hostEngine);
      const adoptedRevision = hostEngine.adoptResidentWorkerLayout(request);
      const snapshot = hostEngine.residentWorkerSnapshot(
        bootstrapping
          ? {}
          : {
              knownStateVector: worker.remoteStateVector(),
              knownFontsRevision: worker.syncedFontsRevision(),
            }
      );
      if (!snapshot) return null;
      const previewKey = layoutPreviewKey(request) ?? '';
      rememberWorkerPreview(workerPreviewKeysRef.current, snapshot, previewKey);
      if (layoutPreviewKeyRef.current !== null && previewKey !== layoutPreviewKeyRef.current) {
        contentEpochRef.current += 1;
        queryEpochGate.invalidate();
      }
      layoutPreviewKeyRef.current = previewKey;
      const contentEpoch = contentEpochRef.current;
      const options = {
        layoutExtras: JSON.stringify(frameExtrasInputs()),
        stateVector: hostEngine.encodeStateVector(),
        displayWindow: displayWindowRef.current,
        ...(bootstrapping
          ? {
              provisionalPages: PROVISIONAL_LAYOUT_PAGES,
              heapLimitBytes: workerHeapLimitRef.current,
              ...followedFrameEpoch(snapshotRef.current.frame),
            }
          : {}),
      };
      const paintCaret =
        !bootstrapping &&
        workerPresentationActiveRef.current &&
        paintedCaretMachine.shouldPaint(performance.now());
      const reply = bootstrapping
        ? worker.bootstrap(snapshot, '', options)
        : worker.sync(snapshot, '', previousFrame?.frameEpoch ?? 0, paintCaret, options);
      // A worker out of memory runs the pass again in a fresh worker; once
      // that one runs out too, the pass rejects and nothing lays out here.
      const unavailable = (
        cause: unknown
      ): Promise<WorkerLayoutComputation | null> | null => {
        const current = workerRef.current;
        if (owner.load !== documentLoadsRef.current || (current && current.engine !== hostEngine)) {
          return null;
        }
        // A session whose replacement worker ran out of memory too lays out nowhere.
        const outOfMemory = cause instanceof ResidentWorkerOutOfMemoryError;
        if (outOfMemory && outOfMemoryRef.current.get(hostEngine)) return Promise.reject(cause);
        // A pass of a session no worker serves any more starts no worker.
        if (!current) return null;
        if (outOfMemory) {
          // A newer layout, here or in a worker, replaced this pass: the host
          // drops it, and a newer worker request recovers the worker it asks.
          if (hostEngine.residentWorkerProbe()?.layoutRevision !== adoptedRevision) return null;
          const outcome = replaceOutOfMemoryWorker(hostEngine, owner, cause);
          if (outcome === 'failed') return Promise.reject(cause);
          if (outcome === 'stale') return null;
          return layoutInWorkerRef.current?.(hostEngine, request) ?? null;
        }
        if (current?.client !== worker) return null;
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; laying out on the main thread',
          cause
        );
        dropWorker(hostEngine);
        return null;
      };
      // `base` is the frame the reply's frame applies to; without one the
      // display builds its own frame for the layout.
      const adopt = (
        result: ResidentEngineWorkerFrame,
        base: RetainedFrame | null | undefined
      ): LayoutComputation => {
        if (result.layoutJson === undefined) {
          throw new ResidentWorkerFailureError('Resident engine worker omitted its layout');
        }
        const computation = workerLayoutComputation(result.layoutJson, result.layoutRevision);
        if (base === undefined) return computation;
        workerLayoutFramesRef.current.set(computation.layout, {
          result,
          previousFrame: base,
          engine: hostEngine,
          contentEpoch,
          layoutExtras: options.layoutExtras,
          provisional: result.layoutProvisional === true,
        });
        return computation;
      };
      return reply
        .then((result): WorkerLayoutComputation => {
          const computation = adopt(result, previousFrame);
          // A display-only preview is replaced by the full document before
          // anything needs the rest of its pages.
          if (!result.layoutProvisional || hostEngine.isDisplayOnly?.()) return computation;
          // The rest is laid out against the provisional frame, once its page
          // surfaces are attached: the worker answers in order, so asking
          // sooner would hold back the first paint until it is done.
          const provisionalEpoch = result.caret.frameEpoch;
          const surfaced = new Promise<void>((resolve) => {
            completionGateRef.current = resolve;
            setTimeout(resolve, PROVISIONAL_SURFACE_WAIT_MS);
          });
          const complete = surfaced
            // A worker handed to another session lays out that session now.
            .then(() =>
              workerRef.current?.engine === hostEngine
                ? worker.completeLayout(provisionalEpoch, false, COMPLETION_SLICE_BLOCKS)
                : null
            )
            .then((completed) => {
              if (!completed) return null;
              const base = frameBase(hostEngine);
              return adopt(completed, base?.frameEpoch === provisionalEpoch ? base : undefined);
            })
            .catch(async (cause: unknown): Promise<LayoutComputation | null> => {
              const retried = await unavailable(cause);
              return retried?.complete ?? retried;
            });
          // A pass the host drops never observes this; the renderer reports the failure.
          complete.catch(() => {});
          return { ...computation, complete };
        })
        .catch(unavailable);
    },
    [
      dropWorker,
      frameBase,
      frameExtrasInputs,
      overrides?.build,
      paintedCaretMachine,
      queryEpochGate,
      replaceOutOfMemoryWorker,
      workerFor,
    ]
  );
  const layoutInWorkerRef = useRef(layoutInWorker);
  layoutInWorkerRef.current = layoutInWorker;

  const attachOffscreenCanvases = useCallback(
    async (
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number,
      caretStyle: ResidentCaretPaintStyle
    ): Promise<boolean> => {
      const current = workerRef.current;
      // Queue even while the worker is mid-invalidation: requests are handled
      // FIFO, so an attach lands after the sync that follows and the worker
      // rasters the newly attached surfaces itself. Refusing here would strand
      // already-transferred canvases (they cannot be re-transferred).
      if (!current) return false;
      const attached = current.client.attachCanvases(
        pages,
        activePageIds,
        devicePixelRatio,
        zoom,
        caretStyle
      );
      completionGateRef.current?.();
      completionGateRef.current = null;
      try {
        await attached;
      } catch (error) {
        if (!(error instanceof ResidentWorkerOutOfMemoryError)) throw error;
        if (
          workerRef.current === current &&
          replaceOutOfMemoryWorker(current.engine, current, error) === 'retry'
        ) {
          requestLayoutRef.current?.();
        }
        return false;
      }
      return true;
    },
    [replaceOutOfMemoryWorker]
  );

  useEffect(() => {
    if (replacedLayoutRef.current && layout && layout !== replacedLayoutRef.current.layout) {
      replacedLayoutRef.current = null;
    }
    if (!layout) {
      // layout reset (document change) — drop the stale pages
      generationRef.current++;
      contentEpochRef.current += 1;
      layoutPreviewKeyRef.current = null;
      queryEpochGate.clear();
      snapshotRef.current = EMPTY_DISPLAY_LIST_SNAPSHOT;
      frameEngineRef.current = null;
      recoveryFrameEpochRef.current = 0;
      setSnapshot(EMPTY_DISPLAY_LIST_SNAPSHOT);
      setError(null);
      setLoading(true);
      settledEpochRef.current = null;
      // A failure of the document being loaded holds until its layout or the next load.
      if (!replacedLayoutRef.current) settleErrorRef.current = null;
      // The worker of a session the renderer let go of never builds again.
      if (workerRef.current && workerRef.current.engine !== residentEngine) {
        workerRef.current.client.destroy();
        workerRef.current = null;
      }
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      notifyCaretInterrupt();
      return;
    }
    queryEpochGate.invalidate();
    // A new layout settles with its own display list, not with the last one.
    settledEpochRef.current = null;
    const previewKey = revisionPreviewKeyOf(layout);
    if (
      previewKey !== null &&
      layoutPreviewKeyRef.current !== null &&
      previewKey !== layoutPreviewKeyRef.current
    ) {
      contentEpochRef.current += 1;
    }
    if (previewKey !== null) layoutPreviewKeyRef.current = previewKey;
    const contentEpoch = contentEpochRef.current;
    const sourceVersion = sourceVersionOf(layout);
    const inputs = (overrides?.getInputs ?? getLayoutKernelInputs)(layout);
    const generation = ++generationRef.current;
    const documentLoad = documentLoadsRef.current;
    // A display-only preview's failure fails no wait: the full session replaces it.
    const settleFailure = (failure: Error): void => {
      const preview = engine as { isDisplayOnly?: () => boolean } | null | undefined;
      if (preview?.isDisplayOnly?.() !== true) markSettled(null, failure);
    };
    if (!inputs) {
      const failure = new Error('No display-list inputs were recorded for the current layout.');
      queryEpochGate.clear();
      setError(failure);
      setErrorEngine(engine ?? null);
      setLoading(false);
      settleFailure(failure);
      return;
    }
    const build = overrides?.build ?? buildRustDisplayList;
    // Merged doc-wide font chains from the Rust measure source (when active).
    // A non-empty map activates GlyphRun emission; absent ⇒ TextRunPrimitive.
    const fontChains = fontChainsProviderRef?.current?.();
    // Getters so the worker-rendered path (extras only) never materializes
    // the retained measured arena; the main-thread fallback pays the fetch once.
    const buildInputs = {
      get measured() {
        return inputs.measured;
      },
      get options() {
        return inputs.options;
      },
      layout,
      ...(inputs.headersFooters ? { headersFooters: inputs.headersFooters } : {}),
      ...(fontChains ? { fontChains } : {}),
      ...(resolvedCommentIds && resolvedCommentIds.size > 0
        ? { resolvedCommentIds: [...resolvedCommentIds].sort((a, b) => a - b) }
        : {}),
    };
    const workerEligible =
      residentEngine !== null && workerFallbackEngineRef.current !== residentEngine;
    // Cheap probe only: the full snapshot (document state, font bytes) is
    // built lazily below, and only for bootstrap/sync — steady-state frame
    // builds never encode state or copy fonts.
    const probe = workerEligible ? residentEngine.residentWorkerProbe() : null;
    const buildOnMainThread = () => {
      if (residentEngine?.residentLayoutInWorker?.()) {
        return Promise.reject(new MainThreadLayoutPendingError());
      }
      if (overrides?.build) {
        return build(buildInputs, engine ?? undefined).then((displayList) => ({
          displayList,
          frame: null as RetainedFrame | null,
          caret: null as YrsResidentCaretSnapshot | null,
          queryEngine: engine,
          workerProduced: false,
          caretPainted: false,
        }));
      }
      // A frame engine paints the pagination it retains, which a newer layout may have replaced.
      const retainedRevision = residentEngine?.residentWorkerProbe()?.layoutRevision;
      const base = frameBase(residentEngine ?? engine ?? null);
      return buildRustDisplayFrame(
        buildInputs,
        engine ?? undefined,
        base,
        base?.frameEpoch ?? recoveryFrameEpochRef.current
      ).then((result) => ({
        ...result,
        caret: null as YrsResidentCaretSnapshot | null,
        queryEngine: engine,
        workerProduced: false,
        caretPainted: false,
        ...(residentEngine && result.frame && retainedRevision !== inputs.layoutRevision
          ? { previewKey: UNKNOWN_REVISION_PREVIEW_KEY }
          : {}),
      }));
    };
    const paintToken = paintedCaretMachine.token();
    let pending: Promise<BuiltDisplay>;
    const outOfMemory = residentEngine ? outOfMemoryRef.current.get(residentEngine) : null;
    if (residentEngine && handedOverEnginesRef.current.has(residentEngine)) {
      // Its successor paints these surfaces now, and its first frame replaces this one.
      pending = Promise.reject(new SupersededPreviewError());
    } else if (outOfMemory && residentEngine) {
      // A replaced document's failure fails nothing of the one being loaded.
      pending = Promise.reject(
        sessionLoad(residentEngine) === documentLoadsRef.current
          ? outOfMemory
          : new SupersededPreviewError()
      );
    } else if (!overrides?.build && probe && canUseResidentEngineWorker()) {
      const hostEngine = residentEngine;
      if (!hostEngine) throw new Error('Resident worker snapshot requires a host engine');
      // The worker this build asks; a failure of one that was since replaced
      // (a StrictMode remount's destroyed worker) must not tear down its successor.
      let requested = workerRef.current;
      const fallback = (
        cause: unknown,
        owner: { client: ResidentEngineWorkerClient; load: number } | null = null
      ): Promise<BuiltDisplay> => {
        // A replaced build falls back on nothing and recovers no worker: its
        // engine may be gone by now.
        if (generation !== generationRef.current) return Promise.reject(cause);
        if (cause instanceof ResidentWorkerOutOfMemoryError) {
          // Another load's build recovers nothing.
          if (documentLoad !== documentLoadsRef.current) return Promise.reject(cause);
          if (owner && owner.load !== documentLoadsRef.current) {
            return Promise.reject(new SupersededPreviewError());
          }
          // So does a layout adopted since this build began, such as a worker
          // pass running again, whose frame this build's snapshot would erase.
          if (hostEngine.residentWorkerProbe()?.layoutRevision !== probe.layoutRevision) {
            return Promise.reject(new SupersededPreviewError());
          }
          const outcome = replaceOutOfMemoryWorker(hostEngine, owner, cause);
          if (outcome === 'retry') {
            try {
              return requestWorkerFrame();
            } catch (error) {
              return fallback(error);
            }
          }
          return outcome === 'stale' && requested !== workerRef.current
            ? buildOnMainThread()
            : Promise.reject(cause);
        }
        if (requested !== workerRef.current) return buildOnMainThread();
        const nextError =
          cause instanceof Error
            ? cause
            : new Error(`Resident engine worker failed: ${String(cause)}`);
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
          nextError
        );
        dropWorker(hostEngine);
        return buildOnMainThread();
      };
      const requestWorkerFrame = (): Promise<BuiltDisplay> => {
        const owner = workerFor(hostEngine);
        requested = owner;
        const worker = owner.client;
        const extras = encodeDisplayListFrameExtras(buildInputs);
        const bootstrapping = !worker.bootstrapSent();
        const previousFrame = bootstrapping ? null : frameBase(hostEngine);
        // On a fresh client both hints are null, so a bootstrap snapshot is
        // always complete; a sync snapshot ships a state diff and skips font
        // bytes the worker already holds.
        const buildSnapshot = () => {
          const snapshot = hostEngine.residentWorkerSnapshot({
            knownStateVector: worker.remoteStateVector(),
            knownFontsRevision: worker.syncedFontsRevision(),
          });
          if (!snapshot) throw new Error('Resident worker snapshot was not available');
          rememberWorkerPreview(workerPreviewKeysRef.current, snapshot);
          return snapshot;
        };
        const sent = () => ({ stateVector: hostEngine.encodeStateVector() });
        // Structural text input reaches the worker as a sync/buildFrame; keep
        // the painted caret glued to those frames while the typing burst lasts.
        const paintCaret =
          !bootstrapping &&
          workerPresentationActiveRef.current &&
          paintedCaretMachine.shouldPaint(performance.now());
        const snapshot =
          bootstrapping || worker.layoutRevision() !== probe.layoutRevision
            ? buildSnapshot()
            : null;
        // The extras carry this layout's headers and footers, so they only go
        // with a worker layout of the same preview.
        const shownKey = snapshot
          ? layoutPreviewKey(snapshot.layoutInput)
          : (workerPreviewKeysRef.current.get(probe.layoutRevision) ?? null);
        const workerFrame =
          previewKey !== null && shownKey !== null && shownKey !== previewKey
            ? Promise.reject(new SupersededPreviewError())
            : !snapshot
              ? worker.buildFrame(extras, previousFrame?.frameEpoch ?? 0, paintCaret)
              : bootstrapping
                ? worker.bootstrap(snapshot, extras, {
                    ...sent(),
                    displayWindow: displayWindowRef.current,
                    heapLimitBytes: workerHeapLimitRef.current,
                    ...followedFrameEpoch(snapshotRef.current.frame),
                  })
                : worker.sync(snapshot, extras, previousFrame?.frameEpoch ?? 0, paintCaret, {
                    ...sent(),
                    displayWindow: displayWindowRef.current,
                  });
        return workerFrame
          .then((result) => {
            const delta = decodeFrameDelta(result.frame);
            const nextFrame = applyFrameDelta(previousFrame, delta);
            return {
              displayList: nextFrame.displayList,
              frame: nextFrame,
              // A replaced build is dropped below; its engine may be gone by now.
              caret:
                generation === generationRef.current
                  ? residentCaretForSelection(
                      result.caret,
                      result.selection,
                      hostEngine.selection(),
                      nextFrame
                    )
                  : null,
              queryEngine: null,
              workerProduced: true,
              caretPainted: result.caretPainted,
              previewKey: workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision),
            };
          })
          .catch((error) => {
            if (error instanceof SupersededPreviewError) throw error;
            return fallback(error, owner);
          });
      };
      const prebuilt = workerLayoutFramesRef.current.get(layout);
      if (prebuilt) workerLayoutFramesRef.current.delete(layout);
      try {
        // The frame is adopted only while nothing newer reached the session
        // or the display since the worker built it.
        if (
          prebuilt &&
          prebuilt.engine === hostEngine &&
          workerRef.current?.engine === hostEngine &&
          prebuilt.contentEpoch === contentEpoch &&
          (frameBase(hostEngine)?.frameEpoch ?? null) ===
            (prebuilt.previousFrame?.frameEpoch ?? null) &&
          prebuilt.layoutExtras === JSON.stringify(frameExtrasInputs())
        ) {
          // The worker ran this layout and built its frame in the same pass.
          const { result, previousFrame } = prebuilt;
          const delta = decodeFrameDelta(result.frame);
          const nextFrame = applyFrameDelta(previousFrame, delta);
          pending = Promise.resolve({
            displayList: nextFrame.displayList,
            frame: nextFrame,
            caret: residentCaretForSelection(
              result.caret,
              result.selection,
              hostEngine.selection(),
              nextFrame
            ),
            queryEngine: null,
            workerProduced: true,
            caretPainted: result.caretPainted,
            provisional: prebuilt.provisional,
            previewKey: workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision),
          });
        } else {
          pending = requestWorkerFrame();
        }
      } catch (error) {
        pending = fallback(error);
      }
    } else {
      pending = buildOnMainThread();
    }
    pending
      .then((result) => {
        if (
          generation !== generationRef.current ||
          contentEpoch !== contentEpochRef.current
        ) {
          if (contentEpoch !== contentEpochRef.current) requestSettleRelayout();
          return;
        }
        const nextSnapshot = createRustDisplayListSnapshot(
          result.displayList,
          result.frame,
          result.caret,
          result.queryEngine,
          snapshotRef.current,
          sourceVersion,
          result.previewKey === undefined ? previewKey : result.previewKey
        );
        snapshotRef.current = nextSnapshot;
        publishQuerySnapshot(nextSnapshot, contentEpoch);
        setSnapshot(nextSnapshot);
        frameEngineRef.current = residentEngine ?? engine ?? null;
        setPresentedEngine(residentEngine ?? engine ?? null);
        setError(null);
        setLoading(false);
        if (!result.provisional && layout.partial !== true) markSettled(contentEpoch);
        const workerProduced = Boolean(
          result.workerProduced && probe && workerRef.current?.client.isReady()
        );
        setWorkerSurfacesActive(workerProduced);
        applyPaintedCaretReply(
          Boolean(workerProduced && result.caretPainted && result.caret?.caretRect),
          paintToken
        );
      })
      .catch((error) => {
        if (error instanceof MainThreadLayoutPendingError) {
          setTimeout(() => requestLayoutRef.current?.(), 0);
          return;
        }
        if (error instanceof SupersededPreviewError) return;
        if (
          generation !== generationRef.current ||
          contentEpoch !== contentEpochRef.current
        ) {
          return;
        }
        const nextError =
          error instanceof Error ? error : new Error(`Display-list build failed: ${String(error)}`);
        console.error('[CanvasRenderer] Rust display-list build failed', nextError);
        queryEpochGate.clear();
        setError(nextError);
        setErrorEngine(engine ?? null);
        setLoading(false);
        settleFailure(nextError);
      });
  }, [
    adoptHostEngine,
    layout,
    overrides,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    residentEngine,
    dropWorker,
    frameExtrasInputs,
    setWorkerPresentationActive,
    paintedCaretMachine,
    applyPaintedCaretReply,
    notifyCaretInterrupt,
    publishQuerySnapshot,
    queryEpochGate,
    markSettled,
    replaceOutOfMemoryWorker,
    requestSettleRelayout,
    sessionLoad,
    workerFor,
  ]);

  const resetSettled = useCallback(
    (failure: Error | null = null): void => {
      if (!failure) {
        contentEpochRef.current += 1;
        documentLoadsRef.current += 1;
        replacedLayoutRef.current = { layout: layoutRef.current };
      }
      markSettled(null, failure, true);
    },
    [markSettled]
  );
  const awaitingDocument = useCallback((): boolean => replacedLayoutRef.current !== null, []);
  const recordSession = useCallback(
    (session: YrsSession | null): void => {
      if (session) sessionLoad(session);
    },
    [sessionLoad]
  );

  const settledDisplayList = useCallback(
    (relayout: (() => void) | null, timeoutMs: number | null = 15_000): Promise<DisplayList> =>
      new Promise<DisplayList>((resolve, reject) => {
        let timer: ReturnType<typeof setTimeout> | undefined;
        const settle = (): boolean => {
          const failure = settleErrorRef.current;
          const displayList = snapshotRef.current.displayList;
          const current =
            displayList !== null &&
            settledEpochRef.current === contentEpochRef.current &&
            !displayList.pages.some((page) => page.unbuilt);
          if (!failure && !current) {
            if (displayList?.pages.some((page) => page.unbuilt)) schedulePageBuilds(0);
            return false;
          }
          settleWaitersRef.current.delete(waiter);
          if (timer !== undefined) clearTimeout(timer);
          if (failure) reject(failure);
          else resolve(displayList!);
          return true;
        };
        const waiter = (): void => {
          settle();
        };
        if (settle()) return;
        if (relayout) settleRelayoutRef.current = relayout;
        settleWaitersRef.current.add(waiter);
        if (timeoutMs !== null) {
          timer = setTimeout(() => {
            settleWaitersRef.current.delete(waiter);
            reject(new Error('The document did not finish rendering'));
          }, timeoutMs);
        }
        relayout?.();
      }),
    [schedulePageBuilds]
  );

  const workerMemory = useCallback(
    (): WasmModuleMemory[] | null => workerRef.current?.client.memory() ?? null,
    []
  );

  return {
    displayList: snapshot.displayList,
    error,
    errorEngine,
    loading,
    frame: snapshot.frame,
    queries: snapshot.queries,
    resolveQueries,
    settledDisplayList,
    resetSettled,
    awaitingDocument,
    recordSession,
    caret: snapshot.caret,
    applyInput,
    applyDelete,
    layoutInWorker,
    presentedEngine,
    shownFrameEngine,
    release,
    setDisplayWindow,
    workerMemory,
    workerSurfacesActive,
    workerPresentationActive,
    setWorkerPresentationActive,
    attachOffscreenCanvases,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  };
}

/** The preview key a region layout request renders with, or null when it cannot be read. */
function layoutPreviewKey(request: unknown): string | null {
  if (typeof request !== 'string') return null;
  try {
    return revisionPreviewKey(
      (JSON.parse(request) as { renderEnv?: YrsRenderEnv }).renderEnv?.revisionPreview
    );
  } catch {
    return null;
  }
}

/** The preview key of the layout revision a worker frame shows; unknown once forgotten. */
function workerPreviewKey(keys: Map<number, string>, revision: number): string {
  return keys.get(revision) ?? UNKNOWN_REVISION_PREVIEW_KEY;
}

/** Records the preview key of the layout a worker snapshot carries, by its layout revision. */
function rememberWorkerPreview(
  keys: Map<number, string>,
  snapshot: YrsResidentWorkerSnapshot,
  key = layoutPreviewKey(snapshot.layoutInput)
): void {
  if (key === null || typeof snapshot.layoutRevision !== 'number') return;
  keys.set(snapshot.layoutRevision, key);
  for (const revision of keys.keys()) {
    if (keys.size <= 8) break;
    keys.delete(revision);
  }
}

/**
 * `sourceVersion` and `previewKey`: the document version and revision preview the frame's pixels
 * and queries show.
 */
function createRustDisplayListSnapshot(
  displayList: DisplayList,
  frame: RetainedFrame | null,
  caret: YrsResidentCaretSnapshot | null,
  engine: RustDisplayListEngine | null | undefined,
  previous: RustDisplayListSnapshot,
  sourceVersion: string | null,
  previewKey: string | null
): RustDisplayListSnapshot {
  const residentQueries = residentDisplayListQueryEngine(engine);
  const queries = createDisplayListQueries(displayList, residentQueries, previous.queries);
  stampSourceVersion(queries, sourceVersion);
  if (previewKey !== null) stampRevisionPreviewKey(queries, previewKey);
  return { displayList, frame, queries, caret };
}

function residentCaretForSelection(
  caret: YrsResidentCaretSnapshot,
  computedFor: YrsSelection | null,
  current: YrsSelection | null,
  frame: RetainedFrame
): YrsResidentCaretSnapshot | null {
  if (!computedFor || !sameYrsSelection(computedFor, current)) return null;
  const validated = residentCaretSnapshotForFrame(caret, frame);
  return validated ? { ...validated, selection: computedFor } : null;
}

/**
 * The editing wasm's own query surface, unless a panic already poisoned that
 * instance — then queries route through the layout module instead of dying on
 * the leaked borrow guard.
 */
function residentDisplayListQueryEngine(
  engine: RustDisplayListEngine | null | undefined
): ResidentDisplayListQueryEngine | undefined {
  const resident =
    engine?.displayHitTestRegionsJson &&
    engine.displayVerticalMoveJson &&
    engine.displayRangeRectsJson &&
    engine.displayRangeRectsRegionJson
      ? (engine as ResidentDisplayListQueryEngine)
      : undefined;
  return isDisplayListQuerySourceDead(resident) ? undefined : resident;
}

function isWorkerHostEngine(
  engine: RustDisplayListEngine | null | undefined
): engine is YrsSession {
  return Boolean(
    engine &&
    'residentWorkerSnapshot' in engine &&
    'onUpdate' in engine &&
    'selection' in engine &&
    'applyUpdate' in engine
  );
}

// memoized per display list: one query facade (and one JSON stringify inside
// it) per build, shared by pointer routing, selection overlay, and sidebar
// anchors. Null while the canvas is off or before the first build lands.
export function useDisplayListQueries(
  displayList: DisplayList | null,
  engine?: RustDisplayListEngine | null
): DisplayListQueries | null {
  const residentQueries = residentDisplayListQueryEngine(engine);
  // The previous facade seeds handle adoption: consecutive builds patch only
  // changed pages into the Rust query store instead of re-serializing the
  // whole display list per build.
  const previousRef = useRef<DisplayListQueries | null>(null);
  return useMemo(() => {
    const next = displayList
      ? createDisplayListQueries(displayList, residentQueries, previousRef.current)
      : null;
    previousRef.current = next;
    return next;
  }, [displayList, residentQueries]);
}

export interface UseCanvasRendererResult {
  /** latest display list built from the live document (null until the first pass lands) */
  displayList: DisplayList | null;
  /** Retained binary frame and its damaged page set. */
  frame: RetainedFrame | null;
  /** sole visible renderer lifecycle */
  status: 'loading' | 'ready' | 'error';
  /** fatal display-list error; non-null exactly while status is `error` */
  error: Error | null;
  /** See {@link UseRustDisplayListResult.errorEngine}. */
  errorEngine: unknown;
  /** The engine of the layout the display list is built from; it lags a replaced session. */
  layoutEngine: unknown;
  /** feed PagedEditor's per-pass Layout into the interaction query source */
  onLayoutComputed: (
    layout: Layout | null,
    engine?: (RustDisplayListEngine & { outlineGlyphJson?: GlyphOutlineProvider }) | null
  ) => void;
  /** Drops the pages, their engine and the resident worker, for a load that failed. */
  reset(): void;
  /** media resolver for CanvasPagesView */
  resolveImage: ImageResolver;
  /** Rust display-list query facade for adapter interactions. */
  queries: DisplayListQueries | null;
  /** Resolve the newest facade after pending edits and relayouts. */
  resolveQueries: ResolveDisplayListQueries;
  /** See {@link UseRustDisplayListResult.settledDisplayList}. */
  settledDisplayList: UseRustDisplayListResult['settledDisplayList'];
  /** See {@link UseRustDisplayListResult.resetSettled}. */
  resetSettled: UseRustDisplayListResult['resetSettled'];
  /** See {@link UseRustDisplayListResult.awaitingDocument}. */
  awaitingDocument: UseRustDisplayListResult['awaitingDocument'];
  /** See {@link UseRustDisplayListResult.recordSession}. */
  recordSession: UseRustDisplayListResult['recordSession'];
  /** Worker caret from the same atomic renderer snapshot. */
  caret: YrsResidentCaretSnapshot | null;
  /** Whether worker-presented pixels make `caret` authoritative. */
  authoritativeCaretActive: boolean;
  /** host element of the canvas pages, so pointer routing can map client → page-local coords */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Glyph outlines sourced from the same resident font store as measurement. */
  glyphOutlineProvider: GlyphOutlineProvider | null;
  /** The engine whose layout the shown frame is of. */
  presentedEngine: unknown;
  /**
   * The image resolver of the frame published last, even before the render that shows it: a
   * display list read from `settledDisplayList` may belong to a session not shown yet.
   */
  imageResolverForShownFrame(): ImageResolver;
  /** One-call ordinary text insertion; false until resident state is ready. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** One-call ordinary deletions/merges; false until resident state is ready. */
  applyDelete(
    direction: 'backward' | 'forward',
    count?: number
  ): Promise<ResidentFrameApplyResult | null>;
  /** Hands a layout pass to the resident worker; see {@link LayoutInWorker}. */
  layoutInWorker: LayoutInWorker;
  /** The pages `[start, end)` near the viewport, built before the others. */
  setDisplayWindow(start: number, end: number): void;
  /** The resident worker's wasm memories as of its latest reply; null without a worker. */
  workerMemory(): WasmModuleMemory[] | null;
  setWorkerPresentationActive(active: boolean): void;
  /** OffscreenCanvas replay bridge; null keeps DOM-canvas replay. */
  offscreenReplay: {
    attach(
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number,
      caretStyle: ResidentCaretPaintStyle
    ): Promise<boolean>;
  } | null;
  /** True while the worker-painted caret line owns the caret (DOM caret hidden). */
  paintedCaretActive: boolean;
  /** Local text input notification for the painted-caret mode machine. */
  notifyCaretInput(): void;
  /** Text input dispatched: hide the DOM caret before the worker round-trip. */
  notifyCaretInputDispatched(): void;
  /** Selection move / blur / IME / mode change: immediate swap to the DOM caret. */
  notifyCaretInterrupt(): void;
}

/**
 * Decoded images of one session's document; the next session starts empty. A frame's replay
 * reads the images of the session that built it, which the renderer holds until another
 * session's frame replaces it.
 */
export function useFrameImageResolver(engine: RustDisplayListEngine | null): ImageResolver {
  return useMemo(() => frameImageResolver(engine), [engine]);
}

function frameImageResolver(engine: RustDisplayListEngine | null): ImageResolver {
  return createCanvasImageResolver({
    media: (token) => engine?.mediaSource?.(token) ?? null,
    mediaScope: () => engine?.mediaScope?.() ?? 0,
  });
}

// bundles the canvas-renderer host wiring for DocxEditor: collects each
// layout pass and rebuilds the display list through the rust engine. Canvas is
// the only visible renderer: the first build has an explicit loading state and
// a hard failure has an explicit error state. `fontChainsProviderRef`
// is the host slot the Rust measure source fills; reading it at build time is
// what activates GlyphRun emission when Rust measurement is on.
export function useCanvasRenderer(
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>,
  // Resolved comment ids whose range wash the canvas hides; identity changes
  // rebuild the display list (resolve / reopen / expand-a-resolved-card).
  resolvedCommentIds?: ReadonlySet<number>,
  /** Asks the host for a layout of the document as it is now. */
  requestLayout?: () => void,
  /** The most a resident worker's editing core may allocate at once. */
  workerHeapLimitBytes?: number,
  /** See `useRustDisplayList`'s `handoffFromRef`. */
  handoffFromRef?: React.RefObject<YrsSession | null>
): UseCanvasRendererResult {
  const [layout, setLayout] = useState<Layout | null>(null);
  const [engine, setEngine] = useState<
    (RustDisplayListEngine & { outlineGlyphJson?: GlyphOutlineProvider }) | null
  >(null);
  const onLayoutComputed = useCallback(
    (
      next: Layout | null,
      nextEngine?: (RustDisplayListEngine & { outlineGlyphJson?: GlyphOutlineProvider }) | null
    ) => {
      setLayout(next);
      setEngine(nextEngine ?? null);
    },
    []
  );
  const {
    displayList,
    error,
    errorEngine,
    loading,
    frame,
    queries: snapshotQueries,
    resolveQueries,
    settledDisplayList,
    resetSettled,
    awaitingDocument,
    recordSession,
    caret,
    applyInput,
    applyDelete,
    layoutInWorker,
    presentedEngine,
    shownFrameEngine,
    release,
    setDisplayWindow,
    workerMemory,
    workerSurfacesActive,
    workerPresentationActive,
    setWorkerPresentationActive,
    attachOffscreenCanvases,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  } = useRustDisplayList(
    layout,
    undefined,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    requestLayout,
    workerHeapLimitBytes,
    handoffFromRef
  );
  const reset = useCallback((): void => {
    setLayout(null);
    setEngine(null);
    release();
  }, [release]);
  const imageSession = (presentedEngine ?? engine) as RustDisplayListEngine | null;
  const resolveImage = useFrameImageResolver(imageSession);
  const shownImagesRef = useRef({ engine: imageSession, resolveImage });
  shownImagesRef.current = { engine: imageSession, resolveImage };
  const imageResolverForShownFrame = useCallback((): ImageResolver => {
    const shown = (shownFrameEngine() ?? shownImagesRef.current.engine) as
      | RustDisplayListEngine
      | null;
    return shown === shownImagesRef.current.engine
      ? shownImagesRef.current.resolveImage
      : frameImageResolver(shown);
  }, [shownFrameEngine]);
  const status: UseCanvasRendererResult['status'] = error
    ? 'error'
    : loading || displayList == null
      ? 'loading'
      : 'ready';
  const [geometryReady, setGeometryReady] = useState(false);
  useEffect(() => {
    if (!snapshotQueries) {
      setGeometryReady(false);
      return;
    }
    if (snapshotQueries.isReady()) {
      setGeometryReady(true);
      return;
    }
    let cancelled = false;
    setGeometryReady(false);
    void snapshotQueries.whenReady().then(
      () => {
        if (!cancelled) setGeometryReady(true);
      },
      () => {
        if (!cancelled) setGeometryReady(false);
      }
    );
    return () => {
      cancelled = true;
    };
  }, [snapshotQueries]);
  const canvasHostRef = useRef<HTMLDivElement | null>(null);
  // Offscreen replay is the default when the browser supports transferable
  // canvas surfaces. `offscreenReplay=0` is a diagnostic escape hatch; pages
  // containing host-resolved media still select the DOM-canvas fallback in
  // CanvasPagesView.
  const offscreenAllowed = (() => {
    if (typeof window === 'undefined') return false;
    return new URLSearchParams(window.location.search).get('offscreenReplay') !== '0';
  })();
  // Keyed off surface ownership, not frame freshness: an invalidated worker
  // frame must not tear down the offscreen canvases (CanvasPagesView keys its
  // page surfaces on this), or every remote/structural edit remounts and
  // blanks the whole document.
  const offscreenReplay = useMemo(
    () =>
      workerSurfacesActive &&
      offscreenAllowed &&
      typeof OffscreenCanvas !== 'undefined' &&
      typeof HTMLCanvasElement !== 'undefined' &&
      'transferControlToOffscreen' in HTMLCanvasElement.prototype
        ? { attach: attachOffscreenCanvases }
        : null,
    [attachOffscreenCanvases, offscreenAllowed, workerSurfacesActive]
  );
  const authoritativeCaretActive = Boolean(
    workerPresentationActive &&
      displayList &&
      caret?.caretRect &&
      !displayListNeedsHostImages(displayList)
  );
  return {
    displayList,
    frame,
    // The display surface is already valid when a delta lands. Keep it mounted
    // while the replacement main-thread geometry cache warms; interactions are
    // briefly gated by a null query facade instead of throwing away every page
    // canvas and forcing a full replay.
    status,
    error,
    errorEngine,
    layoutEngine: engine,
    onLayoutComputed,
    reset,
    resolveImage,
    queries: geometryReady ? snapshotQueries : null,
    resolveQueries,
    settledDisplayList,
    resetSettled,
    awaitingDocument,
    recordSession,
    caret,
    authoritativeCaretActive,
    canvasHostRef,
    // Glyph ids are a session's own: outlines come from the engine whose
    // frame is shown, which lags `engine` while a new session's first frame builds.
    glyphOutlineProvider:
      (presentedEngine as { outlineGlyphJson?: GlyphOutlineProvider } | null)?.outlineGlyphJson ??
      engine?.outlineGlyphJson ??
      null,
    presentedEngine,
    imageResolverForShownFrame,
    applyInput,
    applyDelete,
    layoutInWorker,
    setDisplayWindow,
    workerMemory,
    setWorkerPresentationActive,
    offscreenReplay,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  };
}
