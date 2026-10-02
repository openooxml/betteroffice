import { startTransition, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildRustDisplayList,
  buildRustDisplayFrame,
  applyFrameDelta,
  applyFrameDeltaOwned,
  createCanvasImageResolver,
  createDisplayListQueries,
  endDisplayListQueriesLine,
  decodeFrameDelta,
  decodeFrameDeltaSteps,
  demoDisplayList,
  encodeDisplayListFrameExtras,
  isDisplayListQuerySourceDead,
  RustDisplayListSourceError,
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
  preloadResidentEngineWorker,
  retainPreloadedResidentEngineWorker,
  takePreloadedResidentEngineWorker,
  ResidentWorkerOutOfMemoryError,
  sameYrsSelection,
  type ResidentCaretPaintStyle,
  type ResidentEngineOffscreenPage,
  type ResidentEngineWorkerFrame,
  type ResidentEngineWorkerOpened,
  type ResidentProposalReply,
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
  isLayoutQueued,
  isSupersededLayout,
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
import {
  ensureWorkerOpenReplica,
  failWorkerOpenReplica,
  workerOpenReplicaPending,
  workerOpenSourceVersion,
  type WorkerOpenFallbackReason,
} from '../internals/workerOpenReplica';
import { bindDisplayWindow, type DisplayWindow } from '../internals/displayWindow';
import { sameLayoutInput } from '../internals/layoutInput';
import { SupersededPreviewError } from '../internals/supersededPreview';
import {
  failWorkerProposalAuthority,
  registeredWorkerProposalAuthority,
} from '../internals/workerProposalAuthority';
import { nearestPages } from './pageBuildOrder';
import { scheduleIdlePageBuild, type PageBuildTask } from './pageBuildScheduler';

export interface WorkerOpenedDocument extends ResidentEngineWorkerOpened {
  encodeState(): Promise<Uint8Array>;
  revisionCount(): Promise<number>;
  proposal: ResidentEngineWorkerClient['proposal'];
  documentRead: ResidentEngineWorkerClient['documentRead'];
  handOver: ResidentEngineWorkerClient['handOver'];
  fallback(reason?: WorkerOpenFallbackReason): (() => boolean) | void;
  destroy(): void;
  replicaReady(): void;
  /** Restamps the shown layout after the session started mirroring the worker's version. */
  mirrorReady(): void;
}

const holdsWorkerProposals = (session: YrsSession): boolean =>
  registeredWorkerProposalAuthority(session)?.holdsWorkerState() === true;
const holdsCommittedWorkerProposals = (session: YrsSession): boolean =>
  registeredWorkerProposalAuthority(session)?.holdsCommittedWorkerState() === true;

export type OpenInWorker = (
  session: YrsSession,
  bytes: Uint8Array,
  digest?: string,
  generation?: number
) => Promise<WorkerOpenedDocument | null>;

/** A display-only preview the resident worker opened; see {@link OpenPreviewInWorker}. */
export interface WorkerOpenedPreview {
  hostJson: string;
  /** Resolves once the preview's first layout is queued in the worker, or the worker is gone. */
  bootstrapPosted: Promise<void>;
  /** The preview loaded on this thread: it lays out here from now on, not in the worker. */
  release(): void;
}

/**
 * Opens a display-only preview of the first `blocks` body blocks of `bytes` in the resident
 * worker, which lays it out; null when no worker takes it or the package cannot open as a
 * preview. The {@link OpenInWorker} of the whole document in the same load takes that worker
 * over, so it opens there right behind the preview's queued layout.
 */
export type OpenPreviewInWorker = (
  session: YrsSession,
  bytes: Uint8Array,
  blocks: number
) => Promise<WorkerOpenedPreview | null>;

export type FontRequirementsInWorker = (
  session: YrsSession,
  request: string
) => Promise<string | null> | null;

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
    timeoutMs?: number | null,
    scope?: 'document' | 'window'
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
  openInWorker: OpenInWorker;
  openPreviewInWorker: OpenPreviewInWorker;
  fontRequirementsInWorker: FontRequirementsInWorker;
  /**
   * The pages `[start, end)` near the viewport. Only these are built; every
   * other page arrives as geometry until it comes near.
   */
  setDisplayWindow(start: number, end: number): void;
  setRetainBuiltPages?(retain: boolean): void;
  /** The resident worker's wasm memories as of its latest reply; null without a worker. */
  workerMemory(): WasmModuleMemory[] | null;
  /**
   * True while the worker owns the visible page surfaces. Sticky across
   * invalidation (remote/structural updates) so the canvas keeps its last
   * pixels instead of remounting; drops only on genuine fallback or reset.
   */
  workerSurfacesActive: boolean;
  /** The engine whose provisional layout is shown with the rest not yet asked of the worker. */
  pendingCompletion: YrsSession | null;
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
export type LayoutInWorker = ((
  session: YrsSession,
  request: string
) => Promise<WorkerLayoutComputation | null> | null) & {
  prewarm?: (session: YrsSession) => (() => void) | null;
  /** A live resident worker holds this session's document. */
  ownsDocument?: (session: YrsSession) => boolean;
};

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
  owner: { engine: YrsSession; client: ResidentEngineWorkerClient; load: number };
  /** Content epoch and display extras the worker built the frame for. */
  contentEpoch: number;
  layoutExtras: string;
  /** The frame shows a layout of the first pages only. */
  provisional: boolean;
}

/** The display fallback needs a main-thread layout of a worker-run one. */
class MainThreadLayoutPendingError extends Error {}

/** What the engine reports when asked for a display of a session it never laid out. */
const UNBUILT_PAGINATION = 'resident pagination input is not built';
/** How often a display of a session that has not laid out yet tries again. */
const SESSION_LAYOUT_RETRY_MS = 250;
/** How long a display waits for a session's own layout before failing. */
const SESSION_LAYOUT_WAIT_MS = 5000;

/** The session has not laid out yet; the display tries again shortly. */
class SessionLayoutPendingError extends Error {}

class WorkerPreviewRefusedError extends Error {}

export interface ResidentFrameApplyResult {
  frameEpoch: number | null;
  caretSynchronized: boolean;
  /** Characters a resident deletion removed; absent when unknown. */
  deletedUnits?: number;
}

/** Pages a worker's first frame builds before the viewport is known. */
const INITIAL_DISPLAY_WINDOW: [number, number] = [0, 5];
/** Pages a worker preview's frames build; the full document's frame brings the rest. */
const WORKER_PREVIEW_DISPLAY_WINDOW: [number, number] = [0, 2];
/** Unbuilt pages built per request while the complete list is awaited. */
const SETTLE_BUILD_BATCH_PAGES = 128;
/** Unbuilt pages built per idle period away from the viewport. */
const BACKGROUND_BUILD_BATCH_PAGES = 16;
const WORKER_OPEN_BUILD_MARGIN_PAGES = 2;
const WORKER_OPEN_RETAIN_MARGIN_PAGES = 8;
/** How often a page build waiting behind a newer worker frame checks again. */
const PAGE_BUILD_RETRY_MS = 50;
/** How long a page build waits for the display to adopt a worker frame. */
const UNADOPTED_FRAME_WAIT_MS = 2000;

type PageBuildTimer = ReturnType<typeof setTimeout> | PageBuildTask;
type PageBuildInFlight =
  | { kind: 'release' }
  | { kind: 'build'; background: boolean; cancel(): void; promote(): void };

type DisplayPagesFrame = Pick<ResidentEngineWorkerFrame, 'frame' | 'pageFrames'> &
  Partial<Pick<ResidentEngineWorkerFrame, 'caret' | 'selection' | 'caretPainted' | 'layoutRevision'>>;

function isPageBuildTask(scheduled: PageBuildTimer): scheduled is PageBuildTask {
  return typeof scheduled === 'object' && 'cancel' in scheduled;
}

/** A page build queued on a timer, which runs sooner than an idle-time build. */
function pageBuildTimerQueued(scheduled: PageBuildTimer | null): boolean {
  return scheduled !== null && !isPageBuildTask(scheduled);
}

function cancelPageBuilds(timer: { current: PageBuildTimer | null }): void {
  const scheduled = timer.current;
  timer.current = null;
  if (scheduled === null) return;
  if (isPageBuildTask(scheduled)) scheduled.cancel();
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
  workerOwner?: WorkerLayoutFrame['owner'];
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

function rejectedWorkerLayout(error: unknown): Promise<never> {
  const pending = Promise.reject<never>(error);
  void pending.catch(() => {});
  return pending;
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
  handoffFromRef?: React.RefObject<YrsSession | null>,
  experimentalWorkerOpen = false
): UseRustDisplayListResult {
  const requestLayoutRef = useRef(requestLayout);
  requestLayoutRef.current = requestLayout;
  const engineRef = useRef(engine);
  engineRef.current = engine;
  const unbuiltSinceRef = useRef(new WeakMap<YrsSession, number>());
  const [sessionLayoutRetry, setSessionLayoutRetry] = useState(0);
  const workerHeapLimitRef = useRef(workerHeapLimitBytes);
  workerHeapLimitRef.current = workerHeapLimitBytes;
  // The engine whose layout the shown frame is of: another engine's frames
  // never apply to it as a base.
  const frameEngineRef = useRef<unknown>(null);
  // The worker-open main-engine frame shown, and the content epoch it was built for.
  const mainFrameRef = useRef<{ engine: YrsSession; contentEpoch: number } | null>(null);
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
  const settleWaitersRef = useRef(new Map<() => void, 'document' | 'window'>());
  const settleRelayoutRef = useRef<(() => void) | null>(null);
  const layoutRef = useRef(layout);
  layoutRef.current = layout;
  // The layout shown when another document started loading. What becomes of
  // it settles no wait: waits are for the document being loaded.
  const replacedLayoutRef = useRef<{ layout: Layout | null; line: object } | null>(null);
  const documentLoadsRef = useRef(0);
  // Layouts without a source engine share a line only within a document load.
  const documentLineRef = useRef<object>({});
  const sourceLinesRef = useRef(new WeakMap<object, object>());
  // Lines handed to facades, ended when the session behind them goes away.
  const openLinesRef = useRef(new Set<object>());
  const sourceLine = useCallback((source: RustDisplayListEngine | null | undefined): object => {
    let line = source
      ? sourceLinesRef.current.get(source)
      : (replacedLayoutRef.current?.line ?? documentLineRef.current);
    if (!line) {
      line = {};
      sourceLinesRef.current.set(source!, line);
    }
    openLinesRef.current.add(line);
    return line;
  }, []);
  // `forget: false` keeps the lines listed, so a line revived after a fallback still ends with its session.
  const endOpenLines = useCallback((forget = true): void => {
    for (const line of openLinesRef.current) endDisplayListQueriesLine(line);
    if (forget) openLinesRef.current.clear();
  }, []);
  const markSettled = useCallback(
    (epoch: number | null, failure: Error | null = null, authoritative = false): void => {
      if (replacedLayoutRef.current && !authoritative) return;
      // A queued layout pass may change what shows, such as the revision preview.
      if (!failure && !authoritative && isLayoutQueued(engineRef.current)) return;
      settledEpochRef.current = epoch;
      settleErrorRef.current = failure;
      for (const waiter of [...settleWaitersRef.current.keys()]) waiter();
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
    opened?: boolean;
    stateVector?: Uint8Array;
    proposalFontRequirements?: ResidentProposalReply['fontRequirements'];
    opening?: Promise<ResidentEngineWorkerOpened>;
  } | null>(null);
  const retainBuiltPagesRef = useRef(false);
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
  const workerOpenSourcesRef = useRef(new WeakMap<YrsSession, {
    bytes: Uint8Array;
    digest?: string;
    generation?: number;
    previewBlocks?: number;
  }>());
  // Display-only previews the worker opened and lays out; their load's whole document takes
  // their worker over.
  const workerPreviewEnginesRef = useRef(new WeakSet<YrsSession>());
  // A worker preview's first pass and font requirements: once the whole document took its
  // worker over, a pass of the preview answers with these instead of laying out again.
  const workerPreviewPassesRef = useRef(new WeakMap<YrsSession, Promise<WorkerLayoutComputation | null>>());
  const workerPreviewRequirementsRef = useRef(new WeakMap<YrsSession, Promise<string | null>>());
  const handedOverPreview = useCallback(
    (hostEngine: YrsSession): boolean =>
      workerPreviewEnginesRef.current.has(hostEngine) && handedOverEnginesRef.current.has(hostEngine),
    []
  );
  const workerOpenEnabledRef = useRef(experimentalWorkerOpen);
  workerOpenEnabledRef.current = experimentalWorkerOpen;
  const spawnedWorkerEnginesRef = useRef(new WeakSet<YrsSession>());
  // Sessions whose worker and page surfaces a successor took over: they never take one back.
  const handedOverEnginesRef = useRef(new WeakSet<YrsSession>());
  const workerFallbackEngineRef = useRef<YrsSession | null>(null);
  const recoveredWorkerEpochsRef = useRef(new WeakMap<YrsSession, number>());
  const recoveredEngine = useCallback(
    (engine: YrsSession | null | undefined): boolean =>
      engine != null && recoveredWorkerEpochsRef.current.has(engine),
    []
  );
  const bootstrapFrameEpoch = useCallback(
    (engine: YrsSession): { frameEpoch?: number } =>
      recoveredEngine(engine)
        ? {
            frameEpoch: Math.max(
              recoveredWorkerEpochsRef.current.get(engine)!,
              snapshotRef.current.frame?.frameEpoch ?? 0
            ),
          }
        : followedFrameEpoch(snapshotRef.current.frame),
    [recoveredEngine]
  );
  // No worker starts once the hook is gone, whatever failure arrives late.
  const unmountedRef = useRef(false);
  const isCurrentWorker = useCallback(
    (hostEngine: YrsSession, owner: NonNullable<typeof workerRef.current>): boolean =>
      !unmountedRef.current &&
      workerRef.current === owner &&
      owner.engine === hostEngine &&
      owner.load === documentLoadsRef.current,
    []
  );
  // The engines whose worker was replaced, each with the failure once its
  // replacement failed too. Weak, so a replaced document's session is not kept.
  const outOfMemoryRef = useRef(new WeakMap<YrsSession, Error | null>());
  const workerFailureListenerRef = useRef<(client: ResidentEngineWorkerClient, failure: Error) => void>(() => {});
  const workerFailureRef = useRef(new WeakMap<YrsSession, Error>());
  const displayWindowRef = useRef<[number, number]>(INITIAL_DISPLAY_WINDOW);
  const displayWindowListenersRef = useRef(new Set<() => void>());
  const displayWindow = useMemo<DisplayWindow>(
    () => ({
      read: () => displayWindowRef.current,
      subscribe(listener) {
        displayWindowListenersRef.current.add(listener);
        return () => {
          displayWindowListenersRef.current.delete(listener);
        };
      },
    }),
    []
  );
  const pageBuildInFlightRef = useRef<PageBuildInFlight | null>(
    null
  );
  const pageBuildTimerRef = useRef<PageBuildTimer | null>(null);
  const provisionalPageFrameRef = useRef(false);
  const deferredPageBuildsRef = useRef(new WeakMap<YrsSession, {
    client: ResidentEngineWorkerClient;
    layoutEpoch: number;
    pages: Set<number>;
  }>());
  const schedulePageBuildsWhenIdleRef = useRef<() => void>(() => {});
  const retryPageBuildsRef = useRef<(idle: boolean) => void>(() => {});
  const unadoptedFrameSinceRef = useRef<number | null>(null);
  const workerLayoutFramesRef = useRef(new WeakMap<Layout, WorkerLayoutFrame>());
  const completionGateRef = useRef<(() => void) | null>(null);
  const resolvedCommentIdsRef = useRef(resolvedCommentIds);
  resolvedCommentIdsRef.current = resolvedCommentIds;
  const recoveryFrameEpochRef = useRef(0);
  const workerInputQueueRef = useRef<Promise<void>>(Promise.resolve());
  const suppressWorkerInvalidationRef = useRef(0);
  const [workerSurfacesActive, setWorkerSurfacesActive] = useState(false);
  // One entry per provisional layout, so only its own gate clears it.
  const [pendingCompletion, setPendingCompletion] = useState<{ engine: YrsSession } | null>(null);
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

  const setMainFrameDisplayWindow = useCallback((hostEngine: YrsSession | null): void => {
    if (
      !workerOpenEnabledRef.current ||
      !hostEngine?.setDisplayWindow ||
      hostEngine.isDisplayOnly?.() === true ||
      handedOverEnginesRef.current.has(hostEngine)
    ) return;
    hostEngine.setDisplayWindow(...displayWindowRef.current);
    hostEngine.setDisplayRetainBuiltPages(retainBuiltPagesRef.current);
    hostEngine.setWindowedIncrementalBuilds(true);
  }, []);

  const mainPageBuildEngine = useCallback((): YrsSession | null => {
    const hostEngine = residentEngineRef.current;
    return workerOpenEnabledRef.current &&
      hostEngine?.setDisplayWindow &&
      mainFrameRef.current?.engine === hostEngine &&
      mainFrameRef.current.contentEpoch === contentEpochRef.current &&
      frameEngineRef.current === hostEngine &&
      hostEngine.isDisplayOnly?.() !== true &&
      !workerPreviewEnginesRef.current.has(hostEngine) &&
      !handedOverEnginesRef.current.has(hostEngine)
      ? hostEngine : null;
  }, []);

  const failWorkerDocument = useCallback(
    (
      hostEngine: YrsSession,
      cause: unknown,
      message = '[CanvasRenderer] Resident worker holding proposals failed'
    ): Error => {
      const previous = workerFailureRef.current.get(hostEngine);
      if (previous) return previous;
      const failure = cause instanceof Error
        ? cause
        : new Error(`Resident engine worker failed: ${String(cause)}`);
      if (unmountedRef.current || sessionLoad(hostEngine) !== documentLoadsRef.current) return failure;
      workerFailureRef.current.set(hostEngine, failure);
      failWorkerProposalAuthority(hostEngine, failure);
      failWorkerOpenReplica(hostEngine, failure);
      workerOpenSourcesRef.current.delete(hostEngine);
      if (workerRef.current?.engine === hostEngine) {
        workerRef.current.client.destroy();
        workerRef.current = null;
      }
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      console.error(message, failure);
      queryEpochGate.clear();
      setError(failure);
      setErrorEngine(hostEngine);
      setLoading(false);
      // A display-only preview's failure fails no wait: the full session replaces it.
      if (hostEngine.isDisplayOnly?.() !== true) markSettled(null, failure, true);
      return failure;
    },
    [markSettled, queryEpochGate, sessionLoad, setWorkerPresentationActive]
  );

  const ensureRebuildableReplica = useCallback(
    (hostEngine: YrsSession): void => {
      const failure = workerFailureRef.current.get(hostEngine);
      if (failure) throw failure;
      if (holdsWorkerProposals(hostEngine)) {
        throw failWorkerDocument(
          hostEngine, new Error('The resident worker holds proposals the main thread cannot rebuild')
        );
      }
      ensureWorkerOpenReplica(hostEngine);
    },
    [failWorkerDocument]
  );

  const watchWorkerFailure = (client: ResidentEngineWorkerClient): void =>
    client.onFailure((failure) => workerFailureListenerRef.current(client, failure));

  // The worker for `hostEngine`: its own, the one of the session it
  // takes over from, a spare, or a new one.
  const workerFor = useCallback(
    (hostEngine: YrsSession): NonNullable<typeof workerRef.current> => {
      const current = workerRef.current;
      if (current?.engine === hostEngine) return current;
      const failure = workerFailureRef.current.get(hostEngine);
      if (failure) throw failure;
      if (holdsCommittedWorkerProposals(hostEngine)) {
        throw failWorkerDocument(
          hostEngine, new Error('The resident worker holding this document is gone')
        );
      }
      const load = sessionLoad(hostEngine);
      const replacement = spawnedWorkerEnginesRef.current.has(hostEngine);
      spawnedWorkerEnginesRef.current.add(hostEngine);
      if (current && handoffFromRef?.current === current.engine) {
        handedOverEnginesRef.current.add(current.engine);
        current.client.rebootstrap();
        workerRef.current = { engine: hostEngine, client: current.client, load };
        watchWorkerFailure(current.client);
        return workerRef.current;
      }
      if (
        current &&
        workerPreviewEnginesRef.current.has(current.engine) &&
        !workerPreviewEnginesRef.current.has(hostEngine) &&
        current.load === load &&
        !current.client.hasFailed()
      ) {
        handedOverEnginesRef.current.add(current.engine);
        current.client.rebootstrap();
        workerRef.current = { engine: hostEngine, client: current.client, load };
        workerRef.current.client.setRetainBuiltPages(retainBuiltPagesRef.current);
        watchWorkerFailure(current.client);
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
      workerRef.current.client.setRetainBuiltPages(retainBuiltPagesRef.current);
      watchWorkerFailure(workerRef.current.client);
      return workerRef.current;
    },
    [failWorkerDocument, handoffFromRef, sessionLoad]
  );

  const adoptHostEngine = useCallback(
    (hostEngine: YrsSession): void => {
      if (holdsWorkerProposals(hostEngine)) {
        throw failWorkerDocument(
          hostEngine, new Error('The resident worker holds proposals the main thread cannot rebuild')
        );
      }
      if (workerFallbackEngineRef.current === hostEngine) return;
      setMainFrameDisplayWindow(hostEngine);
      hostEngine.resetFrameBase();
      queryEpochGate.clear();
      recoveryFrameEpochRef.current = snapshotRef.current.frame?.frameEpoch ?? 0;
      const fallbackSnapshot = { ...snapshotRef.current, frame: null, queries: null, caret: null };
      snapshotRef.current = fallbackSnapshot;
      setSnapshot(fallbackSnapshot);
      endOpenLines(false);
      workerFallbackEngineRef.current = hostEngine;
      mainFrameRef.current = null;
    },
    [endOpenLines, failWorkerDocument, queryEpochGate, setMainFrameDisplayWindow]
  );

  useEffect(() => {
    if (!residentEngine) return;
    return residentEngine.onUpdate((update) => {
      if (
        suppressWorkerInvalidationRef.current > 0 ||
        (workerOpenEnabledRef.current && workerOpenReplicaPending(residentEngine))
      ) return;
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
      endOpenLines();
      markSettled(null, new Error('The editor was unmounted'), true);
    };
  }, [endOpenLines, markSettled, queryEpochGate]);

  const publishQuerySnapshot = useCallback(
    (nextSnapshot: RustDisplayListSnapshot, contentEpoch: number): void => {
      const queries = nextSnapshot.queries;
      if (!queries) {
        if (contentEpoch === contentEpochRef.current) queryEpochGate.clear();
        return;
      }
      if (contentEpoch !== contentEpochRef.current) return;
      bindDisplayWindow(queries, displayWindow);
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
    [displayWindow, queryEpochGate]
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

  // A worker that ran out of memory, or with worker-open failed while no
  // request handled it, is replaced by a fresh one once. The main
  // thread never takes over its work: its memory has the same limit and
  // already holds the document. `retry` asks the caller to use the current
  // worker, `stale` means the failed worker no longer serves this engine.
  // A worker of a document that another load replaced fails nothing of the new one.
  const replaceOutOfMemoryWorker = useCallback(
    (
      hostEngine: YrsSession,
      worker: { client: ResidentEngineWorkerClient; load: number } | null,
      failure: Error
    ): 'retry' | 'stale' | 'failed' => {
      const outOfMemory = failure instanceof ResidentWorkerOutOfMemoryError;
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
      if (!previous && !holdsCommittedWorkerProposals(hostEngine)) {
        outOfMemoryRef.current.set(hostEngine, null);
        console.warn(
          outOfMemory
            ? '[CanvasRenderer] Resident engine worker ran out of memory; starting a fresh worker'
            : '[CanvasRenderer] Resident engine worker failed; starting a fresh worker',
          failure
        );
        return 'retry';
      }
      outOfMemoryRef.current.set(hostEngine, failure);
      failWorkerDocument(
        hostEngine,
        failure,
        previous
          ? outOfMemory
            ? '[CanvasRenderer] Resident engine worker ran out of memory again'
            : '[CanvasRenderer] Resident engine worker failed again'
          : outOfMemory
            ? '[CanvasRenderer] Resident engine worker holding proposals ran out of memory'
            : '[CanvasRenderer] Resident engine worker holding proposals failed'
      );
      return 'failed';
    },
    [failWorkerDocument, setWorkerPresentationActive]
  );

  // A failure no request handled by the next task leaves the dead worker current.
  workerFailureListenerRef.current = (client, failure) => {
    if (!workerOpenEnabledRef.current) return;
    setTimeout(() => {
      const owner = workerRef.current;
      if (unmountedRef.current || owner?.client !== client || workerFailureRef.current.has(owner.engine)) return;
      if (replaceOutOfMemoryWorker(owner.engine, owner, failure) !== 'retry') return;
      registeredWorkerProposalAuthority(owner.engine)?.restart();
      requestLayoutRef.current?.();
    }, 0);
  };

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
          UNKNOWN_REVISION_PREVIEW_KEY,
          sourceLine(hostEngine)
        );
        generationRef.current += 1;
        mainFrameRef.current = null;
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
        const line = sourceLine(worker.engine);
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
                  paintCaret,
                  displayWindowRef.current
                )
              : await worker.client.applyDelete(
                  operation.direction,
                  selection,
                  currentFrame.frameEpoch,
                  false,
                  paintCaret,
                  operation.count,
                  displayWindowRef.current
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
          // A worker another document's load replaced rejects the input it held for its own.
          if (dropped() || worker.load !== documentLoadsRef.current) {
            return { frameEpoch: null, caretSynchronized: false };
          }
          if (!isCurrentWorker(worker.engine, worker)) return null;
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
        if (workerRef.current !== worker) {
          return { frameEpoch: null, caretSynchronized: false, deletedUnits: result.deletedUnits };
        }
        const previous = snapshotRef.current;
        if (previous.frame && delta.frameEpoch <= previous.frame.frameEpoch) {
          // Superseded: a newer frame's reply owns the painted-caret verdict.
          return { frameEpoch: null, caretSynchronized: false, deletedUnits: result.deletedUnits };
        }
        const nextFrame = applyFrameDeltaOwned(previous.frame, delta);
        mainFrameRef.current = null;
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
          workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision),
          line
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
      isCurrentWorker,
      paintedCaretMachine,
      publishQuerySnapshot,
      queryEpochGate,
      replaceOutOfMemoryWorker,
      requestSettleRelayout,
      sourceLine,
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
    (hostEngine: YrsSession, cause?: unknown): boolean => {
      if (workerFailureRef.current.has(hostEngine)) return false;
      if (holdsWorkerProposals(hostEngine)) {
        failWorkerDocument(
          hostEngine, cause ?? new Error('The resident worker holds proposals the main thread cannot rebuild')
        );
        return false;
      }
      if (workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine)) {
        ensureRebuildableReplica(hostEngine);
      }
      adoptHostEngine(hostEngine);
      if (workerRef.current?.engine === hostEngine) {
        workerRef.current.client.destroy();
        workerRef.current = null;
      }
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      return true;
    },
    [adoptHostEngine, ensureRebuildableReplica, failWorkerDocument, setWorkerPresentationActive]
  );

  const requestOpenedWorker = useCallback(
    async <T,>(
      hostEngine: YrsSession,
      request: (owner: NonNullable<typeof workerRef.current>) => Promise<T>,
      onOwner?: (owner: NonNullable<typeof workerRef.current>) => void,
    ): Promise<T> => {
      const source = workerOpenSourcesRef.current.get(hostEngine);
      const load = sessionLoad(hostEngine);
      for (;;) {
        const outOfMemory = outOfMemoryRef.current.get(hostEngine);
        const failure = workerFailureRef.current.get(hostEngine);
        if (failure) throw failure;
        if (holdsCommittedWorkerProposals(hostEngine) && workerRef.current?.engine !== hostEngine) {
          throw failWorkerDocument(
            hostEngine, outOfMemory ?? new Error('The resident worker holding this document is gone')
          );
        }
        if (outOfMemory) throw outOfMemory;
        if (
          ((!source || workerOpenSourcesRef.current.get(hostEngine) !== source) &&
            !holdsWorkerProposals(hostEngine)) ||
          unmountedRef.current ||
          load !== documentLoadsRef.current ||
          (workerFallbackEngineRef.current === hostEngine && !holdsWorkerProposals(hostEngine))
        ) {
          throw new SupersededPreviewError();
        }
        if (workerRef.current?.engine !== hostEngine) {
          if (!source || handedOverPreview(hostEngine)) throw new SupersededPreviewError();
          const owner = workerFor(hostEngine);
          owner.opened = true;
          const opening =
            source.previewBlocks === undefined
              ? owner.client.open(source.bytes, {
                  digest: source.digest,
                  ...(source.generation !== undefined ? { generation: String(source.generation) } : {}),
                  heapLimitBytes: workerHeapLimitRef.current,
                })
              : owner.client
                  .openPreview(source.bytes, source.previewBlocks, {
                    heapLimitBytes: workerHeapLimitRef.current,
                  })
                  .then((opened) => {
                    if (!opened) throw new WorkerPreviewRefusedError();
                    return opened;
                  });
          owner.opening = opening.then((opened) => {
            owner.stateVector = opened.stateVector;
            return opened;
          });
        }
        const owner = workerRef.current;
        if (!owner) throw new SupersededPreviewError();
        onOwner?.(owner);
        try {
          await owner.opening;
          if (workerRef.current !== owner) {
            if (
              outOfMemoryRef.current.has(hostEngine) &&
              (!workerRef.current || workerRef.current.engine === hostEngine)
            ) continue;
            throw new SupersededPreviewError();
          }
          if (unmountedRef.current || load !== documentLoadsRef.current) {
            throw new SupersededPreviewError();
          }
          const result = await request(owner);
          if (!isCurrentWorker(hostEngine, owner)) throw new SupersededPreviewError();
          return result;
        } catch (error) {
          const failure = workerFailureRef.current.get(hostEngine);
          if (failure) throw failure;
          if (error instanceof ResidentWorkerOutOfMemoryError) {
            if (replaceOutOfMemoryWorker(hostEngine, owner, error) === 'retry') {
              if (holdsWorkerProposals(hostEngine)) {
                registeredWorkerProposalAuthority(hostEngine)?.restart();
              }
              continue;
            }
          } else {
            if (!isCurrentWorker(hostEngine, owner)) throw new SupersededPreviewError();
            if (owner.client.hasFailed()) {
              if (holdsCommittedWorkerProposals(hostEngine)) {
                throw failWorkerDocument(hostEngine, error);
              }
              if (holdsWorkerProposals(hostEngine)) {
                owner.client.destroy();
                workerRef.current = null;
                setWorkerSurfacesActive(false);
                setWorkerPresentationActive(false);
                registeredWorkerProposalAuthority(hostEngine)?.restart();
              }
            }
          }
          throw error;
        }
      }
    },
    [
      failWorkerDocument,
      handedOverPreview,
      isCurrentWorker,
      replaceOutOfMemoryWorker,
      sessionLoad,
      setWorkerPresentationActive,
      workerFor,
    ]
  );

  const openInWorker = useCallback<OpenInWorker>(
    async (hostEngine, bytes, digest, generation) => {
      if (overrides?.build || !canUseResidentEngineWorker()) return null;
      workerOpenSourcesRef.current.set(hostEngine, { bytes, digest, generation });
      let needsLayout = false;
      const restamp = (): void => {
        if (workerRef.current?.engine !== hostEngine) return;
        const targets = [layoutRef.current, snapshotRef.current.displayList, snapshotRef.current.queries];
        for (const target of targets) {
          if (target) {
            stampSourceVersion(target, workerOpenSourceVersion(hostEngine, sourceVersionOf(target)));
          }
        }
      };
      try {
        const opened = await requestOpenedWorker(hostEngine, (owner) => owner.opening!);
        return {
          ...opened,
          encodeState: () => requestOpenedWorker(hostEngine, (owner) => owner.client.encodeState()),
          revisionCount: () => requestOpenedWorker(hostEngine, (owner) => owner.client.revisionCount()),
          proposal: (op) =>
            requestOpenedWorker(hostEngine, async (owner) => {
              owner.proposalFontRequirements = undefined;
              const reply = await owner.client.proposal(op);
              owner.proposalFontRequirements = reply.fontRequirements;
              return reply;
            }),
          documentRead: (read) =>
            requestOpenedWorker(hostEngine, (owner) => owner.client.documentRead(read)),
          handOver: () =>
            requestOpenedWorker(hostEngine, (owner) => owner.client.handOver()),
          fallback: (reason = 'failure') => {
            const outOfMemory = outOfMemoryRef.current.get(hostEngine);
            if (outOfMemory) throw outOfMemory;
            const owner = workerRef.current;
            if (!owner || !isCurrentWorker(hostEngine, owner)) {
              throw new SupersededPreviewError();
            }
            const recover = workerOpenEnabledRef.current && reason !== 'failure' &&
              !outOfMemoryRef.current.has(hostEngine) &&
              !workerFailureRef.current.has(hostEngine) &&
              !owner.client.hasFailed() && !holdsWorkerProposals(hostEngine);
            if (!dropWorker(hostEngine)) throw workerFailureRef.current.get(hostEngine);
            needsLayout = true;
            const recoveryEpoch = recoveryFrameEpochRef.current;
            if (recover) return () => {
              if (unmountedRef.current || sessionLoad(hostEngine) !== documentLoadsRef.current ||
                workerRef.current || workerFallbackEngineRef.current !== hostEngine ||
                outOfMemoryRef.current.has(hostEngine) || workerFailureRef.current.has(hostEngine)) return false;
              recoveredWorkerEpochsRef.current.set(hostEngine, recoveryEpoch);
              workerOpenSourcesRef.current.delete(hostEngine);
              workerFallbackEngineRef.current = null;
              needsLayout = false;
              return true;
            };
          },
          destroy: () => {
            failWorkerProposalAuthority(hostEngine, new Error('The document changed while opening the replica'));
            workerOpenSourcesRef.current.delete(hostEngine);
            if (workerRef.current?.engine !== hostEngine) return;
            workerRef.current.client.destroy();
            workerRef.current = null;
          },
          replicaReady: () => {
            if (unmountedRef.current || sessionLoad(hostEngine) !== documentLoadsRef.current) return;
            if (
              needsLayout ||
              (workerRef.current?.engine === hostEngine && !workerRef.current.client.bootstrapSent())
            ) {
              const load = documentLoadsRef.current;
              setTimeout(() => {
                if (!unmountedRef.current && load === documentLoadsRef.current) {
                  requestLayoutRef.current?.();
                }
              }, 0);
            }
            restamp();
          },
          mirrorReady: () => {
            if (unmountedRef.current || sessionLoad(hostEngine) !== documentLoadsRef.current) return;
            restamp();
          },
        };
      } catch (error) {
        workerOpenSourcesRef.current.delete(hostEngine);
        const owner = workerRef.current;
        if (owner && isCurrentWorker(hostEngine, owner)) {
          owner.client.destroy();
          workerRef.current = null;
        }
        throw error;
      }
    },
    [dropWorker, isCurrentWorker, overrides?.build, requestOpenedWorker, sessionLoad]
  );

  const openPreviewInWorker = useCallback<OpenPreviewInWorker>(
    async (hostEngine, bytes, blocks) => {
      if (overrides?.build || !canUseResidentEngineWorker()) return null;
      workerOpenSourcesRef.current.set(hostEngine, { bytes, previewBlocks: blocks });
      workerPreviewEnginesRef.current.add(hostEngine);
      let owner: NonNullable<typeof workerRef.current> | null = null;
      try {
        const opened = await requestOpenedWorker(
          hostEngine,
          (current) => current.opening!,
          (current) => {
            owner = current;
          }
        );
        const client = owner!.client;
        return {
          hostJson: opened.hostJson,
          bootstrapPosted: client.whenBootstrapSent(),
          release: () => {
            if (unmountedRef.current || handedOverEnginesRef.current.has(hostEngine)) return;
            if (!dropWorker(hostEngine)) return;
            const load = documentLoadsRef.current;
            setTimeout(() => {
              if (!unmountedRef.current && load === documentLoadsRef.current) requestLayoutRef.current?.();
            }, 0);
          },
        };
      } catch (error) {
        workerOpenSourcesRef.current.delete(hostEngine);
        // A package that cannot open as a preview leaves its worker to the whole document.
        if (error instanceof WorkerPreviewRefusedError) return null;
        workerPreviewEnginesRef.current.delete(hostEngine);
        const current = workerRef.current;
        if (current && isCurrentWorker(hostEngine, current)) {
          current.client.destroy();
          workerRef.current = null;
        }
        return null;
      }
    },
    [dropWorker, isCurrentWorker, overrides?.build, requestOpenedWorker]
  );

  const fontRequirementsInWorker = useCallback<FontRequirementsInWorker>(
    (hostEngine, request) => {
      if (handedOverPreview(hostEngine)) {
        return workerPreviewRequirementsRef.current.get(hostEngine) ?? null;
      }
      if (!holdsWorkerProposals(hostEngine) &&
        (!workerOpenEnabledRef.current || !workerOpenReplicaPending(hostEngine))) return null;
      if (!workerOpenSourcesRef.current.has(hostEngine) && !holdsWorkerProposals(hostEngine)) {
        ensureRebuildableReplica(hostEngine);
        return null;
      }
      const owner = { current: workerRef.current };
      // The worker answers in order, so a reply that arrives after the whole document took
      // the preview's worker over is still the preview's.
      let answered: string | undefined;
      const pending = requestOpenedWorker(
        hostEngine,
        async (current) => {
          answered = await (current.proposalFontRequirements &&
            sameLayoutInput(current.proposalFontRequirements.layoutInput, request)
            ? Promise.resolve(current.proposalFontRequirements.requirementsJson)
            : current.client.fontRequirements(request));
          return answered;
        },
        (current) => { owner.current = current; }
      )
        .then((requirements) => {
          JSON.parse(requirements);
          return requirements;
        })
        .catch((error: unknown) => {
          const failure = workerFailureRef.current.get(hostEngine);
          if (failure) throw failure;
          if (error instanceof ResidentWorkerOutOfMemoryError) throw error;
          if (error instanceof SupersededPreviewError) {
            if (workerFallbackEngineRef.current === hostEngine) return null;
            if (answered !== undefined && handedOverPreview(hostEngine)) return answered;
            throw error;
          }
          if (owner.current && isCurrentWorker(hostEngine, owner.current) && !dropWorker(hostEngine, error)) {
            throw workerFailureRef.current.get(hostEngine);
          }
          return null;
        });
      if (workerPreviewEnginesRef.current.has(hostEngine)) {
        workerPreviewRequirementsRef.current.set(hostEngine, pending);
      }
      return pending;
    },
    [dropWorker, ensureRebuildableReplica, handedOverPreview, isCurrentWorker, requestOpenedWorker]
  );

  const shownFrameEngine = useCallback((): unknown => frameEngineRef.current, []);
  const release = useCallback((): void => {
    // A failed load's failure holds for later waits until the next load.
    replacedLayoutRef.current = {
      layout: null,
      line: replacedLayoutRef.current?.line ?? documentLineRef.current,
    };
    documentLineRef.current = {};
    // A build still running publishes nothing, so no line of the released session comes back.
    generationRef.current += 1;
    cancelPageBuilds(pageBuildTimerRef);
    endOpenLines();
    workerRef.current?.client.destroy();
    workerRef.current = null;
    frameEngineRef.current = null;
    setPresentedEngine(null);
    setWorkerSurfacesActive(false);
    setWorkerPresentationActive(false);
  }, [endOpenLines, setWorkerPresentationActive]);

  /** The frame `engine`'s next frame applies to. */
  const frameBase = useCallback(
    (engine: unknown): RetainedFrame | null =>
      frameEngineRef.current === engine ? snapshotRef.current.frame : null,
    []
  );

  const mainCaretPage = (): number | undefined => {
    const main = mainFrameRef.current;
    if (!main || frameEngineRef.current !== main.engine) return undefined;
    try {
      return main.engine.residentCaretSnapshot().caretRect?.pageIndex;
    } catch {
      return undefined;
    }
  };

  const pagesToRelease = useCallback((frame: RetainedFrame): number[] => {
    if (
      !workerOpenEnabledRef.current ||
      retainBuiltPagesRef.current ||
      settleWaitersRef.current.size > 0
    ) {
      return [];
    }
    const [start, end] = displayWindowRef.current;
    const first = Math.max(0, start - WORKER_OPEN_RETAIN_MARGIN_PAGES);
    const last = Math.min(frame.pages.length, end + WORKER_OPEN_RETAIN_MARGIN_PAGES);
    const caretPage = snapshotRef.current.caret?.caretRect?.pageIndex ?? mainCaretPage();
    const candidates: number[] = [];
    for (let index = 0; index < frame.pages.length; index += 1) {
      if (
        !frame.displayList.pages[index]?.unbuilt &&
        (index < first || index >= last) &&
        index !== caretPage
      ) {
        candidates.push(index);
      }
    }
    return candidates;
  }, []);

  const buildUnbuiltPages = useCallback(
    (idle = false): void => {
      const workerOpen = workerOpenEnabledRef.current;
      pageBuildTimerRef.current = null;
      const mainEngine = mainPageBuildEngine();
      // A main-engine frame of an older content epoch waits for the frame of its relayout.
      if (!mainEngine && mainFrameRef.current?.engine === frameEngineRef.current) return;
      const worker = mainEngine ? null : workerRef.current;
      const targetEngine = worker?.engine ?? mainEngine;
      const frame = snapshotRef.current.frame;
      if (!targetEngine || !frame || (worker && !worker.client.isReady())) return;
      // A worker handed to another session builds its pages once it has
      // presented that session's frame.
      if (frameEngineRef.current !== targetEngine) return;
      if (workerPreviewEnginesRef.current.has(targetEngine)) return;
      if (isLayoutQueued(targetEngine) || isSupersededLayout(layoutRef.current)) {
        retryPageBuildsRef.current(idle);
        return;
      }
      const pages = frame.displayList.pages;
      const [start, end] = displayWindowRef.current;
      const settling = settleWaitersRef.current.size > 0;
      const windowOnly = settling
        ? ![...settleWaitersRef.current.values()].includes('document')
        : workerOpen;
      const first = windowOnly ? Math.max(0, start - WORKER_OPEN_BUILD_MARGIN_PAGES) : 0;
      const last = windowOnly
        ? Math.min(pages.length, end + WORKER_OPEN_BUILD_MARGIN_PAGES)
        : pages.length;
      const deferred = worker ? deferredPageBuildsRef.current.get(worker.engine) : undefined;
      const deferredPages = deferred && deferred.client === worker?.client &&
        deferred.layoutEpoch === frame.layoutEpoch
        ? deferred.pages
        : undefined;
      const unbuilt: number[] = [];
      for (let index = first; index < last; index += 1) {
        if (pages[index]?.unbuilt && !deferredPages?.has(index)) unbuilt.push(index);
      }
      const release = unbuilt.length === 0 ? pagesToRelease(frame) : [];
      if (unbuilt.length === 0 && release.length === 0) return;
      let batch = unbuilt.filter((index) => index >= start && index < end);
      const background = batch.length === 0 && release.length === 0;
      const inFlight = pageBuildInFlightRef.current;
      // Pages the worker has built for a background request come back as a
      // whole-document frame to a request based on the display's older frame.
      const supersedingBackground =
        workerOpen &&
        !background && release.length === 0 && inFlight?.kind === 'build' && inFlight.background;
      if (inFlight) {
        if (inFlight.kind === 'release') return;
        if (settling || (!background && !supersedingBackground)) inFlight.promote();
        if (!supersedingBackground) return;
        inFlight.cancel();
        pageBuildInFlightRef.current = null;
      }
      if (release.length > 0) {
        if (!idle) {
          schedulePageBuildsWhenIdleRef.current();
          return;
        }
        if (
          worker && (
            worker.client.frameRequestPending() ||
            worker.client.answeredFrame() > frame.frameEpoch ||
            paintedCaretMachine.shouldPaint(performance.now())
          )
        ) {
          retryPageBuildsRef.current(true);
          return;
        }
      }
      // Behind a worker frame the display has not adopted, the pages would
      // come back as a whole-document recovery frame: wait for it.
      const framePending = worker?.client.frameRequestPending() ?? false;
      if (
        worker && (framePending ||
          (!supersedingBackground && worker.client.answeredFrame() > frame.frameEpoch))
      ) {
        const now = performance.now();
        if (framePending || unadoptedFrameSinceRef.current === null) {
          unadoptedFrameSinceRef.current = now;
        }
        if (now - unadoptedFrameSinceRef.current < UNADOPTED_FRAME_WAIT_MS) {
          retryPageBuildsRef.current(idle);
          return;
        }
      } else {
        unadoptedFrameSinceRef.current = null;
      }
      if (background) {
        if (provisionalPageFrameRef.current) return;
        if (!idle && !settling) {
          schedulePageBuildsWhenIdleRef.current();
          return;
        }
        batch = nearestPages(
          unbuilt,
          start,
          end,
          settling ? SETTLE_BUILD_BATCH_PAGES : BACKGROUND_BUILD_BATCH_PAGES
        );
      }
      let attachment: PageBuildTask | null = null;
      let promoted = false;
      let promote = (): void => {
        promoted = true;
      };
      const build: PageBuildInFlight = release.length > 0
        ? { kind: 'release' }
        : {
            kind: 'build', background: background && workerOpen,
            cancel: () => attachment?.cancel(), promote: () => promote(),
          };
      pageBuildInFlightRef.current = build;
      const buildBase = frame;
      const dispatchedEpoch = contentEpochRef.current;
      const dispatchedGeneration = generationRef.current;
      const targetCurrent = (): boolean => worker
        ? isCurrentWorker(targetEngine, worker)
        : !unmountedRef.current && mainPageBuildEngine() === targetEngine;
      const current = (): boolean =>
        targetCurrent() &&
        ((worker && !workerOpen) || generationRef.current === dispatchedGeneration) &&
        frameEngineRef.current === targetEngine &&
        (!workerOpen || (
          pageBuildInFlightRef.current === build &&
          ((worker && build.kind === 'release') ||
            (contentEpochRef.current === dispatchedEpoch && !worker?.client.frameRequestPending())) &&
          (!background || snapshotRef.current.frame?.frameEpoch === buildBase.frameEpoch)
        ));
      const finish = (): void => {
        if (pageBuildInFlightRef.current !== build) return;
        pageBuildInFlightRef.current = null;
        const keepTimer =
          !workerOpenEnabledRef.current &&
          settleWaitersRef.current.size === 0 &&
          pageBuildTimerQueued(pageBuildTimerRef.current);
        if (!keepTimer) schedulePageBuildsWhenIdleRef.current();
      };
      const line = sourceLine(targetEngine);
      const paintToken = paintedCaretMachine.token();
      const paintCaret =
        Boolean(worker && workerPresentationActiveRef.current &&
          paintedCaretMachine.shouldPaint(performance.now()));
      const failed = (cause: unknown): void => {
        if (!targetCurrent()) return;
        if (!worker) {
          if (
            generationRef.current !== dispatchedGeneration ||
            contentEpochRef.current !== dispatchedEpoch
          ) return;
          const nextError = cause instanceof Error
            ? cause : new Error(`Display-list build failed: ${String(cause)}`);
          console.error('[CanvasRenderer] Building display pages failed', nextError);
          cancelPageBuilds(pageBuildTimerRef);
          queryEpochGate.clear();
          setError(nextError);
          setErrorEngine(targetEngine);
          setLoading(false);
          markSettled(null, nextError);
          return;
        }
        if (cause instanceof ResidentWorkerOutOfMemoryError) {
          if (replaceOutOfMemoryWorker(worker.engine, worker, cause) === 'retry') {
            requestLayoutRef.current?.();
          }
          return;
        }
        if (!dropWorker(worker.engine, cause)) return;
        console.error(
          '[CanvasRenderer] Building display pages failed; falling back to the main-thread engine',
          cause
        );
        requestLayoutRef.current?.();
      };
      const request: Promise<DisplayPagesFrame | { superseded: true } | null> = worker
        ? release.length > 0
          ? worker.client.releasePages(
              release.map((index) => ({ index, pageId: frame.pages[index]!.pageId.toString() })),
              frame.frameEpoch,
              false
            )
          : worker.client.buildPages(batch, frame.frameEpoch, paintCaret, background && workerOpen)
        : (async () => {
            setMainFrameDisplayWindow(targetEngine);
            const bytes = release.length > 0
              ? targetEngine.releaseDisplayPagesFrame(release, frame.frameEpoch)
              : targetEngine.buildDisplayPagesFrame(batch, frame.frameEpoch);
            return bytes === null ? null : { frame: bytes };
          })();
      void request.then(
        (result) => {
          if (!result || 'superseded' in result || !current()) {
            finish();
            return;
          }
          const attach = (nextFrame: RetainedFrame): void => {
            if (!current()) {
              finish();
              return;
            }
            try {
              const previous = snapshotRef.current;
              if (previous.frame && nextFrame.frameEpoch <= previous.frame.frameEpoch) {
                finish();
                return;
              }
              const caret = worker && result.caret
                ? residentCaretForSelection(
                    result.caret,
                    result.selection ?? null,
                    targetEngine.selection(),
                    nextFrame
                  )
                : null;
              const nextSnapshot =
                contentEpochRef.current === dispatchedEpoch
                  ? createRustDisplayListSnapshot(
                      nextFrame.displayList,
                      nextFrame,
                      caret,
                      worker ? null : targetEngine,
                      previous,
                      sourceVersionOf(previous.queries),
                      worker
                        ? workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision!)
                        : revisionPreviewKeyOf(previous.queries),
                      line
                    )
                  : { displayList: nextFrame.displayList, frame: nextFrame, queries: null, caret };
              if (worker) {
                const deferred = deferredPageBuildsRef.current.get(worker.engine);
                const deferredPages = deferred?.client === worker.client &&
                  deferred.layoutEpoch === nextFrame.layoutEpoch
                  ? deferred.pages
                  : new Set<number>();
                for (const index of batch) {
                  if (nextFrame.displayList.pages[index]?.unbuilt) deferredPages.add(index);
                }
                deferredPageBuildsRef.current.set(worker.engine, {
                  client: worker.client, layoutEpoch: nextFrame.layoutEpoch, pages: deferredPages,
                });
              }
              snapshotRef.current = nextSnapshot;
              publishQuerySnapshot(nextSnapshot, contentEpochRef.current);
              if (background && workerOpen) startTransition(() => setSnapshot(nextSnapshot));
              else setSnapshot(nextSnapshot);
              if (worker) applyPaintedCaretReply(Boolean(result.caretPainted && caret?.caretRect), paintToken);
            } catch (error) {
              finish();
              failed(error);
              return;
            }
            finish();
            for (const waiter of [...settleWaitersRef.current.keys()]) waiter();
          };
          if (!background || !workerOpen) {
            try {
              const previous = snapshotRef.current.frame;
              const delta = decodeFrameDelta(result.frame);
              if (previous && delta.frameEpoch <= previous.frameEpoch) {
                finish();
                return;
              }
              attach(applyFrameDeltaOwned(previous, delta));
            } catch (error) {
              finish();
              failed(error);
            }
            return;
          }
          const pageFrames = result.pageFrames ?? [result.frame];
          function* decodePages(): Generator<void, RetainedFrame> {
            let nextFrame = buildBase;
            const damagedPageIds = new Set<bigint>();
            const removedPageIds = new Set<bigint>();
            for (const bytes of pageFrames) {
              const delta = yield* decodeFrameDeltaSteps(bytes);
              nextFrame = applyFrameDelta(nextFrame, delta);
              for (const id of nextFrame.damagedPageIds) damagedPageIds.add(id);
              for (const id of nextFrame.removedPageIds) removedPageIds.add(id);
              yield;
            }
            return { ...nextFrame, damagedPageIds, removedPageIds };
          }
          const steps = decodePages();
          const decode = (deadline: Pick<IdleDeadline, 'timeRemaining'>): void => {
            if (!current()) {
              finish();
              return;
            }
            try {
              while (deadline.timeRemaining() > 0) {
                const step = steps.next();
                if (step.done) {
                  attach(step.value);
                  return;
                }
              }
              attachDecode();
            } catch (error) {
              finish();
              failed(error);
            }
          };
          let urgent = false;
          const attachDecode = (): void => {
            urgent = promoted || settleWaitersRef.current.size > 0;
            attachment = scheduleIdlePageBuild(decode, urgent);
          };
          promote = () => {
            promoted = true;
            if (urgent || !attachment) return;
            attachment.cancel();
            attachDecode();
          };
          attachDecode();
        },
        (error) => {
          finish();
          failed(error);
        }
      );
    },
    [
      applyPaintedCaretReply,
      dropWorker,
      isCurrentWorker,
      mainPageBuildEngine,
      markSettled,
      paintedCaretMachine,
      pagesToRelease,
      publishQuerySnapshot,
      queryEpochGate,
      replaceOutOfMemoryWorker,
      setMainFrameDisplayWindow,
      sourceLine,
    ]
  );

  const schedulePageBuilds = useCallback(
    (delay: number): void => {
      cancelPageBuilds(pageBuildTimerRef);
      pageBuildTimerRef.current = setTimeout(() => buildUnbuiltPages(), delay);
    },
    [buildUnbuiltPages]
  );
  retryPageBuildsRef.current = (idle) => {
    cancelPageBuilds(pageBuildTimerRef);
    pageBuildTimerRef.current = setTimeout(() => {
      if (idle) schedulePageBuildsWhenIdleRef.current();
      else buildUnbuiltPages();
    }, PAGE_BUILD_RETRY_MS);
  };
  schedulePageBuildsWhenIdleRef.current = () => {
    cancelPageBuilds(pageBuildTimerRef);
    pageBuildTimerRef.current = scheduleIdlePageBuild((deadline) => {
      if (deadline.timeRemaining() > 0) buildUnbuiltPages(true);
      else schedulePageBuildsWhenIdleRef.current();
    }, settleWaitersRef.current.size > 0);
  };

  const setDisplayWindow = useCallback(
    (start: number, end: number): void => {
      const current = displayWindowRef.current;
      if (current[0] === start && current[1] === end) return;
      displayWindowRef.current = [start, end];
      setMainFrameDisplayWindow(mainPageBuildEngine());
      schedulePageBuilds(0);
      // A window wait may already be met by the pages the new window shows.
      for (const waiter of [...settleWaitersRef.current.keys()]) waiter();
      for (const listener of [...displayWindowListenersRef.current]) listener();
    },
    [mainPageBuildEngine, schedulePageBuilds, setMainFrameDisplayWindow]
  );

  const setRetainBuiltPages = useCallback(
    (retain: boolean): void => {
      retainBuiltPagesRef.current = retain;
      workerRef.current?.client.setRetainBuiltPages(retain);
      setMainFrameDisplayWindow(mainPageBuildEngine());
      const frame = snapshotRef.current.frame;
      if (!retain && frame && pagesToRelease(frame).length > 0) {
        schedulePageBuildsWhenIdleRef.current();
      }
    },
    [mainPageBuildEngine, pagesToRelease, setMainFrameDisplayWindow]
  );

  useEffect(() => {
    const frame = snapshot.frame;
    if (!frame) return;
    if (frame.displayList.pages.some((page) => page.unbuilt)) {
      schedulePageBuilds(pageBuildInFlightRef.current ? 50 : 16);
    } else if (pagesToRelease(frame).length > 0) {
      schedulePageBuildsWhenIdleRef.current();
    }
  }, [pagesToRelease, schedulePageBuilds, snapshot.frame]);

  useEffect(() => () => {
    cancelPageBuilds(pageBuildTimerRef);
    if (pageBuildInFlightRef.current?.kind === 'build') pageBuildInFlightRef.current.cancel();
    pageBuildInFlightRef.current = null;
  }, []);

  const canLayoutInWorker = useCallback(
    (hostEngine: YrsSession): boolean =>
      !(
        overrides?.build ||
        !canUseResidentEngineWorker() ||
        !isWorkerHostEngine(hostEngine) ||
        !hostEngine.adoptResidentWorkerLayout ||
        (workerOpenEnabledRef.current &&
          hostEngine.isDisplayOnly?.() === true &&
          !workerPreviewEnginesRef.current.has(hostEngine)) ||
        workerFallbackEngineRef.current === hostEngine ||
        handedOverEnginesRef.current.has(hostEngine)
      ),
    [overrides?.build]
  );
  const ownsDocument = useCallback(
    (hostEngine: YrsSession): boolean => {
      const owner = workerRef.current;
      return workerOpenEnabledRef.current &&
        owner !== null &&
        isCurrentWorker(hostEngine, owner) &&
        !owner.client.hasFailed() &&
        // A successful bootstrap also retains documents supplied by sync.
        ((owner.opened === true && owner.stateVector !== undefined) ||
          (owner.client.bootstrapSent() && owner.client.isReady())) &&
        canLayoutInWorker(hostEngine);
    },
    [canLayoutInWorker, isCurrentWorker]
  );
  const prewarmLayoutWorker = useCallback(
    (hostEngine: YrsSession): (() => void) | null => {
      const current = workerRef.current;
      if (
        !canLayoutInWorker(hostEngine) ||
        outOfMemoryRef.current.get(hostEngine) ||
        spawnedWorkerEnginesRef.current.has(hostEngine) ||
        workerOpenSourcesRef.current.has(hostEngine) ||
        current?.engine === hostEngine ||
        (current && handoffFromRef?.current === current.engine)
      ) return null;
      void preloadResidentEngineWorker().catch(() => {});
      return retainPreloadedResidentEngineWorker();
    },
    [canLayoutInWorker, handoffFromRef]
  );

  const layoutInWorker: LayoutInWorker = useCallback<LayoutInWorker>(
    (hostEngine, request) => {
      if (handedOverPreview(hostEngine)) return workerPreviewPassesRef.current.get(hostEngine) ?? null;
      if (!canLayoutInWorker(hostEngine) || !hostEngine.adoptResidentWorkerLayout) {
        if (holdsWorkerProposals(hostEngine)) {
          workerFor(hostEngine);
          throw failWorkerDocument(
            hostEngine,
            new Error('The resident worker holding proposals cannot lay out the document')
          );
        }
        if (workerOpenEnabledRef.current) ensureRebuildableReplica(hostEngine);
        return null;
      }
      const outOfMemory = outOfMemoryRef.current.get(hostEngine);
      if (outOfMemory) return rejectedWorkerLayout(outOfMemory);
      if (
        workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine) &&
        workerRef.current?.engine !== hostEngine &&
        workerOpenSourcesRef.current.has(hostEngine)
      ) {
        const owner = { current: workerRef.current };
        return requestOpenedWorker(hostEngine, async () => null, (current) => {
          owner.current = current;
        })
          .then(() => layoutInWorkerRef.current?.(hostEngine, request) ?? null)
          .catch((error: unknown) => {
            const failure = workerFailureRef.current.get(hostEngine);
            if (failure) throw failure;
            if (error instanceof ResidentWorkerOutOfMemoryError) throw error;
            if (error instanceof SupersededPreviewError) return null;
            if (owner.current && isCurrentWorker(hostEngine, owner.current) && !dropWorker(hostEngine, error)) {
              throw workerFailureRef.current.get(hostEngine);
            }
            return null;
          });
      }
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
      if (!snapshot) {
        if (holdsWorkerProposals(hostEngine)) {
          throw failWorkerDocument(
            hostEngine,
            new Error('The resident worker holding proposals cannot lay out the document')
          );
        }
        if (workerOpenEnabledRef.current) ensureRebuildableReplica(hostEngine);
        return null;
      }
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
        stateVector: workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine)
          ? owner.stateVector
          : hostEngine.encodeStateVector(),
        ...(bootstrapping && owner.opened ? { opened: true } : {}),
        displayWindow: workerPreviewEnginesRef.current.has(hostEngine)
          ? WORKER_PREVIEW_DISPLAY_WINDOW
          : displayWindowRef.current,
        ...(bootstrapping
          ? {
              provisionalPages: PROVISIONAL_LAYOUT_PAGES,
              heapLimitBytes: workerHeapLimitRef.current,
              ...bootstrapFrameEpoch(hostEngine),
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
        if (recoveredEngine(hostEngine) &&
          !isCurrentWorker(hostEngine, owner)) return rejectedWorkerLayout(new SupersededPreviewError());
        const current = workerRef.current;
        if (
          unmountedRef.current ||
          owner.load !== documentLoadsRef.current ||
          (current && current.engine !== hostEngine)
        ) {
          return null;
        }
        const failure = workerFailureRef.current.get(hostEngine);
        if (failure) return rejectedWorkerLayout(failure);
        if (holdsWorkerProposals(hostEngine) && cause instanceof SupersededPreviewError) return null;
        // A session whose replacement worker ran out of memory too lays out nowhere.
        const outOfMemory = cause instanceof ResidentWorkerOutOfMemoryError;
        if (outOfMemory && outOfMemoryRef.current.get(hostEngine)) return rejectedWorkerLayout(cause);
        // A pass of a session no worker serves any more starts no worker.
        if (!current) {
          if (holdsWorkerProposals(hostEngine)) {
            return rejectedWorkerLayout(failWorkerDocument(hostEngine, cause));
          }
          return null;
        }
        if (outOfMemory) {
          // A newer layout, here or in a worker, replaced this pass: the host
          // drops it, and a newer worker request recovers the worker it asks.
          if (hostEngine.residentWorkerProbe()?.layoutRevision !== adoptedRevision) return null;
          const outcome = replaceOutOfMemoryWorker(hostEngine, owner, cause);
          if (outcome === 'failed') return rejectedWorkerLayout(cause);
          if (outcome === 'stale') return null;
          return layoutInWorkerRef.current?.(hostEngine, request) ?? null;
        }
        if (!isCurrentWorker(hostEngine, owner)) return null;
        if (!dropWorker(hostEngine, cause)) {
          return rejectedWorkerLayout(workerFailureRef.current.get(hostEngine)!);
        }
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; laying out on the main thread',
          cause
        );
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
          owner,
          contentEpoch,
          layoutExtras: options.layoutExtras,
          provisional: result.layoutProvisional === true,
        });
        return computation;
      };
      const pass = reply
        .then((result): WorkerLayoutComputation | null => {
          if (recoveredEngine(hostEngine) && !isCurrentWorker(hostEngine, owner)) {
            throw new SupersededPreviewError();
          }
          if (holdsWorkerProposals(hostEngine) &&
            (!isCurrentWorker(hostEngine, owner) ||
              hostEngine.residentWorkerProbe()?.layoutRevision !== adoptedRevision)) return null;
          const computation = adopt(result, previousFrame);
          // A display-only preview is replaced by the full document before
          // anything needs the rest of its pages.
          if (!result.layoutProvisional || hostEngine.isDisplayOnly?.()) return computation;
          // The rest is laid out against the provisional frame, once its page
          // surfaces are attached: the worker answers in order, so asking
          // sooner would hold back the first paint until it is done.
          const provisionalEpoch = result.caret.frameEpoch;
          const provisionalDocEpoch = decodeFrameDelta(result.frame).docEpoch;
          const isCurrentPass = (): boolean =>
            isCurrentWorker(hostEngine, owner) &&
            hostEngine.residentWorkerProbe()?.layoutRevision === adoptedRevision;
          const gate = { engine: hostEngine };
          if (workerOpenEnabledRef.current) setPendingCompletion(gate);
          const surfaced = new Promise<void>((resolve) => {
            completionGateRef.current = resolve;
            setTimeout(resolve, PROVISIONAL_SURFACE_WAIT_MS);
          });
          const complete = surfaced
            // A worker handed to another session lays out that session now.
            .then(async () => {
              setPendingCompletion((current) => (current === gate ? null : current));
              while (isCurrentPass()) {
                const completed = await worker.completeLayout(
                  provisionalEpoch, false, COMPLETION_SLICE_BLOCKS
                );
                if (!isCurrentPass()) {
                  if (recoveredEngine(hostEngine) &&
                    !isCurrentWorker(hostEngine, owner)) throw new SupersededPreviewError();
                  return null;
                }
                if (completed) {
                  const base = frameBase(hostEngine);
                  return adopt(completed, base?.docEpoch === provisionalDocEpoch ? base : undefined);
                }
                if (!holdsWorkerProposals(hostEngine)) return null;
              }
              if (recoveredEngine(hostEngine) &&
                !isCurrentWorker(hostEngine, owner)) throw new SupersededPreviewError();
              return null;
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
      if (workerPreviewEnginesRef.current.has(hostEngine)) {
        workerPreviewPassesRef.current.set(hostEngine, pass);
      }
      return pass;
    },
    [
      bootstrapFrameEpoch,
      canLayoutInWorker,
      dropWorker,
      handedOverPreview,
      ensureRebuildableReplica,
      failWorkerDocument,
      frameBase,
      frameExtrasInputs,
      isCurrentWorker,
      paintedCaretMachine,
      queryEpochGate,
      recoveredEngine,
      replaceOutOfMemoryWorker,
      requestOpenedWorker,
      workerFor,
    ]
  );
  const layoutInWorkerRef: { current: LayoutInWorker } = useRef<LayoutInWorker>(layoutInWorker);
  layoutInWorkerRef.current = layoutInWorker;
  const prewarmableLayoutInWorker = useMemo(
    () => Object.assign(layoutInWorker, { prewarm: prewarmLayoutWorker, ownsDocument }),
    [layoutInWorker, prewarmLayoutWorker, ownsDocument]
  );

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
    if (residentEngine && workerFailureRef.current.has(residentEngine)) return;
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
      endOpenLines();
      documentLineRef.current = {};
      frameEngineRef.current = null;
      recoveryFrameEpochRef.current = 0;
      setSnapshot(EMPTY_DISPLAY_LIST_SNAPSHOT);
      setError(null);
      setLoading(true);
      settledEpochRef.current = null;
      // A failure of the document being loaded holds until its layout or the next load.
      if (!replacedLayoutRef.current) settleErrorRef.current = null;
      // The worker of a session the renderer let go of never builds again; one
      // opening the current load's document has no layout to show yet.
      const current = workerRef.current;
      if (
        current &&
        current.engine !== residentEngine &&
        !(
          current.opened &&
          current.load === documentLoadsRef.current &&
          workerOpenSourcesRef.current.has(current.engine)
        )
      ) {
        current.client.destroy();
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
    const line = sourceLine(engine);
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
    const previewOnMainThread =
      workerOpenEnabledRef.current &&
      residentEngine?.isDisplayOnly?.() === true &&
      !workerPreviewEnginesRef.current.has(residentEngine);
    const workerEligible =
      residentEngine !== null &&
      workerFallbackEngineRef.current !== residentEngine &&
      !previewOnMainThread;
    // Cheap probe only: the full snapshot (document state, font bytes) is
    // built lazily below, and only for bootstrap/sync — steady-state frame
    // builds never encode state or copy fonts.
    const probe = workerEligible ? residentEngine.residentWorkerProbe() : null;
    const buildOnMainThread = () => {
      if (residentEngine && holdsWorkerProposals(residentEngine)) {
        try {
          workerFor(residentEngine);
          return Promise.reject(new MainThreadLayoutPendingError());
        } catch (error) {
          return Promise.reject(error);
        }
      }
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
      try {
        setMainFrameDisplayWindow(residentEngine);
      } catch (error) {
        return Promise.reject(error);
      }
      // A frame engine paints the pagination it retains, which a newer layout may have replaced.
      const retainedRevision = residentEngine?.residentWorkerProbe()?.layoutRevision;
      const base = frameBase(residentEngine ?? engine ?? null);
      return buildRustDisplayFrame(
        buildInputs,
        engine ?? undefined,
        base,
        base?.frameEpoch ?? recoveryFrameEpochRef.current
      ).then(
        (result) => ({
          ...result,
          caret: null as YrsResidentCaretSnapshot | null,
          // The preview's session is retired after the handover; its retained
          // queries answer from their own pages, as a worker preview's do.
          queryEngine: previewOnMainThread ? null : engine,
          workerProduced: false,
          caretPainted: false,
          ...(residentEngine && result.frame && retainedRevision !== inputs.layoutRevision
            ? { previewKey: UNKNOWN_REVISION_PREVIEW_KEY }
            : {}),
        }),
        (error: unknown) => {
          // A layout published with a session that has laid nothing out yet (it
          // replaced the session mid-pass) is another session's: ask for the
          // session's own, and try again until it has one or the wait runs out.
          if (
            residentEngine &&
            retainedRevision === undefined &&
            error instanceof RustDisplayListSourceError &&
            error.stage === 'build' &&
            error.message.includes(UNBUILT_PAGINATION)
          ) {
            const now = performance.now();
            const since = unbuiltSinceRef.current.get(residentEngine);
            if (since === undefined) {
              unbuiltSinceRef.current.set(residentEngine, now);
              setTimeout(() => requestLayoutRef.current?.(), 0);
            }
            if (since === undefined || now - since < SESSION_LAYOUT_WAIT_MS) {
              setTimeout(() => {
                if (
                  generationRef.current === generation &&
                  documentLoadsRef.current === documentLoad
                ) {
                  setSessionLayoutRetry((retry) => retry + 1);
                }
              }, SESSION_LAYOUT_RETRY_MS);
              throw new SessionLayoutPendingError();
            }
            // Another document's load is under way: this one's failure is stale.
            if (documentLoadsRef.current !== documentLoad) throw new SessionLayoutPendingError();
          }
          throw error;
        }
      );
    };
    const paintToken = paintedCaretMachine.token();
    let pending: Promise<BuiltDisplay>;
    const outOfMemory = residentEngine ? outOfMemoryRef.current.get(residentEngine) : null;
    // A worker preview's first frame, built before the whole document's open took its worker over.
    const handedOverPreviewFrame =
      residentEngine !== null &&
      handedOverEnginesRef.current.has(residentEngine) &&
      workerPreviewEnginesRef.current.has(residentEngine) &&
      workerLayoutFramesRef.current.has(layout);
    if (residentEngine && handedOverEnginesRef.current.has(residentEngine) && !handedOverPreviewFrame) {
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
        if (
          unmountedRef.current ||
          documentLoad !== documentLoadsRef.current ||
          (requested?.engine === hostEngine && !isCurrentWorker(hostEngine, requested))
        ) {
          return Promise.reject(new SupersededPreviewError());
        }
        // A successor that failed to construct leaves this build to the host engine.
        if (requested !== workerRef.current) {
          if (workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine) &&
            !holdsWorkerProposals(hostEngine)) {
            ensureRebuildableReplica(hostEngine);
          }
          return buildOnMainThread();
        }
        const nextError =
          cause instanceof Error
            ? cause
            : new Error(`Resident engine worker failed: ${String(cause)}`);
        if (!dropWorker(hostEngine, nextError)) {
          return Promise.reject(workerFailureRef.current.get(hostEngine));
        }
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
          nextError
        );
        return buildOnMainThread();
      };
      const requestWorkerFrame = (): Promise<BuiltDisplay> => {
        if (
          workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine) &&
          workerRef.current?.engine !== hostEngine &&
          workerOpenSourcesRef.current.has(hostEngine)
        ) {
          return requestOpenedWorker(hostEngine, async () => null, (owner) => { requested = owner; })
            .then(requestWorkerFrame)
            .catch((error: unknown) => {
              if (
                error instanceof ResidentWorkerOutOfMemoryError ||
                error instanceof SupersededPreviewError
              ) {
                throw error;
              }
              return fallback(error);
            });
        }
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
        const sent = () => ({
          stateVector: workerOpenEnabledRef.current && workerOpenReplicaPending(hostEngine)
            ? owner.stateVector
            : hostEngine.encodeStateVector(),
        });
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
              ? worker.buildFrame(
                  extras,
                  previousFrame?.frameEpoch ?? 0,
                  paintCaret,
                  displayWindowRef.current
                )
              : bootstrapping
                ? worker.bootstrap(snapshot, extras, {
                    ...sent(),
                    ...(owner.opened ? { opened: true } : {}),
                    displayWindow: workerPreviewEnginesRef.current.has(hostEngine)
                      ? WORKER_PREVIEW_DISPLAY_WINDOW
                      : displayWindowRef.current,
                    heapLimitBytes: workerHeapLimitRef.current,
                    ...bootstrapFrameEpoch(hostEngine),
                  })
                : worker.sync(snapshot, extras, previousFrame?.frameEpoch ?? 0, paintCaret, {
                    ...sent(),
                    displayWindow: displayWindowRef.current,
                  });
        return workerFrame
          .then((result) => {
            if (recoveredEngine(hostEngine) && !isCurrentWorker(hostEngine, owner)) {
              throw new SupersededPreviewError();
            }
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
              workerOwner: owner,
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
        const delta = prebuilt ? decodeFrameDelta(prebuilt.result.frame) : null;
        const base = frameBase(hostEngine);
        // The frame is adopted only while nothing newer reached the session
        // or the display since the worker built it. A whole frame (the worker
        // sends one when a page build ran first) replaces any older base.
        const appliesTo = !prebuilt
          ? undefined
          : (base?.frameEpoch ?? null) === (prebuilt.previousFrame?.frameEpoch ?? null)
            ? prebuilt.previousFrame
            : delta?.full && (!base || delta.frameEpoch > base.frameEpoch)
              ? base
              : undefined;
        if (
          prebuilt &&
          delta &&
          appliesTo !== undefined &&
          prebuilt.engine === hostEngine &&
          (!recoveredEngine(prebuilt.engine) || prebuilt.owner === workerRef.current || handedOverPreviewFrame) &&
          (workerRef.current?.engine === hostEngine || handedOverPreviewFrame) &&
          prebuilt.contentEpoch === contentEpoch &&
          prebuilt.layoutExtras === JSON.stringify(frameExtrasInputs())
        ) {
          // The worker ran this layout and built its frame in the same pass.
          const { result } = prebuilt;
          const nextFrame = applyFrameDelta(appliesTo, delta);
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
            workerOwner: prebuilt.owner,
            caretPainted: result.caretPainted,
            provisional: prebuilt.provisional,
            previewKey: workerPreviewKey(workerPreviewKeysRef.current, result.layoutRevision),
          });
        } else {
          pending = handedOverPreviewFrame
            ? Promise.reject(new SupersededPreviewError())
            : requestWorkerFrame();
        }
      } catch (error) {
        pending = fallback(error);
      }
    } else {
      pending = buildOnMainThread();
    }
    pending
      .then((result) => {
        if (residentEngine && workerFailureRef.current.has(residentEngine)) return;
        if (recoveredEngine(residentEngine) && result.workerOwner &&
          result.workerOwner !== workerRef.current && !handedOverPreviewFrame) return;
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
          workerOpenEnabledRef.current && residentEngine
            ? workerOpenSourceVersion(residentEngine, sourceVersion)
            : sourceVersion,
          result.previewKey === undefined ? previewKey : result.previewKey,
          line
        );
        snapshotRef.current = nextSnapshot;
        provisionalPageFrameRef.current = result.provisional === true;
        publishQuerySnapshot(nextSnapshot, contentEpoch);
        setSnapshot(nextSnapshot);
        frameEngineRef.current = residentEngine ?? engine ?? null;
        mainFrameRef.current =
          workerOpenEnabledRef.current && residentEngine && result.frame && !result.workerProduced
            ? { engine: residentEngine, contentEpoch }
            : null;
        setPresentedEngine(residentEngine ?? engine ?? null);
        setError(null);
        setLoading(false);
        if (!result.provisional && layout.partial !== true && !isSupersededLayout(layout)) {
          markSettled(contentEpoch);
        }
        const workerProduced = Boolean(
          result.workerProduced && probe && workerRef.current?.client.isReady()
        );
        // A worker preview paints here: the whole document's open runs in that worker next.
        setWorkerSurfacesActive(
          workerProduced && !(residentEngine && workerPreviewEnginesRef.current.has(residentEngine))
        );
        applyPaintedCaretReply(
          Boolean(workerProduced && result.caretPainted && result.caret?.caretRect),
          paintToken
        );
      })
      .catch((error) => {
        if (residentEngine && workerFailureRef.current.has(residentEngine)) return;
        if (error instanceof MainThreadLayoutPendingError) {
          setTimeout(() => requestLayoutRef.current?.(), 0);
          return;
        }
        if (error instanceof SupersededPreviewError || error instanceof SessionLayoutPendingError) {
          return;
        }
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
    bootstrapFrameEpoch,
    ensureRebuildableReplica,
    failWorkerDocument,
    layout,
    sessionLayoutRetry,
    overrides,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    residentEngine,
    setMainFrameDisplayWindow,
    dropWorker,
    endOpenLines,
    frameExtrasInputs,
    isCurrentWorker,
    setWorkerPresentationActive,
    paintedCaretMachine,
    applyPaintedCaretReply,
    notifyCaretInterrupt,
    publishQuerySnapshot,
    queryEpochGate,
    markSettled,
    recoveredEngine,
    replaceOutOfMemoryWorker,
    requestOpenedWorker,
    requestSettleRelayout,
    sessionLoad,
    sourceLine,
    workerFor,
  ]);

  const resetSettled = useCallback(
    (failure: Error | null = null): void => {
      if (!failure) {
        contentEpochRef.current += 1;
        documentLoadsRef.current += 1;
        // Until the next document's layout arrives, what is built is the shown document's.
        replacedLayoutRef.current = {
          layout: layoutRef.current,
          line: replacedLayoutRef.current?.line ?? documentLineRef.current,
        };
        documentLineRef.current = {};
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
    (
      relayout: (() => void) | null,
      timeoutMs: number | null = 15_000,
      scope: 'document' | 'window' = 'document'
    ): Promise<DisplayList> =>
      new Promise<DisplayList>((resolve, reject) => {
        let timer: ReturnType<typeof setTimeout> | undefined;
        const settle = (): boolean => {
          const failure = settleErrorRef.current;
          const displayList = snapshotRef.current.displayList;
          const [start, end] = displayWindowRef.current;
          const unbuilt = displayList?.pages.some(
            (page, index) =>
              page.unbuilt &&
              (scope === 'document' ||
                (index >= start - WORKER_OPEN_BUILD_MARGIN_PAGES &&
                  index < end + WORKER_OPEN_BUILD_MARGIN_PAGES))
          );
          const current =
            displayList !== null &&
            settledEpochRef.current === contentEpochRef.current &&
            !isLayoutQueued(engineRef.current) &&
            !unbuilt;
          if (!failure && !current) {
            if (unbuilt) schedulePageBuilds(0);
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
        settleWaitersRef.current.set(waiter, scope);
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
    layoutInWorker: prewarmableLayoutInWorker,
    presentedEngine,
    shownFrameEngine,
    release,
    openInWorker,
    openPreviewInWorker,
    fontRequirementsInWorker,
    setDisplayWindow,
    setRetainBuiltPages,
    workerMemory,
    workerSurfacesActive,
    pendingCompletion: pendingCompletion?.engine ?? null,
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
 * and queries show; `line`, the document load it lays out.
 */
function createRustDisplayListSnapshot(
  displayList: DisplayList,
  frame: RetainedFrame | null,
  caret: YrsResidentCaretSnapshot | null,
  engine: RustDisplayListEngine | null | undefined,
  previous: RustDisplayListSnapshot,
  sourceVersion: string | null,
  previewKey: string | null,
  line: object
): RustDisplayListSnapshot {
  const residentQueries = residentDisplayListQueryEngine(engine);
  const queries = createDisplayListQueries(displayList, residentQueries, previous.queries, line);
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
  /** True while the worker owns the visible page surfaces. */
  workerSurfacesActive: boolean;
  /** See {@link UseRustDisplayListResult.pendingCompletion}. */
  pendingCompletion: YrsSession | null;
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
  openInWorker: OpenInWorker;
  openPreviewInWorker: OpenPreviewInWorker;
  fontRequirementsInWorker: FontRequirementsInWorker;
  /** The pages `[start, end)` near the viewport, built before the others. */
  setDisplayWindow(start: number, end: number): void;
  setRetainBuiltPages?(retain: boolean): void;
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
  handoffFromRef?: React.RefObject<YrsSession | null>,
  experimentalWorkerOpen = false
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
    openInWorker,
    openPreviewInWorker,
    fontRequirementsInWorker,
    setDisplayWindow,
    setRetainBuiltPages,
    workerMemory,
    workerSurfacesActive,
    pendingCompletion,
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
    handoffFromRef,
    experimentalWorkerOpen
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
    openInWorker,
    openPreviewInWorker,
    fontRequirementsInWorker,
    setDisplayWindow,
    setRetainBuiltPages,
    workerMemory,
    setWorkerPresentationActive,
    workerSurfacesActive,
    pendingCompletion,
    offscreenReplay,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  };
}
