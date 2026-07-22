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
  type DisplayList,
  type DisplayListQueries,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RustDisplayListEngine,
  type RetainedFrame,
  type ResidentDisplayListQueryEngine,
} from '@betteroffice/docx/layout/render';
import { getLayoutKernelInputs } from '@betteroffice/docx/editor';
import {
  canUseResidentEngineWorker,
  residentCaretSnapshotForFrame,
  ResidentEngineWorkerClient,
  sameYrsSelection,
  type ResidentCaretPaintStyle,
  type ResidentEngineOffscreenPage,
  type YrsResidentCaretSnapshot,
  type YrsSelection,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { RustFontChainsProvider } from './useRustMeasurement';
import { displayListNeedsHostImages } from '../canvasPresentation';
import { CARET_PAINT_IDLE_MS, PaintedCaretMachine } from '../paintedCaret';

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
  /** true until the first display list for the current document is ready */
  loading: boolean;
  /** Binary retained-frame state; null on the compatibility JSON path. */
  frame: RetainedFrame | null;
  /** Query facade built from the same display list as `frame`. */
  queries: DisplayListQueries | null;
  /** Worker-computed caret tagged to `frame`. */
  caret: YrsResidentCaretSnapshot | null;
  /** Apply a plain-text edit through the resident engine and publish its frame. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** Apply one collapsed deletion/paragraph merge through the resident engine. */
  applyDelete(direction: 'backward' | 'forward'): Promise<ResidentFrameApplyResult | null>;
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

export interface ResidentFrameApplyResult {
  frameEpoch: number | null;
  caretSynchronized: boolean;
}

/** test seam: unit tests inject a fake engine/inputs-resolver instead of the wasm module */
export interface RustDisplayListHookOverrides {
  build?: typeof buildRustDisplayList;
  getInputs?: typeof getLayoutKernelInputs;
}

type ResidentInputOperation =
  | { kind: 'insert'; text: string }
  | { kind: 'delete'; direction: 'backward' | 'forward' };

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
  engine?: RustDisplayListEngine | null
): UseRustDisplayListResult {
  const [snapshot, setSnapshot] = useState<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const snapshotRef = useRef<RustDisplayListSnapshot>(EMPTY_DISPLAY_LIST_SNAPSHOT);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(true);
  const generationRef = useRef(0);
  const workerRef = useRef<{
    engine: YrsSession;
    client: ResidentEngineWorkerClient;
  } | null>(null);
  const workerFallbackEngineRef = useRef<YrsSession | null>(null);
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

  useEffect(() => {
    if (!residentEngine || !canUseResidentEngineWorker()) return;
    return residentEngine.onUpdate((update) => {
      if (suppressWorkerInvalidationRef.current > 0) return;
      // Update observers fire from inside the wasm transaction. Calling any
      // other EditSession method here would re-enter the borrowed wasm object
      // (wasm-bindgen correctly rejects that unsafe alias). Selection is sent
      // with the next input request after the transaction has returned.
      workerRef.current?.client.invalidate(update, null);
      // Only frame freshness drops here. The worker keeps the page surfaces
      // (workerSurfacesActive) so the canvas retains its pixels until the
      // post-sync frame lands — flipping surfaces would remount every page.
    });
  }, [residentEngine]);

  useEffect(
    () => () => {
      if (paintedCaretIdleTimerRef.current !== null) {
        clearTimeout(paintedCaretIdleTimerRef.current);
      }
      if (dispatchHoldTimerRef.current !== null) {
        clearTimeout(dispatchHoldTimerRef.current);
      }
      workerRef.current?.client.destroy();
      workerRef.current = null;
    },
    []
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

  const applyResidentInput = useCallback(
    (operation: ResidentInputOperation): Promise<ResidentFrameApplyResult | null> => {
      const run = async (): Promise<ResidentFrameApplyResult | null> => {
        const worker = workerRef.current;
        const currentFrame = snapshotRef.current.frame;
        if (!worker || !worker.client.isReady() || !currentFrame) return null;
        const selection = worker.engine.selection();
        if (!selection) return null;
        paintedCaretMachine.noteInput(performance.now());
        const paintCaret = workerPresentationActiveRef.current;
        const paintToken = paintedCaretMachine.token();
        residentPaintInflightRef.current += 1;
        let result;
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
                  paintCaret
                );
        } finally {
          residentPaintInflightRef.current -= 1;
        }
        if (!result.applied) return null;
        const delta = decodeFrameDelta(result.frame);
        suppressWorkerInvalidationRef.current += 1;
        try {
          for (const update of result.updates) worker.engine.applyLocalUpdate(update, 'body');
        } finally {
          suppressWorkerInvalidationRef.current -= 1;
        }
        const previous = snapshotRef.current;
        if (previous.frame && delta.frameEpoch <= previous.frame.frameEpoch) {
          // Superseded: a newer frame's reply owns the painted-caret verdict.
          return { frameEpoch: null, caretSynchronized: false };
        }
        const nextFrame = applyFrameDeltaOwned(previous.frame, delta);
        const caret = residentCaretForSelection(
          result.caret,
          result.selection,
          worker.engine.selection(),
          nextFrame
        );
        const nextSnapshot = createRustDisplayListSnapshot(
          nextFrame.displayList,
          nextFrame,
          caret,
          null,
          previous
        );
        // Supersede an older async compatibility build before publishing the
        // frame produced by the edit transaction.
        generationRef.current += 1;
        snapshotRef.current = nextSnapshot;
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
        applyPaintedCaretReply(Boolean(result.caretPainted && caret?.caretRect), paintToken);
        return {
          frameEpoch: nextFrame.frameEpoch,
          caretSynchronized: Boolean(
            caret?.caretRect &&
              workerPresentationActiveRef.current &&
              !displayListNeedsHostImages(nextFrame.displayList)
          ),
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
        setError(nextError);
        // Once invoked, never fall through to the legacy op: a worker failure
        // may have happened after committing the transaction.
        return { frameEpoch: null, caretSynchronized: false };
      });
    },
    [applyPaintedCaretReply, paintedCaretMachine]
  );

  const applyInput = useCallback(
    (text: string) => applyResidentInput({ kind: 'insert', text }),
    [applyResidentInput]
  );

  const applyDelete = useCallback(
    (direction: 'backward' | 'forward') =>
      applyResidentInput({ kind: 'delete', direction }),
    [applyResidentInput]
  );

  const attachOffscreenCanvases = useCallback(
    async (
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number,
      caretStyle: ResidentCaretPaintStyle
    ): Promise<boolean> => {
      const worker = workerRef.current?.client;
      // Queue even while the worker is mid-invalidation: requests are handled
      // FIFO, so an attach lands after the sync that follows and the worker
      // rasters the newly attached surfaces itself. Refusing here would strand
      // already-transferred canvases (they cannot be re-transferred).
      if (!worker) return false;
      await worker.attachCanvases(pages, activePageIds, devicePixelRatio, zoom, caretStyle);
      return true;
    },
    []
  );

  useEffect(() => {
    if (!layout) {
      // layout reset (document change) — drop the stale pages
      generationRef.current++;
      snapshotRef.current = EMPTY_DISPLAY_LIST_SNAPSHOT;
      setSnapshot(EMPTY_DISPLAY_LIST_SNAPSHOT);
      setError(null);
      setLoading(true);
      setWorkerSurfacesActive(false);
      setWorkerPresentationActive(false);
      notifyCaretInterrupt();
      return;
    }
    const inputs = (overrides?.getInputs ?? getLayoutKernelInputs)(layout);
    const generation = ++generationRef.current;
    if (!inputs) {
      setError(new Error('No display-list inputs were recorded for the current layout.'));
      setLoading(false);
      return;
    }
    const build = overrides?.build ?? buildRustDisplayList;
    // Merged doc-wide font chains from the Rust measure source (when active).
    // A non-empty map activates GlyphRun emission; absent ⇒ TextRunPrimitive.
    const fontChains = fontChainsProviderRef?.current?.();
    const buildInputs = {
      measured: inputs.measured,
      options: inputs.options,
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
    const buildOnMainThread = () =>
      overrides?.build
        ? build(buildInputs, engine ?? undefined).then((displayList) => ({
            displayList,
            frame: null as RetainedFrame | null,
            caret: null as YrsResidentCaretSnapshot | null,
            queryEngine: engine,
            workerProduced: false,
            caretPainted: false,
          }))
        : buildRustDisplayFrame(buildInputs, engine ?? undefined, snapshotRef.current.frame).then(
            (result) => ({
              ...result,
              caret: null as YrsResidentCaretSnapshot | null,
              queryEngine: engine,
              workerProduced: false,
              caretPainted: false,
            })
          );
    const paintToken = paintedCaretMachine.token();
    let pending: Promise<{
      displayList: DisplayList;
      frame: RetainedFrame | null;
      caret: YrsResidentCaretSnapshot | null;
      queryEngine: RustDisplayListEngine | null | undefined;
      workerProduced: boolean;
      caretPainted: boolean;
    }>;
    if (!overrides?.build && probe && canUseResidentEngineWorker()) {
      const hostEngine = residentEngine;
      if (!hostEngine) throw new Error('Resident worker snapshot requires a host engine');
      const fallback = (cause: unknown) => {
        const nextError =
          cause instanceof Error
            ? cause
            : new Error(`Resident engine worker failed: ${String(cause)}`);
        console.error(
          '[CanvasRenderer] Resident engine worker unavailable; falling back to the main-thread engine',
          nextError
        );
        workerFallbackEngineRef.current = hostEngine;
        if (workerRef.current?.engine === hostEngine) {
          workerRef.current.client.destroy();
          workerRef.current = null;
        }
        setWorkerSurfacesActive(false);
        setWorkerPresentationActive(false);
        return buildOnMainThread();
      };
      try {
        if (workerRef.current?.engine !== hostEngine) {
          workerRef.current?.client.destroy();
          workerRef.current = {
            engine: hostEngine,
            client: new ResidentEngineWorkerClient(),
          };
        }
        const worker = workerRef.current.client;
        const extras = encodeDisplayListFrameExtras(buildInputs);
        const bootstrapping = worker.layoutRevision() === 0;
        const previousFrame = bootstrapping ? null : snapshotRef.current.frame;
        // On a fresh client both hints are null, so a bootstrap snapshot is
        // always complete; a sync snapshot ships a state diff and skips font
        // bytes the worker already holds.
        const buildSnapshot = () => {
          const snapshot = hostEngine.residentWorkerSnapshot({
            knownStateVector: worker.remoteStateVector(),
            knownFontsRevision: worker.syncedFontsRevision(),
          });
          if (!snapshot) throw new Error('Resident worker snapshot was not available');
          return snapshot;
        };
        // Structural text input reaches the worker as a sync/buildFrame; keep
        // the painted caret glued to those frames while the typing burst lasts.
        const paintCaret =
          !bootstrapping &&
          workerPresentationActiveRef.current &&
          paintedCaretMachine.shouldPaint(performance.now());
        const workerFrame = bootstrapping
          ? worker.bootstrap(buildSnapshot(), extras)
          : worker.layoutRevision() !== probe.layoutRevision
            ? worker.sync(buildSnapshot(), extras, previousFrame?.frameEpoch ?? 0, paintCaret)
            : worker.buildFrame(extras, previousFrame?.frameEpoch ?? 0, paintCaret);
        pending = workerFrame
          .then((result) => {
            const delta = decodeFrameDelta(result.frame);
            const nextFrame = applyFrameDelta(previousFrame, delta);
            return {
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
            };
          })
          .catch(fallback);
      } catch (error) {
        pending = fallback(error);
      }
    } else {
      pending = buildOnMainThread();
    }
    pending
      .then((result) => {
        if (generation !== generationRef.current) return;
        const nextSnapshot = createRustDisplayListSnapshot(
          result.displayList,
          result.frame,
          result.caret,
          result.queryEngine,
          snapshotRef.current
        );
        snapshotRef.current = nextSnapshot;
        setSnapshot(nextSnapshot);
        setError(null);
        setLoading(false);
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
        if (generation !== generationRef.current) return;
        const nextError =
          error instanceof Error ? error : new Error(`Display-list build failed: ${String(error)}`);
        console.error('[CanvasRenderer] Rust display-list build failed', nextError);
        setError(nextError);
        setLoading(false);
      });
  }, [
    layout,
    overrides,
    fontChainsProviderRef,
    resolvedCommentIds,
    engine,
    residentEngine,
    setWorkerPresentationActive,
    paintedCaretMachine,
    applyPaintedCaretReply,
    notifyCaretInterrupt,
  ]);

  return {
    displayList: snapshot.displayList,
    error,
    loading,
    frame: snapshot.frame,
    queries: snapshot.queries,
    caret: snapshot.caret,
    applyInput,
    applyDelete,
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

function createRustDisplayListSnapshot(
  displayList: DisplayList,
  frame: RetainedFrame | null,
  caret: YrsResidentCaretSnapshot | null,
  engine: RustDisplayListEngine | null | undefined,
  previous: RustDisplayListSnapshot
): RustDisplayListSnapshot {
  const residentQueries = residentDisplayListQueryEngine(engine);
  return {
    displayList,
    frame,
    queries: createDisplayListQueries(displayList, residentQueries, previous.queries),
    caret,
  };
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

function residentDisplayListQueryEngine(
  engine: RustDisplayListEngine | null | undefined
): ResidentDisplayListQueryEngine | undefined {
  return engine?.displayHitTestRegionsJson &&
    engine.displayRangeRectsJson &&
    engine.displayRangeRectsRegionJson
    ? (engine as ResidentDisplayListQueryEngine)
    : undefined;
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
  /** feed PagedEditor's per-pass Layout into the interaction query source */
  onLayoutComputed: (
    layout: Layout | null,
    engine?: (RustDisplayListEngine & { outlineGlyphJson?: GlyphOutlineProvider }) | null
  ) => void;
  /** media resolver for CanvasPagesView */
  resolveImage: ImageResolver;
  /** Rust display-list query facade for adapter interactions. */
  queries: DisplayListQueries | null;
  /** Worker caret from the same atomic renderer snapshot. */
  caret: YrsResidentCaretSnapshot | null;
  /** Whether worker-presented pixels make `caret` authoritative. */
  authoritativeCaretActive: boolean;
  /** host element of the canvas pages, so pointer routing can map client → page-local coords */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Glyph outlines sourced from the same resident font store as measurement. */
  glyphOutlineProvider: GlyphOutlineProvider | null;
  /** One-call ordinary text insertion; false until resident state is ready. */
  applyInput(text: string): Promise<ResidentFrameApplyResult | null>;
  /** One-call ordinary deletion/merge; false until resident state is ready. */
  applyDelete(
    direction: 'backward' | 'forward'
  ): Promise<ResidentFrameApplyResult | null>;
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
  resolvedCommentIds?: ReadonlySet<number>
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
    loading,
    frame,
    queries: snapshotQueries,
    caret,
    applyInput,
    applyDelete,
    workerSurfacesActive,
    workerPresentationActive,
    setWorkerPresentationActive,
    attachOffscreenCanvases,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  } = useRustDisplayList(layout, undefined, fontChainsProviderRef, resolvedCommentIds, engine);
  const resolveImage = useMemo(() => createCanvasImageResolver(), []);
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
    onLayoutComputed,
    resolveImage,
    queries: geometryReady ? snapshotQueries : null,
    caret,
    authoritativeCaretActive,
    canvasHostRef,
    glyphOutlineProvider: engine?.outlineGlyphJson ?? null,
    applyInput,
    applyDelete,
    setWorkerPresentationActive,
    offscreenReplay,
    paintedCaretActive,
    notifyCaretInput,
    notifyCaretInputDispatched,
    notifyCaretInterrupt,
  };
}
