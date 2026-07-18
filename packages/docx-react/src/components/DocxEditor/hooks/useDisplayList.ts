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
  ResidentEngineWorkerClient,
  type ResidentEngineOffscreenPage,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { RustFontChainsProvider } from './useRustMeasurement';

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
  /** Apply a plain-text edit through the resident engine and publish its frame. */
  applyInput(text: string): Promise<boolean>;
  /** Apply one collapsed deletion/paragraph merge through the resident engine. */
  applyDelete(direction: 'backward' | 'forward'): Promise<boolean>;
  /** True once the dedicated worker owns resident frame production. */
  workerActive: boolean;
  attachOffscreenCanvases(
    pages: ResidentEngineOffscreenPage[],
    activePageIds: string[],
    devicePixelRatio: number,
    zoom: number
  ): Promise<boolean>;
}

/** test seam: unit tests inject a fake engine/inputs-resolver instead of the wasm module */
export interface RustDisplayListHookOverrides {
  build?: typeof buildRustDisplayList;
  getInputs?: typeof getLayoutKernelInputs;
}

type ResidentInputOperation =
  | { kind: 'insert'; text: string }
  | { kind: 'delete'; direction: 'backward' | 'forward' };

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
  const [displayList, setDisplayList] = useState<DisplayList | null>(null);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(true);
  const generationRef = useRef(0);
  const frameRef = useRef<RetainedFrame | null>(null);
  const [frame, setFrame] = useState<RetainedFrame | null>(null);
  const workerRef = useRef<{
    engine: YrsSession;
    client: ResidentEngineWorkerClient;
  } | null>(null);
  const workerFallbackEngineRef = useRef<YrsSession | null>(null);
  const workerInputQueueRef = useRef<Promise<void>>(Promise.resolve());
  const suppressWorkerInvalidationRef = useRef(0);
  const [workerActive, setWorkerActive] = useState(false);

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
      setWorkerActive(false);
    });
  }, [residentEngine]);

  useEffect(
    () => () => {
      workerRef.current?.client.destroy();
      workerRef.current = null;
    },
    []
  );

  const applyResidentInput = useCallback(
    (operation: ResidentInputOperation): Promise<boolean> => {
      const run = async (): Promise<boolean> => {
        const worker = workerRef.current;
        const currentFrame = frameRef.current;
        if (!worker || !worker.client.isReady() || !currentFrame) return false;
        const selection = worker.engine.selection();
        if (!selection) return false;
        const result =
          operation.kind === 'insert'
            ? await worker.client.applyInput(
                operation.text,
                selection,
                currentFrame.frameEpoch,
                false
              )
            : await worker.client.applyDelete(
                operation.direction,
                selection,
                currentFrame.frameEpoch,
                false
              );
        if (!result.applied) return false;
        const delta = decodeFrameDelta(result.frame);
        const nextFrame = applyFrameDeltaOwned(currentFrame, delta);
        suppressWorkerInvalidationRef.current += 1;
        try {
          for (const update of result.updates) worker.engine.applyLocalUpdate(update, 'body');
        } finally {
          suppressWorkerInvalidationRef.current -= 1;
        }
        // Supersede an older async compatibility build before publishing the
        // frame produced by the edit transaction.
        generationRef.current += 1;
        frameRef.current = nextFrame;
        setFrame(nextFrame);
        setDisplayList(nextFrame.displayList);
        setError(null);
        setLoading(false);
        return true;
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
        if (nextError.message.includes('resident input state is not ready')) return false;
        console.error('[CanvasRenderer] Resident input failed', nextError);
        setError(nextError);
        // Once invoked, never fall through to the legacy op: a worker failure
        // may have happened after committing the transaction.
        return true;
      });
    },
    []
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
      zoom: number
    ): Promise<boolean> => {
      const worker = workerRef.current?.client;
      if (!worker?.isReady()) return false;
      await worker.attachCanvases(pages, activePageIds, devicePixelRatio, zoom);
      return true;
    },
    []
  );

  useEffect(() => {
    if (!layout) {
      // layout reset (document change) — drop the stale pages
      generationRef.current++;
      setDisplayList(null);
      setError(null);
      setLoading(true);
      frameRef.current = null;
      setFrame(null);
      setWorkerActive(false);
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
    const snapshot = workerEligible ? residentEngine.residentWorkerSnapshot() : null;
    const buildOnMainThread = () =>
      overrides?.build
        ? build(buildInputs, engine ?? undefined).then((displayList) => ({
            displayList,
            frame: null as RetainedFrame | null,
          }))
        : buildRustDisplayFrame(buildInputs, engine ?? undefined, frameRef.current);
    let pending: Promise<{ displayList: DisplayList; frame: RetainedFrame | null }>;
    if (!overrides?.build && snapshot && canUseResidentEngineWorker()) {
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
        setWorkerActive(false);
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
        const workerFrame = bootstrapping
          ? worker.bootstrap(snapshot, extras)
          : worker.layoutRevision() !== snapshot.layoutRevision
            ? worker.sync(snapshot, extras, frameRef.current?.frameEpoch ?? 0)
            : worker.buildFrame(extras, frameRef.current?.frameEpoch ?? 0);
        pending = workerFrame
          .then((result) => {
            const previous = bootstrapping ? null : frameRef.current;
            const delta = decodeFrameDelta(result.frame);
            const nextFrame = applyFrameDelta(previous, delta);
            return { displayList: nextFrame.displayList, frame: nextFrame };
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
        frameRef.current = result.frame;
        setFrame(result.frame);
        setDisplayList(result.displayList);
        setError(null);
        setLoading(false);
        setWorkerActive(Boolean(snapshot && workerRef.current?.client.isReady()));
      })
      .catch((error) => {
        if (generation !== generationRef.current) return;
        const nextError =
          error instanceof Error ? error : new Error(`Display-list build failed: ${String(error)}`);
        console.error('[CanvasRenderer] Rust display-list build failed', nextError);
        setError(nextError);
        setLoading(false);
      });
  }, [layout, overrides, fontChainsProviderRef, resolvedCommentIds, engine, residentEngine]);

  return {
    displayList,
    error,
    loading,
    frame,
    applyInput,
    applyDelete,
    workerActive,
    attachOffscreenCanvases,
  };
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
  const residentQueries =
    engine?.displayHitTestRegionsJson &&
    engine.displayRangeRectsJson &&
    engine.displayRangeRectsRegionJson
      ? (engine as ResidentDisplayListQueryEngine)
      : undefined;
  return useMemo(
    () => (displayList ? createDisplayListQueries(displayList, residentQueries) : null),
    [displayList, residentQueries]
  );
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
  /** host element of the canvas pages, so pointer routing can map client → page-local coords */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Glyph outlines sourced from the same resident font store as measurement. */
  glyphOutlineProvider: GlyphOutlineProvider | null;
  /** One-call ordinary text insertion; false until resident state is ready. */
  applyInput(text: string): Promise<boolean>;
  /** One-call ordinary deletion/merge; false until resident state is ready. */
  applyDelete(direction: 'backward' | 'forward'): Promise<boolean>;
  /** OffscreenCanvas replay bridge; null keeps DOM-canvas replay. */
  offscreenReplay: {
    attach(
      pages: ResidentEngineOffscreenPage[],
      activePageIds: string[],
      devicePixelRatio: number,
      zoom: number
    ): Promise<boolean>;
  } | null;
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
    applyInput,
    applyDelete,
    workerActive,
    attachOffscreenCanvases,
  } = useRustDisplayList(layout, undefined, fontChainsProviderRef, resolvedCommentIds, engine);
  const resolveImage = useMemo(() => createCanvasImageResolver(), []);
  const status: UseCanvasRendererResult['status'] = error
    ? 'error'
    : loading || displayList == null
      ? 'loading'
      : 'ready';
  // Worker-produced deltas update the retained DisplayList geometry cache on
  // the main thread. Resident queries on the main replica are intentionally
  // bypassed because that replica no longer owns the current display list.
  const queries = useDisplayListQueries(displayList, workerActive ? null : engine);
  const [geometryReady, setGeometryReady] = useState(false);
  useEffect(() => {
    if (!queries) {
      setGeometryReady(false);
      return;
    }
    if (queries.isReady()) {
      setGeometryReady(true);
      return;
    }
    let cancelled = false;
    setGeometryReady(false);
    void queries.whenReady().then(
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
  }, [queries]);
  const canvasHostRef = useRef<HTMLDivElement | null>(null);
  // Offscreen replay is the default when the browser supports transferable
  // canvas surfaces. `offscreenReplay=0` is a diagnostic escape hatch; pages
  // containing host-resolved media still select the DOM-canvas fallback in
  // CanvasPagesView.
  const offscreenAllowed = (() => {
    if (typeof window === 'undefined') return false;
    return new URLSearchParams(window.location.search).get('offscreenReplay') !== '0';
  })();
  const offscreenReplay = useMemo(
    () =>
      workerActive &&
      offscreenAllowed &&
      typeof OffscreenCanvas !== 'undefined' &&
      typeof HTMLCanvasElement !== 'undefined' &&
      'transferControlToOffscreen' in HTMLCanvasElement.prototype
        ? { attach: attachOffscreenCanvases }
        : null,
    [attachOffscreenCanvases, offscreenAllowed, workerActive]
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
    queries: geometryReady ? queries : null,
    canvasHostRef,
    glyphOutlineProvider: engine?.outlineGlyphJson ?? null,
    applyInput,
    applyDelete,
    offscreenReplay,
  };
}
