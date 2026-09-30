import {
  memo,
  useCallback,
  useEffect,
  useInsertionEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
  type Ref,
} from 'react';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import {
  bindDisplayPageRegistry,
  DisplayPageRegistry,
  displayPageHoldsMirrorId,
  displayPageRevision,
  presentDisplayPageBackBuffer,
  rasterizeDisplayPageToBackBuffer,
  GlyphCache,
  loadGlyphOutlineProvider,
  displayPageNoteAnchorRevision,
  type DisplayList,
  type DisplayPage,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import type { UseCanvasRendererResult } from './hooks/useDisplayList';
import { CanvasPageMirror } from './CanvasPageMirror';
import { CanvasInteractiveOverlay } from './CanvasInteractiveOverlay';
import type { PageChromeHandle } from './usePageChrome';
import { CanvasA11yLiveRegion, type CanvasA11yLiveRegionProps } from './CanvasA11yLiveRegion';
import { CANVAS_PAGE_GAP_PX, CANVAS_PAGES_PADDING_PX } from '@betteroffice/docx/layout/render';
import { SIDEBAR_DOCUMENT_SHIFT } from '../sidebar/constants';
import { ParseError } from '../DocxEditorHelpers';
import { displayListNeedsHostImages } from './canvasPresentation';
import { CanvasReplayState, presentCanvasReplay, type CanvasReplayPreparation } from './canvasReplay';
import { resolveCaretPaintColor } from './paintedCaret';
import { clearPresented, markPresented, markReplayFailed } from './internals/layoutProvenance';
import { viewportColumnBand } from './internals/viewportBand';
import { DEFAULT_CARET_WIDTH } from './overlays/SelectionOverlay';

// Canvas is the sole visible renderer. The editing/input subtree stays mounted
// independently so hidden input focus and IME state survive initial
// loading and renderer errors.
export function CanvasPagedArea({
  renderer,
  a11y,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  fontFamilies,
  children,
}: {
  renderer: UseCanvasRendererResult;
  /** The CSS family each document font family paints browser text with, where they differ. */
  fontFamilies?: ReadonlyMap<string, string>;
  /** live-region wiring (host notify ref + Yrs session getter) — see CanvasA11yLiveRegion */
  a11y?: Omit<CanvasA11yLiveRegionProps, 'active'>;
  /** shifts the canvas pages left to make room for the comments sidebar, mirroring the DOM painter's viewport transform */
  sidebarOpen?: boolean;
  /** zoom level (1 = 100%); the canvas re-rasters at `zoom * DPR` so text stays crisp */
  zoom?: number;
  /** mounts the focusable content-control (SDT) overlay above each page; off in read-only mode */
  interactive?: boolean;
  children: ReactNode;
}) {
  return (
    <>
      {renderer.status === 'ready' && renderer.displayList ? (
        <CanvasPagesView
          displayList={renderer.displayList}
          frame={renderer.frame}
          resolveImage={renderer.resolveImage}
          hostRef={renderer.canvasHostRef}
          sidebarOpen={sidebarOpen}
          zoom={zoom}
          interactive={interactive}
          glyphOutlineProvider={renderer.glyphOutlineProvider}
          fontFamilies={fontFamilies}
          offscreenReplay={renderer.offscreenReplay}
          onWorkerPresentationChange={renderer.setWorkerPresentationActive}
          onPageWindowChange={renderer.setDisplayWindow}
        />
      ) : renderer.status === 'error' ? (
        <div data-testid="canvas-renderer-error" role="alert" style={{ minHeight: 240 }}>
          <ParseError message={renderer.error?.message ?? 'Canvas renderer failed.'} />
        </div>
      ) : null}
      {children}
      {a11y ? <CanvasA11yLiveRegion active={renderer.status === 'ready'} {...a11y} /> : null}
    </>
  );
}

// Pages within this many pages of the viewport keep live bitmaps; everything
// farther keeps its canvas ELEMENT (stable identity, exact geometry for
// pointer routing/overlays/scroll math) but releases its backing store. Only
// pixels are windowed — never DOM structure.
const PAGE_WINDOW_BUFFER = 2;
// Documents at or below this page count never window — zero behavior change
// for ordinary documents.
const PAGE_WINDOW_MIN_PAGES = 12;
/** Pages built on demand, most recent first, that keep their chrome past the task that built them. */
const ON_DEMAND_CHROME_PAGES = 12;
// A page already mounted stays mounted until it drifts one page beyond the
// mount band, so slow scrolling at a boundary cannot thrash mount/unmount.
const PAGE_WINDOW_HYSTERESIS = 1;
type ChromeKind = 'mirror' | 'overlay';
type ChromeHandles = Partial<Record<ChromeKind, PageChromeHandle>>;

interface PageWindowRange {
  start: number;
  end: number;
}

function nextPageWindow(
  previous: PageWindowRange | null,
  firstVisible: number,
  lastVisible: number,
  totalPages: number
): PageWindowRange {
  const mountStart = Math.max(0, firstVisible - PAGE_WINDOW_BUFFER);
  const mountEnd = Math.min(totalPages - 1, lastVisible + PAGE_WINDOW_BUFFER);
  if (!previous) return { start: mountStart, end: mountEnd };
  const keepStart = Math.max(0, mountStart - PAGE_WINDOW_HYSTERESIS);
  const keepEnd = Math.min(totalPages - 1, mountEnd + PAGE_WINDOW_HYSTERESIS);
  const start = Math.min(mountStart, Math.max(previous.start, keepStart));
  const end = Math.max(mountEnd, Math.min(previous.end, keepEnd));
  if (start === previous.start && end === previous.end) return previous;
  return { start, end };
}

/**
 * One page's surface: canvas + a11y mirror + optional interactive overlay.
 * Memoized so a keystroke's snapshot commit re-renders only the pages whose
 * `DisplayPage` identity actually changed — the owned frame-delta path keeps
 * untouched pages' identity stable across keystrokes. A page without
 * `chrome` keeps its sized canvas and empty chrome hosts, which
 * `registerChrome` can fill on demand.
 */
const CanvasPageSurface = memo(function CanvasPageSurface({
  page,
  noteAnchorRevision,
  pageKey,
  zoom,
  interactive,
  chrome,
  inWindow,
  deferChrome,
  registerCanvas,
  registerChrome,
}: {
  page: DisplayPage;
  /**
   * `displayPageRevision(page)`: re-renders the surface when an owned delta
   * changes the page in place.
   */
  revision: number;
  noteAnchorRevision: number;
  pageKey: string;
  zoom: number;
  interactive: boolean;
  chrome: boolean;
  inWindow: boolean;
  deferChrome: boolean;
  registerCanvas: (pageKey: string, el: HTMLCanvasElement | null) => void;
  registerChrome: (pageKey: string, kind: ChromeKind, handle: PageChromeHandle | null) => void;
}) {
  const registerMirror = useCallback(
    (handle: PageChromeHandle | null) => registerChrome(pageKey, 'mirror', handle),
    [pageKey, registerChrome]
  );
  const registerOverlay = useCallback(
    (handle: PageChromeHandle | null) => registerChrome(pageKey, 'overlay', handle),
    [pageKey, registerChrome]
  );
  return (
    <div
      className="canvas-page"
      data-page-index={page.pageIndex}
      data-page-key={pageKey}
      style={{ position: 'relative', width: page.width * zoom, height: page.height * zoom }}
    >
      <canvas
        ref={(el) => registerCanvas(pageKey, el)}
        data-page-index={page.pageIndex}
        style={{
          display: 'block',
          width: page.width * zoom,
          height: page.height * zoom,
          background: '#ffffff',
          boxShadow: '0 1px 3px var(--doc-shadow)',
        }}
      />
      <CanvasPageMirror
        page={page}
        zoom={zoom}
        active={chrome}
        defer={deferChrome}
        visible={inWindow}
        register={registerMirror}
        noteAnchorRevision={noteAnchorRevision}
      />
      {interactive ? (
        <CanvasInteractiveOverlay
          page={page}
          zoom={zoom}
          active={chrome}
          defer={deferChrome}
          register={registerOverlay}
        />
      ) : null}
    </div>
  );
});

export function CanvasPagesView({
  displayList,
  frame,
  resolveImage,
  hostRef,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  glyphOutlineProvider,
  fontFamilies,
  offscreenReplay,
  onWorkerPresentationChange,
  onPageWindowChange,
}: {
  displayList: DisplayList;
  /** Binary retained-frame metadata used to scope page replay. */
  frame?: RetainedFrame | null;
  resolveImage?: ImageResolver;
  /** pointer routing maps client coords → page-local through this host element */
  hostRef?: Ref<HTMLDivElement>;
  /** shift the page column left to reserve room for the comments sidebar */
  sidebarOpen?: boolean;
  /**
   * Zoom level (1 = 100%). Instead of CSS-scaling the page column (which would
   * blur the rastered text), each page canvas is re-sized and re-drawn at
   * `zoom * devicePixelRatio` so glyph outlines stay crisp — see
   * `sizeCanvasForPage`. The a11y mirror is CSS-scaled to keep its nodes 1:1
   * over the enlarged canvas.
   */
  zoom?: number;
  /**
   * Mounts the interactive content-control overlay (focusable SDT widgets)
   * above each page. The a11y mirror stays pointer-inert; this separate layer
   * owns the only clickable/focusable SDT controls on the canvas path.
   */
  interactive?: boolean;
  /** Outline source sharing the display engine's resident font store. */
  glyphOutlineProvider?: GlyphOutlineProvider | null;
  /** The CSS family each document font family paints browser text with, where they differ. */
  fontFamilies?: ReadonlyMap<string, string>;
  /** Dedicated worker replay surface; unsupported/media-heavy pages use DOM canvas. */
  offscreenReplay?: UseCanvasRendererResult['offscreenReplay'];
  onWorkerPresentationChange?: (active: boolean) => void;
  /** The pages `[start, end)` that hold bitmaps, reported as the viewport moves. */
  onPageWindowChange?: (start: number, end: number) => void;
}) {
  const canvasesRef = useRef(new Map<string, HTMLCanvasElement>());
  // Page lookups (pointer, overlays, caret) read this instead of searching
  // the host, which also holds every page's accessibility mirror.
  const [pageRegistry] = useState(() => new DisplayPageRegistry());
  const registerCanvas = useCallback(
    (pageKey: string, el: HTMLCanvasElement | null) => {
      const previous = canvasesRef.current.get(pageKey);
      if (previous && previous !== el) pageRegistry.delete(previous);
      if (el) {
        canvasesRef.current.set(pageKey, el);
        pageRegistry.add(el);
      } else {
        canvasesRef.current.delete(pageKey);
      }
    },
    [pageRegistry]
  );
  // Runs after this render's page DOM is in place and before any layout effect
  // reads it: memoized pages can move or renumber without their refs rerunning.
  useInsertionEffect(() => pageRegistry.invalidate());
  const transferredCanvasesRef = useRef(new WeakSet<HTMLCanvasElement>());
  const [replayState] = useState(() => new CanvasReplayState());
  const offscreenSignatureRef = useRef('');
  const surfaceRef = useRef('');
  const replayGenerationRef = useRef(0);
  const [offscreenFailed, setOffscreenFailed] = useState(false);
  const offscreenFailedRef = useRef(false);
  const offscreenAttachedRef = useRef(false);
  const pendingAttachRef = useRef<{ generation: number; displayList: DisplayList } | null>(null);
  const workerPresentationRef = useRef(false);
  const publishWorkerPresentation = useCallback(
    (active: boolean) => {
      if (workerPresentationRef.current === active) return;
      workerPresentationRef.current = active;
      onWorkerPresentationChange?.(active);
    },
    [onWorkerPresentationChange]
  );
  const offscreenEligible = useMemo(
    () => Boolean(offscreenReplay && frame && !displayListNeedsHostImages(displayList)),
    [displayList, frame, offscreenReplay]
  );
  useEffect(() => {
    if (!offscreenEligible || offscreenFailed) publishWorkerPresentation(false);
  }, [offscreenEligible, offscreenFailed, publishWorkerPresentation]);
  useEffect(
    () => () => {
      publishWorkerPresentation(false);
    },
    [publishWorkerPresentation]
  );
  // ===========================================================================
  // Page windowing: only pages near the viewport hold rastered bitmaps. Every
  // page keeps its canvas element (stable keys and CSS-sized boxes, so scroll
  // geometry, pointer routing, and canvas-rect overlays are untouched); an
  // off-window page's backing store is released (attributes zeroed on the DOM
  // path, offscreen buffer zeroed by the worker) and repainted on re-entry.
  // The window moves only with scrolling/resize/zoom, never with document
  // invalidation.
  // ===========================================================================
  const innerHostRef = useRef<HTMLDivElement | null>(null);
  const setHostRef = useMemo(
    () =>
      (element: HTMLDivElement | null): void => {
        if (innerHostRef.current && innerHostRef.current !== element) {
          bindDisplayPageRegistry(innerHostRef.current, null);
        }
        innerHostRef.current = element;
        if (element) bindDisplayPageRegistry(element, pageRegistry);
        if (typeof hostRef === 'function') hostRef(element);
        else if (hostRef) (hostRef as { current: HTMLDivElement | null }).current = element;
      },
    [hostRef, pageRegistry]
  );
  const pageWindowAllowed = useMemo(() => {
    if (typeof window === 'undefined') return false;
    // diagnostic escape hatch, mirroring `offscreenReplay=0`
    return new URLSearchParams(window.location.search).get('pageWindow') !== '0';
  }, []);
  const windowingEnabled = pageWindowAllowed && displayList.pages.length > PAGE_WINDOW_MIN_PAGES;
  const [pageWindow, setPageWindow] = useState<PageWindowRange | null>(null);
  const windowMeasuredRef = useRef(false);
  // Column-space page tops/bottoms from display-list geometry alone (no DOM
  // reads): padding, then each page height at the current zoom plus the gap.
  const pageOffsets = useMemo(() => {
    const tops = new Array<number>(displayList.pages.length);
    const bottoms = new Array<number>(displayList.pages.length);
    let y = CANVAS_PAGES_PADDING_PX;
    displayList.pages.forEach((page, index) => {
      tops[index] = y;
      bottoms[index] = y + page.height * zoom;
      y = bottoms[index] + CANVAS_PAGE_GAP_PX;
    });
    return { tops, bottoms };
  }, [displayList, zoom]);
  useLayoutEffect(() => {
    if (!windowingEnabled) {
      windowMeasuredRef.current = false;
      setPageWindow(null);
      return;
    }
    const host = innerHostRef.current;
    if (!host) return;
    const scrollParent = findVerticalScrollParentOrRoot(host);
    const scrollTarget: EventTarget =
      scrollParent === document.scrollingElement || scrollParent === document.documentElement
        ? window
        : scrollParent;
    let rafId: number | null = null;
    const recompute = (): void => {
      rafId = null;
      const column = host.firstElementChild as HTMLElement | null;
      if (!column || !scrollParent.isConnected) {
        // unmeasurable — fail open (all pages live) so replay is never
        // deferred forever
        setPageWindow(
          (previous) => previous ?? { start: 0, end: displayList.pages.length - 1 }
        );
        return;
      }
      const band = viewportColumnBand(scrollTarget === window ? null : scrollParent, column);
      const viewTop = band.top - band.columnTop;
      const viewBottom = viewTop + band.height;
      const { tops, bottoms } = pageOffsets;
      let first = tops.length - 1;
      for (let index = 0; index < tops.length; index += 1) {
        if (bottoms[index] >= viewTop) {
          first = index;
          break;
        }
      }
      let last = first;
      for (let index = tops.length - 1; index >= first; index -= 1) {
        if (tops[index] <= viewBottom) {
          last = index;
          break;
        }
      }
      windowMeasuredRef.current = true;
      setPageWindow((previous) => nextPageWindow(previous, first, last, tops.length));
    };
    const schedule = (): void => {
      if (rafId === null) rafId = requestAnimationFrame(recompute);
    };
    // The first measurement must land before paint; later re-measures coalesce.
    if (windowMeasuredRef.current) schedule();
    else recompute();
    scrollTarget.addEventListener('scroll', schedule, { passive: true });
    window.addEventListener('resize', schedule);
    return () => {
      if (rafId !== null) cancelAnimationFrame(rafId);
      scrollTarget.removeEventListener('scroll', schedule);
      window.removeEventListener('resize', schedule);
    };
  }, [windowingEnabled, pageOffsets]);
  // Until the first measurement lands (set pre-paint by the layout effect
  // above), the replay effect is deferred entirely — never guess a window
  // that could blank a visible page, and never raster every page of a large
  // document just because the window is not measured yet.
  const windowPending = windowingEnabled && pageWindow === null;
  const effectiveWindow: PageWindowRange | null = windowingEnabled ? pageWindow : null;
  const pageInWindow = (index: number): boolean =>
    effectiveWindow === null || (index >= effectiveWindow.start && index <= effectiveWindow.end);
  const chromeInWindow = (index: number): boolean =>
    effectiveWindow === null
      ? !windowingEnabled || index < PAGE_WINDOW_MIN_PAGES
      : pageInWindow(index);
  // The page holding focus (an SDT widget, or assistive-technology focus in
  // its mirror) keeps its chrome when it leaves the window. Pages are pinned
  // by their surface key, which renumbering keeps.
  const [focusedPageKey, setFocusedPageKey] = useState<string | null>(null);
  useEffect(() => {
    const host = innerHostRef.current;
    if (!host) return;
    const pageKeyOf = (target: EventTarget | null): string | null =>
      (target instanceof Element
        ? target.closest<HTMLElement>('.canvas-page')?.dataset.pageKey
        : undefined) ?? null;
    const onFocusIn = (event: FocusEvent) => setFocusedPageKey(pageKeyOf(event.target));
    const onFocusOut = (event: FocusEvent) => {
      if (!(event.relatedTarget instanceof Node) || !host.contains(event.relatedTarget)) {
        setFocusedPageKey(null);
      }
    };
    host.addEventListener('focusin', onFocusIn);
    host.addEventListener('focusout', onFocusOut);
    return () => {
      host.removeEventListener('focusin', onFocusIn);
      host.removeEventListener('focusout', onFocusOut);
    };
  }, []);
  // Focus removed along with its element fires no focusout.
  useEffect(() => {
    const host = innerHostRef.current;
    if (focusedPageKey !== null && !host?.contains(document.activeElement)) {
      setFocusedPageKey(null);
    }
  });

  // Chrome built on demand for pages outside the window: each page's mirror
  // and overlay register a function that builds them at once.
  const pageKeys = useMemo(
    () =>
      displayList.pages.map((page, index) => {
        const retainedPage = frame?.pages[index];
        return retainedPage ? retainedPage.pageId.toString() : `index:${page.pageIndex}`;
      }),
    [displayList, frame]
  );
  const pageKeysRef = useRef(pageKeys);
  pageKeysRef.current = pageKeys;
  const displayListRef = useRef(displayList);
  displayListRef.current = displayList;
  const chromeHandlesRef = useRef(new Map<string, ChromeHandles>());
  const registerChrome = useCallback(
    (pageKey: string, kind: ChromeKind, handle: PageChromeHandle | null) => {
      const registry = chromeHandlesRef.current;
      const entry = registry.get(pageKey) ?? {};
      if (handle) entry[kind] = handle;
      else delete entry[kind];
      if (entry.mirror || entry.overlay) registry.set(pageKey, entry);
      else registry.delete(pageKey);
    },
    []
  );
  // Pages built on demand keep their chrome until the page window moves, so
  // what a plugin query returned stays connected while the view stays put.
  // Past the most recent ON_DEMAND_CHROME_PAGES, they keep it only until the
  // task that built them ends: a scan of the document does not build it all.
  const onDemandKeysRef = useRef(new Set<string>());
  const [onDemandPageKeys, setOnDemandPageKeys] = useState<ReadonlySet<string>>(
    () => new Set()
  );
  const releaseOnDemand = useCallback((keys: Iterable<string>) => {
    const onDemand = onDemandKeysRef.current;
    for (const key of [...keys]) {
      const handles = chromeHandlesRef.current.get(key);
      handles?.mirror?.release();
      handles?.overlay?.release();
      onDemand.delete(key);
    }
    setOnDemandPageKeys(new Set(onDemand));
  }, []);
  const trimTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (trimTimerRef.current !== null) clearTimeout(trimTimerRef.current);
    },
    []
  );
  const materializePages = useCallback(
    (pageIndices: readonly number[]) => {
      const onDemand = onDemandKeysRef.current;
      let added = false;
      for (const index of pageIndices) {
        const key = pageKeysRef.current[index];
        if (key === undefined) continue;
        const handles = chromeHandlesRef.current.get(key);
        handles?.mirror?.build();
        handles?.overlay?.build();
        added ||= !onDemand.has(key);
        // Least recently built first.
        onDemand.delete(key);
        onDemand.add(key);
      }
      if (added) setOnDemandPageKeys(new Set(onDemand));
      if (onDemand.size > ON_DEMAND_CHROME_PAGES && trimTimerRef.current === null) {
        trimTimerRef.current = setTimeout(() => {
          trimTimerRef.current = null;
          const keys = [...onDemandKeysRef.current];
          if (keys.length > ON_DEMAND_CHROME_PAGES) {
            releaseOnDemand(keys.slice(0, keys.length - ON_DEMAND_CHROME_PAGES));
          }
        }, 0);
      }
    },
    [releaseOnDemand]
  );
  useEffect(() => {
    pageRegistry.setMaterializer(materializePages);
    return () => pageRegistry.setMaterializer(null);
  }, [materializePages, pageRegistry]);

  // A fragment link to content on a page whose chrome is not built builds
  // that page first, so the browser finds the target it follows.
  useEffect(() => {
    const host = innerHostRef.current;
    if (!host) return;
    const onClick = (event: MouseEvent): void => {
      const link =
        event.target instanceof Element
          ? event.target.closest<HTMLAnchorElement>('a[href^="#"]')
          : null;
      if (!link || !host.contains(link)) return;
      let id: string;
      try {
        id = decodeURIComponent(link.getAttribute('href')!.slice(1));
      } catch {
        return;
      }
      if (!id || host.ownerDocument.getElementById(id)) return;
      const index = displayListRef.current.pages.findIndex((page) =>
        displayPageHoldsMirrorId(page, id)
      );
      if (index >= 0) materializePages([index]);
    };
    host.addEventListener('click', onClick, true);
    return () => host.removeEventListener('click', onClick, true);
  }, [materializePages]);

  // One glyph-outline cache for the canvas lifetime (task contract: not
  // per-render). The wasm-backed outline provider loads lazily through the
  // SAME module the display-list builder already resolved — no extra fetch.
  // `glyphCacheReady` re-runs the draw effect once the provider lands so the
  // first shaped frame repaints as real glyph outlines (until then a glyphRun
  // falls back to fillText inside the backend, so text is never blank).
  const glyphCacheRef = useRef<GlyphCache | null>(null);
  const [glyphCacheReady, setGlyphCacheReady] = useState(false);
  useEffect(() => {
    setOffscreenFailed(false);
    offscreenFailedRef.current = false;
    offscreenAttachedRef.current = false;
    offscreenSignatureRef.current = '';
  }, [offscreenReplay]);
  useEffect(() => {
    let cancelled = false;
    glyphCacheRef.current = null;
    setGlyphCacheReady(false);
    // A replay still rasterizing once its engine is replaced or unmounted reads no outline from
    // it (the engine may be freed) and falls back to text.
    const outlines = glyphOutlineProvider;
    const provider = outlines
      ? Promise.resolve<GlyphOutlineProvider>((fontId, glyphId) => {
          if (cancelled) throw new Error('The glyph outlines belong to a released engine');
          return outlines(fontId, glyphId);
        })
      : loadGlyphOutlineProvider();
    void provider
      .then((provider) => {
        if (cancelled) return;
        glyphCacheRef.current = new GlyphCache({ provider });
        setGlyphCacheReady(true);
      })
      .catch(() => {
        // outline export absent → the backend keeps painting glyph runs with
        // fillText; nothing to do here.
      });
    return () => {
      cancelled = true;
    };
  }, [glyphOutlineProvider]);

  const windowStart = effectiveWindow?.start ?? -1;
  const windowEnd = effectiveWindow?.end ?? -1;
  const pageCount = displayList.pages.length;
  useLayoutEffect(() => {
    if (windowPending) return;
    if (windowStart < 0) onPageWindowChange?.(0, pageCount);
    else onPageWindowChange?.(windowStart, windowEnd + 1);
  }, [onPageWindowChange, pageCount, windowEnd, windowPending, windowStart]);
  useEffect(() => {
    if (onDemandKeysRef.current.size > 0) releaseOnDemand(onDemandKeysRef.current);
  }, [releaseOnDemand, windowStart, windowEnd]);
  useEffect(() => {
    // The window measurement lands pre-paint (layout effect) and re-runs this
    // effect; rastering before it exists would process every page.
    if (windowPending) return;
    const replayGeneration = ++replayGenerationRef.current;
    const dpr = window.devicePixelRatio || 1;
    if (offscreenEligible && !offscreenFailed && frame && offscreenReplay) {
      const pages: Array<{ pageId: string; canvas: OffscreenCanvas }> = [];
      const activePageIds: string[] = [];
      for (let index = 0; index < frame.pages.length; index += 1) {
        if (!pageInWindow(index)) continue;
        const retainedPage = frame.pages[index];
        const page = displayList.pages[index];
        const pageId = retainedPage.pageId.toString();
        activePageIds.push(pageId);
        const canvas = canvasesRef.current.get(pageId);
        if (!canvas || !page) continue;
        canvas.style.width = `${page.width * zoom}px`;
        canvas.style.height = `${page.height * zoom}px`;
        if (transferredCanvasesRef.current.has(canvas)) continue;
        try {
          pages.push({ pageId, canvas: canvas.transferControlToOffscreen() });
          transferredCanvasesRef.current.add(canvas);
        } catch {
          offscreenFailedRef.current = true;
          publishWorkerPresentation(false);
          setOffscreenFailed(true);
          return;
        }
      }
      const host = innerHostRef.current;
      const caretColor = resolveCaretPaintColor(host);
      const caretStyle = { color: caretColor, width: DEFAULT_CARET_WIDTH };
      const signature = `${activePageIds.join(',')}|${dpr}|${zoom}|${caretColor}`;
      if (pages.length > 0 || signature !== offscreenSignatureRef.current) {
        offscreenSignatureRef.current = signature;
        if (host) clearPresented(host);
        const pendingAttach = { generation: replayGeneration, displayList };
        pendingAttachRef.current = pendingAttach;
        void offscreenReplay.attach(pages, activePageIds, dpr, zoom, caretStyle).then((attached) => {
          // Publish on resolution regardless of replay generation: attach
          // resolutions are FIFO, so the last one reflects the worker's real
          // attachment state. Gating on the generation dropped the publish
          // whenever a frame landed while the first attach was still
          // rastering — i.e. on every document load — leaving presentation
          // permanently unpublished while the worker was in fact presenting.
          offscreenAttachedRef.current = attached;
          if (!attached && pages.length > 0) {
            // No worker took these canvases, and a canvas transfers only
            // once: they can never paint, so the pages remount on the DOM path.
            offscreenFailedRef.current = true;
            setOffscreenFailed(true);
          }
          if (!offscreenFailedRef.current) publishWorkerPresentation(attached);
          if (!attached) {
            // transient (no worker client yet) — clear the signature so the
            // next pass retries instead of permanently flipping surfaces
            offscreenSignatureRef.current = '';
          }
          if (pendingAttachRef.current !== pendingAttach) return;
          pendingAttachRef.current = null;
          const current = pendingAttach.generation === replayGenerationRef.current;
          if (attached && current && innerHostRef.current) {
            markPresented(innerHostRef.current, pendingAttach.displayList, { worker: true });
          }
        }, () => {
          if (pendingAttachRef.current === pendingAttach) pendingAttachRef.current = null;
          offscreenFailedRef.current = true;
          publishWorkerPresentation(false);
          setOffscreenFailed(true);
        });
      } else {
        const pendingAttach = pendingAttachRef.current;
        if (pendingAttach) {
          // The worker replays its latest frame onto attached pages before the attach replies.
          pendingAttach.generation = replayGeneration;
          pendingAttach.displayList = displayList;
        } else if (offscreenAttachedRef.current && host) {
          // The worker presents a frame before it replies with it, so these pages show no other.
          markPresented(host, displayList, { worker: true });
        }
        // Heal any publish lost to ordering (StrictMode remount, late
        // resolution): the worker is attached and this pass kept it active.
        if (offscreenAttachedRef.current) publishWorkerPresentation(true);
      }
      return;
    }
    const surface = `${dpr}|${zoom}`;
    if (surface !== surfaceRef.current) {
      surfaceRef.current = surface;
      if (innerHostRef.current) clearPresented(innerHostRef.current);
    }
    const glyphCache = glyphCacheRef.current ?? undefined;
    replayState.updateFrame(frame);
    const environment = { dpr, zoom, glyphCache, resolveImage, fontFamilies };
    const preparations: CanvasReplayPreparation[] = [];
    for (const [i, page] of displayList.pages.entries()) {
      const retainedPage = frame?.pages[i];
      const pageKey = retainedPage ? retainedPage.pageId.toString() : `index:${page.pageIndex}`;
      const canvas = canvasesRef.current.get(pageKey);
      const ctx = canvas?.getContext('2d');
      if (!canvas || !ctx) continue;
      if (!pageInWindow(i)) {
        // release the off-window bitmap; the CSS-sized element stays for
        // geometry consumers. Dropping the presented mark makes re-entry
        // repaint through the ordinary remount rule below.
        if (canvas.width !== 0 || canvas.height !== 0) {
          canvas.width = 0;
          canvas.height = 0;
          replayState.release(canvas);
        }
        continue;
      }
      // A remounted canvas (surface-mode flip) has no pixels regardless of
      // the retained frame's damage set — always paint it.
      const presentation = replayState.prepare(canvas, retainedPage?.pageId, environment);
      if (!presentation) continue;
      // Raster off-DOM first. The connected canvas keeps its previous pixels
      // until every damaged page has finished all async image/glyph work.
      const buffer = document.createElement('canvas');
      preparations.push({
        buffer,
        ready: rasterizeDisplayPageToBackBuffer(
          buffer,
          page,
          { resolveImage, glyphCache, fontFamilies },
          dpr,
          zoom
        ),
        present() {
          presentDisplayPageBackBuffer(canvas, buffer, page, zoom);
          replayState.didPresent(presentation);
        },
      });
    }
    void presentCanvasReplay(
      preparations,
      () => replayGeneration === replayGenerationRef.current
    ).then(
      (presented) => {
        if (presented && innerHostRef.current) markPresented(innerHostRef.current, displayList);
      },
      (error) => {
        if (replayGeneration === replayGenerationRef.current) {
          console.error('[CanvasRenderer] Canvas replay failed', error);
          markReplayFailed(displayList, error);
        }
      }
    );
    return () => {
      replayGenerationRef.current += 1;
    };
    // glyphCacheReady is a redraw trigger (the cache itself is read via ref);
    // zoom re-runs the raster so the enlarged canvas paints at full resolution;
    // windowStart/windowEnd re-run it so pages entering the window paint and
    // the offscreen active set prunes pages that left it
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    displayList,
    frame,
    resolveImage,
    fontFamilies,
    glyphCacheReady,
    offscreenEligible,
    offscreenFailed,
    offscreenReplay,
    zoom,
    windowPending,
    windowStart,
    windowEnd,
    publishWorkerPresentation,
  ]);

  // The host stays a full-width, un-transformed positioned box so the
  // interactive comment overlays (portalled in by DocxEditorPagedArea) anchor
  // their `50%`-centered X / host-relative Y to the page's un-shifted center.
  // Only the inner page column shifts left when the sidebar opens, mirroring
  // the DOM painter's viewport `translateX(-SIDEBAR_DOCUMENT_SHIFT)`. Pointer
  // routing reads each canvas's live `getBoundingClientRect`, so the transform
  // is factored out for free.
  return (
    <div ref={setHostRef} className="canvas-pages" style={{ position: 'relative' }}>
      <div
        className="canvas-pages__column"
        style={{
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          gap: CANVAS_PAGE_GAP_PX,
          padding: CANVAS_PAGES_PADDING_PX,
          transform: sidebarOpen ? `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px)` : undefined,
          transition: 'transform 0.2s ease',
        }}
      >
        {displayList.pages.map((page, i) => {
          const pageKey = pageKeys[i]!;
          const surfaceKey = `${pageKey}:${offscreenEligible && !offscreenFailed ? 'offscreen' : 'dom'}`;
          // per-page wrapper so the mirror positions 1:1 over its canvas.
          // Every page keeps its sized canvas, so page geometry never
          // changes; the a11y mirror and SDT overlay hold content only for
          // pages in the window, the page holding focus, and pages built on
          // demand. A page in the measured window builds at once, others at
          // idle time; after a content change a page in the window rebuilds
          // at once.
          return (
            <CanvasPageSurface
              key={surfaceKey}
              page={page}
              revision={displayPageRevision(page)}
              noteAnchorRevision={displayPageNoteAnchorRevision(page)}
              pageKey={pageKey}
              zoom={zoom}
              interactive={interactive}
              chrome={
                chromeInWindow(i) ||
                pageKey === focusedPageKey ||
                onDemandPageKeys.has(pageKey)
              }
              inWindow={chromeInWindow(i)}
              deferChrome={windowPending || !chromeInWindow(i)}
              registerCanvas={registerCanvas}
              registerChrome={registerChrome}
            />
          );
        })}
      </div>
    </div>
  );
}
