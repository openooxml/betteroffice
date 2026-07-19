import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
  type Ref,
} from 'react';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import {
  presentDisplayPageBackBuffer,
  rasterizeDisplayPageToBackBuffer,
  GlyphCache,
  loadGlyphOutlineProvider,
  type DisplayList,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import type { UseCanvasRendererResult } from './hooks/useDisplayList';
import { CanvasPageMirror } from './CanvasPageMirror';
import { CanvasInteractiveOverlay } from './CanvasInteractiveOverlay';
import { CanvasA11yLiveRegion, type CanvasA11yLiveRegionProps } from './CanvasA11yLiveRegion';
import { CANVAS_PAGE_GAP_PX, CANVAS_PAGES_PADDING_PX } from '@betteroffice/docx/layout/render';
import { SIDEBAR_DOCUMENT_SHIFT } from '../sidebar/constants';
import { DefaultLoadingIndicator, ParseError } from '../DocxEditorHelpers';

// Canvas is the sole visible renderer. The editing/input subtree stays mounted
// independently so hidden ProseMirror focus and IME state survive initial
// loading and renderer errors.
export function CanvasPagedArea({
  renderer,
  a11y,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  children,
}: {
  renderer: UseCanvasRendererResult;
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
          offscreenReplay={renderer.offscreenReplay}
        />
      ) : renderer.status === 'error' ? (
        <div data-testid="canvas-renderer-error" role="alert" style={{ minHeight: 240 }}>
          <ParseError message={renderer.error?.message ?? 'Canvas renderer failed.'} />
        </div>
      ) : (
        <div data-testid="canvas-renderer-loading" role="status" style={{ minHeight: 240 }}>
          <DefaultLoadingIndicator />
        </div>
      )}
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
// A page already mounted stays mounted until it drifts one page beyond the
// mount band, so slow scrolling at a boundary cannot thrash mount/unmount.
const PAGE_WINDOW_HYSTERESIS = 1;

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

// experimental canvas replay of a display list: one <canvas> per page, sized
// for devicePixelRatio and painted by the core backend, with the a11y mirror
// mounted 1:1 under each page canvas. dumb glue — every layout and style
// decision already happened upstream in the display list. the white page
// background is document content (word-faithful), not themed ui chrome, hence
// the inline color.
export function CanvasPagesView({
  displayList,
  frame,
  resolveImage,
  hostRef,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  glyphOutlineProvider,
  offscreenReplay,
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
  /** Dedicated worker replay surface; unsupported/media-heavy pages use DOM canvas. */
  offscreenReplay?: UseCanvasRendererResult['offscreenReplay'];
}) {
  const canvasesRef = useRef(new Map<string, HTMLCanvasElement>());
  const transferredCanvasesRef = useRef(new WeakSet<HTMLCanvasElement>());
  const presentedCanvasesRef = useRef(new WeakSet<HTMLCanvasElement>());
  const offscreenSignatureRef = useRef('');
  const replayGenerationRef = useRef(0);
  const [offscreenFailed, setOffscreenFailed] = useState(false);
  const offscreenEligible = useMemo(
    () => Boolean(offscreenReplay && frame && !displayListNeedsHostImages(displayList)),
    [displayList, frame, offscreenReplay]
  );
  const rasterEnvironmentRef = useRef<{
    dpr: number;
    zoom: number;
    glyphCacheReady: boolean;
    resolveImage?: ImageResolver;
  } | null>(null);

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
        innerHostRef.current = element;
        if (typeof hostRef === 'function') hostRef(element);
        else if (hostRef) (hostRef as { current: HTMLDivElement | null }).current = element;
      },
    [hostRef]
  );
  const pageWindowAllowed = useMemo(() => {
    if (typeof window === 'undefined') return false;
    // diagnostic escape hatch, mirroring `offscreenReplay=0`
    return new URLSearchParams(window.location.search).get('pageWindow') !== '0';
  }, []);
  const windowingEnabled = pageWindowAllowed && displayList.pages.length > PAGE_WINDOW_MIN_PAGES;
  const [pageWindow, setPageWindow] = useState<PageWindowRange | null>(null);
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
      // client rects are viewport-relative: the visible band starts at the
      // scroller's client top for an element scroller, at 0 for the root
      const viewportTop = scrollTarget === window ? 0 : scrollParent.getBoundingClientRect().top;
      const viewportHeight =
        scrollTarget === window ? window.innerHeight : scrollParent.clientHeight;
      const columnRect = column.getBoundingClientRect();
      const viewTop = viewportTop - columnRect.top;
      const viewBottom = viewTop + viewportHeight;
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
      setPageWindow((previous) => nextPageWindow(previous, first, last, tops.length));
    };
    const schedule = (): void => {
      if (rafId === null) rafId = requestAnimationFrame(recompute);
    };
    recompute();
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
    offscreenSignatureRef.current = '';
  }, [offscreenReplay]);
  useEffect(() => {
    let cancelled = false;
    glyphCacheRef.current = null;
    setGlyphCacheReady(false);
    const provider = glyphOutlineProvider
      ? Promise.resolve(glyphOutlineProvider)
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
          setOffscreenFailed(true);
          return;
        }
      }
      const signature = `${activePageIds.join(',')}|${dpr}|${zoom}`;
      if (pages.length > 0 || signature !== offscreenSignatureRef.current) {
        offscreenSignatureRef.current = signature;
        void offscreenReplay.attach(pages, activePageIds, dpr, zoom).then((attached) => {
          if (!attached) {
            // transient (no worker client yet) — clear the signature so the
            // next pass retries instead of permanently flipping surfaces
            offscreenSignatureRef.current = '';
          }
        }, () => setOffscreenFailed(true));
      }
      return;
    }
    const glyphCache = glyphCacheRef.current ?? undefined;
    const previousEnvironment = rasterEnvironmentRef.current;
    const rasterEnvironmentChanged =
      !previousEnvironment ||
      previousEnvironment.dpr !== dpr ||
      previousEnvironment.zoom !== zoom ||
      previousEnvironment.glyphCacheReady !== glyphCacheReady ||
      previousEnvironment.resolveImage !== resolveImage;
    rasterEnvironmentRef.current = { dpr, zoom, glyphCacheReady, resolveImage };
    const preparations: Array<
      Promise<{
        canvas: HTMLCanvasElement;
        buffer: HTMLCanvasElement;
        page: DisplayList['pages'][number];
      }>
    > = [];
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
          presentedCanvasesRef.current.delete(canvas);
        }
        continue;
      }
      // A remounted canvas (surface-mode flip) has no pixels regardless of
      // the retained frame's damage set — always paint it.
      const damaged =
        !frame ||
        !retainedPage ||
        rasterEnvironmentChanged ||
        !presentedCanvasesRef.current.has(canvas) ||
        frame.damagedPageIds.has(retainedPage.pageId);
      if (!damaged) continue;
      // Raster off-DOM first. The connected canvas keeps its previous pixels
      // until every damaged page has finished all async image/glyph work.
      const buffer = document.createElement('canvas');
      preparations.push(
        rasterizeDisplayPageToBackBuffer(
          buffer,
          page,
          { resolveImage, glyphCache },
          dpr,
          zoom
        ).then(() => ({ canvas, buffer, page }))
      );
    }
    const present = (prepared: Awaited<(typeof preparations)[number]>[]) => {
      if (replayGeneration !== replayGenerationRef.current) return;
      for (const { canvas, buffer, page } of prepared) {
        presentDisplayPageBackBuffer(canvas, buffer, page, zoom);
        presentedCanvasesRef.current.add(canvas);
      }
    };
    void Promise.all(preparations).then(present, (error) => {
      if (replayGeneration === replayGenerationRef.current) {
        console.error('[CanvasRenderer] Atomic canvas replay failed', error);
      }
    });
    // glyphCacheReady is a redraw trigger (the cache itself is read via ref);
    // zoom re-runs the raster so the enlarged canvas paints at full resolution;
    // windowStart/windowEnd re-run it so pages entering the window paint and
    // the offscreen active set prunes pages that left it
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    displayList,
    frame,
    resolveImage,
    glyphCacheReady,
    offscreenEligible,
    offscreenFailed,
    offscreenReplay,
    zoom,
    windowPending,
    windowStart,
    windowEnd,
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
          const retainedPage = frame?.pages[i];
          const pageKey = retainedPage ? retainedPage.pageId.toString() : `index:${page.pageIndex}`;
          const surfaceKey = `${pageKey}:${offscreenEligible && !offscreenFailed ? 'offscreen' : 'dom'}`;
          return (
            // per-page wrapper so the mirror positions 1:1 over its canvas.
            // Every page keeps its full DOM (canvas element, a11y mirror, SDT
            // overlay) — the page window releases only bitmap backing stores,
            // so the accessible document and page geometry never shrink.
            <div key={surfaceKey} className="canvas-page" style={{ position: 'relative' }}>
              <canvas
                ref={(el) => {
                  if (el) canvasesRef.current.set(pageKey, el);
                  else canvasesRef.current.delete(pageKey);
                }}
                data-page-index={page.pageIndex}
                style={{
                  display: 'block',
                  width: page.width * zoom,
                  height: page.height * zoom,
                  background: '#ffffff',
                  boxShadow: '0 1px 3px var(--doc-shadow)',
                }}
              />
              <CanvasPageMirror page={page} zoom={zoom} />
              {interactive ? <CanvasInteractiveOverlay page={page} zoom={zoom} /> : null}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function displayListNeedsHostImages(displayList: DisplayList): boolean {
  const visit = (value: unknown): boolean => {
    if (!value || typeof value !== 'object') return false;
    if (Array.isArray(value)) return value.some(visit);
    const record = value as Record<string, unknown>;
    if (record.kind === 'image' || record.kind === 'picture') return true;
    return Object.values(record).some(visit);
  };
  return visit(displayList.pages);
}
