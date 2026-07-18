import { useEffect, useMemo, useRef, useState, type ReactNode, type Ref } from 'react';
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
import { getPerfTrace } from './internals/perfTrace';

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

// experimental canvas replay of a display list: one <canvas> per page, sized
// for devicePixelRatio and painted by the core backend, with the a11y mirror
// mounted 1:1 under each page canvas (the page's accessible content and the
// e2e assertion surface). dumb glue — every layout and style decision already
// happened upstream in the display list. the white page background is
// document content (word-faithful), not themed ui chrome, hence the inline
// color.
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
   * `sizeCanvasForPage`. The a11y/e2e mirror is CSS-scaled to keep its nodes 1:1
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

  useEffect(() => {
    const replayGeneration = ++replayGenerationRef.current;
    const trace = getPerfTrace();
    const started = trace ? performance.now() : 0;
    const dpr = window.devicePixelRatio || 1;
    if (offscreenEligible && !offscreenFailed && frame && offscreenReplay) {
      const pages: Array<{ pageId: string; canvas: OffscreenCanvas }> = [];
      const activePageIds: string[] = [];
      for (let index = 0; index < frame.pages.length; index += 1) {
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
          if (!attached) setOffscreenFailed(true);
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
      const damaged =
        !frame ||
        !retainedPage ||
        rasterEnvironmentChanged ||
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
      }
      if (trace) {
        trace.record('canvasReplay', performance.now() - started, {
          calls: prepared.length,
          detail: `${prepared.length}/${displayList.pages.length} pages; atomic`,
        });
      }
    };
    void Promise.all(preparations).then(present, (error) => {
      if (replayGeneration === replayGenerationRef.current) {
        console.error('[CanvasRenderer] Atomic canvas replay failed', error);
      }
    });
    // glyphCacheReady is a redraw trigger (the cache itself is read via ref);
    // zoom re-runs the raster so the enlarged canvas paints at full resolution
  }, [
    displayList,
    frame,
    resolveImage,
    glyphCacheReady,
    offscreenEligible,
    offscreenFailed,
    offscreenReplay,
    zoom,
  ]);

  // The host stays a full-width, un-transformed positioned box so the
  // interactive comment overlays (portalled in by DocxEditorPagedArea) anchor
  // their `50%`-centered X / host-relative Y to the page's un-shifted center.
  // Only the inner page column shifts left when the sidebar opens, mirroring
  // the DOM painter's viewport `translateX(-SIDEBAR_DOCUMENT_SHIFT)`. Pointer
  // routing reads each canvas's live `getBoundingClientRect`, so the transform
  // is factored out for free.
  return (
    <div ref={hostRef} className="canvas-pages" style={{ position: 'relative' }}>
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
            // per-page wrapper so the mirror positions 1:1 over its canvas
            <div key={surfaceKey} className="canvas-page" style={{ position: 'relative' }}>
              <canvas
                ref={(el) => {
                  if (el) canvasesRef.current.set(pageKey, el);
                  else canvasesRef.current.delete(pageKey);
                }}
                data-page-index={page.pageIndex}
                // header/footer band geometry (page-local px) — test-only hooks so
                // the e2e suite can assert ink inside the bands without reaching
                // into the display list
                data-header-y={page.header?.y}
                data-header-h={page.header?.height}
                data-footer-y={page.footer?.y}
                data-footer-h={page.footer?.height}
                style={{
                  display: 'block',
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
