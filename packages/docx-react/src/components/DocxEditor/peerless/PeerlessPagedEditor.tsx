import { memo, useCallback, useEffect, useInsertionEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type PointerEvent, type ReactElement, type ReactNode, type RefObject } from 'react';
import type { DocxEchoRect } from '@betteroffice/docx/peerless';
import type { DocxCommandPageIdentity } from '@betteroffice/docx/peerless';
import { displayPageNoteAnchorRevision, displayPageRevision } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import { nextPageWindow } from '../CanvasPagesView';
import { viewportColumnBand } from '../internals/viewportBand';
import {
  usePeerlessEditSession,
  type PeerlessEditSessionBinding,
  type UsePeerlessEditSessionOptions,
} from '../hooks/usePeerlessEditSession';
import { PeerlessInput } from './PeerlessInput';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';
import type { PeerlessCompositionState } from './composition';
import { captureSelectionIntent } from './selection';
import type { PeerlessFrameGeometry } from './viewStore';
import { clientPointToPage } from './pagePoint';
import { CanvasPageMirror } from '../CanvasPageMirror';
import { createPageChromeHold, type PageChromeHold } from '../usePageChrome';
import { CanvasA11yLiveRegion } from '../CanvasA11yLiveRegion';
import { DEFAULT_BLINK_INTERVAL } from '../overlays/SelectionOverlay';

type PageGeometry = PeerlessFrameGeometry['pages'][number];
const emptyPages: readonly PageGeometry[] = [];
const SCROLL_SETTLE_MS = 150;
const pageLayouts = new WeakMap<readonly PageGeometry[], { width: number; height: number }>();

function pageLayout(pages: readonly PageGeometry[]) {
  const cached = pageLayouts.get(pages);
  if (cached) return cached;
  const layout = pages.reduce((extent, { bounds }) => ({
    width: Math.max(extent.width, bounds.x + bounds.width), height: Math.max(extent.height, bounds.y + bounds.height),
  }), { width: 0, height: 0 });
  pageLayouts.set(pages, layout);
  return layout;
}

const PeerlessPage = memo(function PeerlessPage({ geometry, canvasForPage, visible, noteAnchorRevision, hold, fullMirror }: {
  geometry: PageGeometry;
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  visible: boolean;
  revision: number;
  noteAnchorRevision: number;
  hold: PageChromeHold;
  fullMirror: boolean;
}): ReactElement {
  const surface = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (!visible) return;
    const canvas = canvasForPage(geometry.page);
    if (!canvas || !surface.current) return;
    canvas.dataset.pageIndex = String(geometry.page.pageIndex);
    canvas.style.width = '100%';
    canvas.style.height = '100%';
    canvas.style.display = 'block';
    surface.current.appendChild(canvas);
    return () => { canvas.remove(); };
  }, [canvasForPage, geometry.page.pageId, geometry.page.pageIndex, visible]);
  return <div
    ref={surface}
    data-page-id={geometry.page.pageId}
    data-page-index={geometry.page.pageIndex}
    className="paged-editor__page canvas-page"
    style={{ position: 'absolute', left: geometry.bounds.x, top: geometry.bounds.y, width: geometry.bounds.width, height: geometry.bounds.height, background: 'white', touchAction: 'none' }}
  >{geometry.displayPage && <CanvasPageMirror page={geometry.displayPage} active={visible && fullMirror} visible={visible}
    noteAnchorRevision={noteAnchorRevision} hold={hold} />}</div>;
}, (previous, next) => previous.canvasForPage === next.canvasForPage &&
  previous.visible === next.visible && previous.hold === next.hold && previous.fullMirror === next.fullMirror &&
  previous.geometry.displayPage === next.geometry.displayPage &&
  previous.revision === next.revision && previous.noteAnchorRevision === next.noteAnchorRevision &&
  previous.geometry.page.pageId === next.geometry.page.pageId &&
  previous.geometry.page.pageIndex === next.geometry.page.pageIndex &&
  previous.geometry.bounds.x === next.geometry.bounds.x && previous.geometry.bounds.y === next.geometry.bounds.y &&
  previous.geometry.bounds.width === next.geometry.bounds.width && previous.geometry.bounds.height === next.geometry.bounds.height);

const PeerlessPages = memo(function PeerlessPages({ pages, visiblePages, canvasForPage, hold, textMirror }: {
  pages: readonly PageGeometry[];
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  visiblePages: readonly PageGeometry[];
  hold: PageChromeHold;
  textMirror: boolean;
}) {
  const visible = new Set(visiblePages.map(({ page }) => page.pageId));
  return <>{pages.filter((page) => visible.has(page.page.pageId) || page.displayPage?.unbuilt === false ||
    (page.displayPage && !page.displayPage.unbuilt)).map((page) => <PeerlessPage key={page.page.pageId}
      geometry={page} visible={visible.has(page.page.pageId)} hold={hold} fullMirror={!textMirror} canvasForPage={canvasForPage}
      revision={page.displayPage ? displayPageRevision(page.displayPage) : 0}
      noteAnchorRevision={page.displayPage ? displayPageNoteAnchorRevision(page.displayPage) : 0} />)}</>;
});

function PeerlessCaret({ caret, input }: { caret: DocxEchoRect; input: RefObject<HTMLTextAreaElement | null> }): ReactElement {
  const [focused, setFocused] = useState(false);
  const [lit, setLit] = useState(true);
  useLayoutEffect(() => {
    const element = input.current;
    if (!element) return;
    const focus = () => setFocused(true);
    const blur = () => setFocused(false);
    setFocused(document.activeElement === element);
    element.addEventListener('focus', focus);
    element.addEventListener('blur', blur);
    return () => {
      element.removeEventListener('focus', focus);
      element.removeEventListener('blur', blur);
    };
  }, [input]);
  useLayoutEffect(() => {
    setLit(true);
    if (!focused) return;
    const blink = setInterval(() => setLit((shown) => !shown), DEFAULT_BLINK_INTERVAL);
    return () => clearInterval(blink);
  }, [focused, caret.x, caret.y]);
  return <div aria-hidden style={{ position: 'absolute', left: caret.x, top: caret.y, width: Math.max(1, caret.width), height: caret.height, background: 'currentColor', opacity: focused && lit ? 1 : 0, pointerEvents: 'none' }} />;
}

function PeerlessReady({ binding, canvasForPage, zoom, autoFocus, onRecoverableText, displayWindow, onDisplayWindowChange, onClipboard, onPasteFragment, highlights, onNavigateCell, onPartPoint, pageOverlays, pageMirror }: {
  binding: PeerlessEditSessionBinding;
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  zoom: number;
  autoFocus?: boolean;
  onRecoverableText(text: string): void;
  displayWindow: readonly [number, number];
  onDisplayWindowChange(start: number, end: number): void;
  onClipboard?(cut: boolean): Promise<void>;
  onPasteFragment?(fragment: DocxClipboardFragment, inputTime: number): boolean;
  highlights?: readonly (DocxEchoRect & { current: boolean })[];
  onNavigateCell?(direction: 'forward' | 'backward'): void;
  onPartPoint?(point: { page: PageGeometry; x: number; y: number; clickCount: number; intent: import('@betteroffice/docx/yrs').DocxSelectionPointIntent }): boolean;
  pageOverlays?: ReactNode;
  pageMirror?: 'full' | 'text';
}): ReactElement {
  const snapshot = useSyncExternalStore(binding.viewStore.subscribe, binding.viewStore.getSnapshot, binding.viewStore.getServerSnapshot);
  const notify = useRef<(() => void) | null>(null);
  const getSelectionContext = useCallback(() => binding.viewStore.getSnapshot().editState?.toolbar, [binding]);
  useLayoutEffect(() => { notify.current?.(); }, [snapshot.editState]);
  const [composition, setComposition] = useState<PeerlessCompositionState | null>(null);
  const onCompositionChange = useCallback((state: PeerlessCompositionState | null) => {
    setComposition((current) => state && current && state.id !== current.id && state.phase !== 'active' ? current : state);
  }, []);
  const input = useRef<HTMLTextAreaElement>(null);
  const click = useRef<{ time: number; x: number; y: number; pageId: string; count: number; press: number } | null>(null);
  const presses = useRef(0);
  const geometry = snapshot.geometry;
  const pages = geometry?.pages ?? emptyPages;
  const layout = useMemo(() => pageLayout(pages), [pages]);
  const width = Math.max(layout.width, geometry?.documentWidth ?? 0);
  const { height } = layout;
  const column = useRef<HTMLDivElement>(null);
  const measuredWindow = useRef<{ start: number; end: number } | null>(null);
  const chromeHold = useMemo(createPageChromeHold, []);
  const scrollSettle = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => () => {
    if (scrollSettle.current !== null) clearTimeout(scrollSettle.current);
    scrollSettle.current = null;
  }, []);
  useLayoutEffect(() => {
    const owner = column.current?.ownerDocument;
    if (!owner) return;
    const press = () => { presses.current += 1; };
    owner.addEventListener('pointerdown', press, true);
    return () => owner.removeEventListener('pointerdown', press, true);
  }, []);
  useLayoutEffect(() => {
    const host = column.current;
    if (!host || !pages.length) return;
    const scrollParent = findVerticalScrollParentOrRoot(host);
    const root = scrollParent === document.scrollingElement || scrollParent === document.documentElement;
    const target = root ? window : scrollParent;
    let raf: number | null = null;
    const measure = () => {
      raf = null;
      let next: { start: number; end: number };
      if (pages.length <= 12) next = { start: 0, end: pages.length - 1 };
      else {
        const band = viewportColumnBand(root ? null : scrollParent, host);
        const top = (band.top - band.columnTop) / zoom;
        const bottom = top + band.height / zoom;
        let first = pages.length - 1;
        for (let index = 0; index < pages.length; index += 1) {
          const bounds = pages[index]!.bounds;
          if (bounds.y + bounds.height >= top) { first = index; break; }
        }
        let last = first;
        for (let index = pages.length - 1; index >= first; index -= 1) {
          if (pages[index]!.bounds.y <= bottom) { last = index; break; }
        }
        next = nextPageWindow(measuredWindow.current, first, last, pages.length);
      }
      measuredWindow.current = next;
      onDisplayWindowChange(pages[next.start]!.page.pageIndex, pages[next.end]!.page.pageIndex + 1);
    };
    const schedule = () => { if (raf === null) raf = requestAnimationFrame(measure); };
    const scroll = () => {
      schedule();
      chromeHold.set(true);
      if (scrollSettle.current !== null) clearTimeout(scrollSettle.current);
      scrollSettle.current = setTimeout(() => {
        scrollSettle.current = null;
        chromeHold.set(false);
      }, SCROLL_SETTLE_MS);
    };
    measure();
    target.addEventListener('scroll', scroll, { passive: true });
    window.addEventListener('resize', schedule);
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(schedule);
    observer?.observe(host);
    if (!root) observer?.observe(scrollParent);
    return () => {
      if (raf !== null) cancelAnimationFrame(raf);
      target.removeEventListener('scroll', scroll);
      window.removeEventListener('resize', schedule);
      observer?.disconnect();
    };
  }, [pages, zoom, onDisplayWindowChange, chromeHold]);
  useLayoutEffect(() => {
    if (geometry && displayWindow[1] > displayWindow[0]) binding.setDisplayWindow?.(displayWindow[0], displayWindow[1]);
  }, [binding, displayWindow[0], displayWindow[1], geometry?.identity.incarnation, geometry?.identity.epoch]);
  const caretPage = geometry?.caret ? clientPointToPage(geometry.caret.x, geometry.caret.y, { left: 0, top: 0 }, 1, pages)?.page.page.pageId : undefined;
  const visiblePages = useMemo(() => pages.filter(({ page }) => page.pageId === caretPage ||
    (page.pageIndex >= displayWindow[0] && page.pageIndex < displayWindow[1])), [pages, caretPage, displayWindow]);
  const dragPointer = useRef<number | null>(null);
  const onPoint = (event: PointerEvent<HTMLDivElement>, drag: boolean) => {
    event.preventDefault();
    const current = binding.viewStore.getSnapshot().geometry;
    if (!current) return;
    const bounds = event.currentTarget.getBoundingClientRect();
    if (bounds.width === 0 || bounds.height === 0) return;
    const point = clientPointToPage(event.clientX, event.clientY, bounds, zoom, current.pages);
    if (!point) return;
    const { page, x, y } = point;
    if (!drag && (x < 0 || x > page.bounds.width || y < 0 || y > page.bounds.height)) return;
    let clickCount = 1;
    if (!drag) {
      const previous = click.current;
      clickCount = previous && previous.press === presses.current - 1 && event.timeStamp - previous.time < 500 && previous.pageId === page.page.pageId &&
        Math.hypot(event.clientX - previous.x, event.clientY - previous.y) < 5 ? previous.count % 3 + 1 : 1;
      click.current = { time: event.timeStamp, x: event.clientX, y: event.clientY, pageId: page.page.pageId, count: clickCount, press: presses.current };
    } else click.current = null;
    const intent = captureSelectionIntent({
      kind: drag ? 'drag' : 'point',
      point: {
        page: page.page, x, y,
        frameEpoch: current.frameEpoch, layoutRevision: current.layoutRevision,
        clickCount,
        modifiers: { shift: event.shiftKey, alt: event.altKey, ctrl: event.ctrlKey, meta: event.metaKey },
      },
    }, current);
    if (!drag && intent.kind === 'selectPoint' && onPartPoint?.({ page, x, y, clickCount, intent: intent.payload })) return true;
    binding.commands.submit(intent, { source: 'selection', fence: true });
    input.current?.focus({ preventScroll: true });
    return true;
  };
  const canDisplaySelection = composition?.phase !== 'active' && geometry?.coversThroughSeq === snapshot.throughSeq;
  const caret = canDisplaySelection ? geometry?.caret : null;
  const releasePointer = (event: PointerEvent<HTMLDivElement>) => {
    if (dragPointer.current !== event.pointerId) return;
    dragPointer.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
  };
  return <>
    <CanvasA11yLiveRegion active getSelectionContext={getSelectionContext} notifyRef={notify} />
    <PeerlessInput admission={binding.admission} viewStore={binding.viewStore} textareaRef={input} autoFocus={autoFocus}
      onCompositionChange={onCompositionChange} onRecoverableText={onRecoverableText} onClipboard={onClipboard} onPasteFragment={onPasteFragment}
      caretHostRef={column} zoom={zoom} onNavigateCell={onNavigateCell} />
    <div ref={column} style={{ position: 'relative', width: width * zoom, height: height * zoom, marginLeft: 'auto', marginRight: 'auto' }}>
      <div data-peerless-pages-host style={{ position: 'relative', width, height, transform: `scale(${zoom})`, transformOrigin: 'top left', touchAction: 'none' }}
        onPointerDown={(event) => {
          if (event.button !== 0 || dragPointer.current !== null || (event.target as HTMLElement).closest('button, input, select, [data-docx-escape-layer]')) return;
          if (!onPoint(event, false)) return;
          event.currentTarget.setPointerCapture(event.pointerId);
          dragPointer.current = event.pointerId;
        }}
        onPointerMove={(event) => {
          if (dragPointer.current === event.pointerId && event.buttons === 1) onPoint(event, true);
        }}
        onPointerUp={releasePointer}
        onPointerCancel={releasePointer}
        onLostPointerCapture={(event) => { if (dragPointer.current === event.pointerId) dragPointer.current = null; }}>
        {pageOverlays}
        <PeerlessPages pages={pages} visiblePages={visiblePages} canvasForPage={canvasForPage} hold={chromeHold}
          textMirror={pageMirror === 'text'} />
        {highlights?.map((rect, index) => <div key={`find-${index}`} aria-hidden style={{ position: 'absolute', left: rect.x, top: rect.y, width: rect.width, height: rect.height, background: rect.current ? 'rgba(255, 150, 0, 0.4)' : 'rgba(255, 220, 0, 0.3)', pointerEvents: 'none' }} />)}
        {canDisplaySelection && geometry?.selectionRects.map((rect, index) => <div key={index} aria-hidden style={{ position: 'absolute', left: rect.x, top: rect.y, width: rect.width, height: rect.height, background: 'rgba(59, 130, 246, 0.25)', pointerEvents: 'none' }} />)}
        {caret && <PeerlessCaret caret={caret} input={input} />}
      </div>
    </div>
  </>;
}

/** @internal */
export interface PeerlessPagedEditorProps {
  sessionOptions: UsePeerlessEditSessionOptions;
  zoom?: number;
  className?: string;
  autoFocus?: boolean;
  onReady?(binding: PeerlessEditSessionBinding): void;
  displayWindow?: readonly [number, number];
  onDisplayWindowChange?(window: readonly [number, number]): void;
  onClipboard?(cut: boolean): Promise<void>;
  onPasteFragment?(fragment: DocxClipboardFragment, inputTime: number): boolean;
  highlights?: readonly (DocxEchoRect & { current: boolean })[];
  onNavigateCell?(direction: 'forward' | 'backward'): void;
  onPartPoint?(point: { page: PageGeometry; x: number; y: number; clickCount: number; intent: import('@betteroffice/docx/yrs').DocxSelectionPointIntent }): boolean;
  pageOverlays?: ReactNode;
  pageMirror?: 'full' | 'text';
}

/** @internal */
export function PeerlessPagedEditor({ sessionOptions, zoom = 1, className, autoFocus, onReady, displayWindow, onDisplayWindowChange, onClipboard, onPasteFragment, highlights, onNavigateCell, onPartPoint, pageOverlays, pageMirror }: PeerlessPagedEditorProps): ReactElement | null {
  const [devicePixelRatio, setDevicePixelRatio] = useState(() => typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1);
  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList;
    const arm = () => {
      media?.removeEventListener('change', change);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', change);
    };
    const change = () => {
      setDevicePixelRatio(window.devicePixelRatio || 1);
      arm();
    };
    change();
    return () => media.removeEventListener('change', change);
  }, []);
  const pageZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const rasterScale = devicePixelRatio * pageZoom;
  const canvases = useRef(new Map<string, HTMLCanvasElement>());
  const [measuredWindow, setMeasuredWindow] = useState<readonly [number, number]>([0, 3]);
  const pageWindow = displayWindow ?? measuredWindow;
  const canvasWindow = useRef(pageWindow);
  useInsertionEffect(() => { canvasWindow.current = pageWindow; }, [pageWindow]);
  const changeWindow = useCallback((start: number, end: number) => {
    setMeasuredWindow((previous) => previous[0] === start && previous[1] === end ? previous : [start, end]);
    if (!displayWindow || displayWindow[0] !== start || displayWindow[1] !== end) onDisplayWindowChange?.([start, end]);
  }, [displayWindow, onDisplayWindowChange]);
  const [recoverableText, setRecoverableText] = useState<readonly string[]>([]);
  const retainText = useCallback((text: string) => setRecoverableText((drafts) => [...drafts, text]), []);
  const { canvasForPage: suppliedCanvas, ...options } = sessionOptions;
  const suppliedCanvasRef = useRef(suppliedCanvas);
  suppliedCanvasRef.current = suppliedCanvas;
  const canvasForPage = useCallback((page: DocxCommandPageIdentity) => {
    let canvas = canvases.current.get(page.pageId);
    if (!canvas) {
      canvas = suppliedCanvasRef.current(page) ?? document.createElement('canvas');
      canvases.current.set(page.pageId, canvas);
    }
    return canvas;
  }, []);
  const evictCanvases = useCallback((geometry: PeerlessFrameGeometry) => {
    const [start, end] = canvasWindow.current;
    const caretPage = geometry.caret ? clientPointToPage(geometry.caret.x, geometry.caret.y, { left: 0, top: 0 }, 1, geometry.pages)?.page.page.pageId : undefined;
    const retained = new Set(geometry.pages.filter(({ page }) => page.pageId === caretPage ||
      (page.pageIndex >= start && page.pageIndex < end)).map(({ page }) => page.pageId));
    for (const [pageId, canvas] of canvases.current) {
      if (retained.has(pageId)) continue;
      canvas.remove();
      canvas.width = 0;
      canvas.height = 0;
      canvases.current.delete(pageId);
    }
  }, []);
  const residentCanvasForPage = useCallback((page: DocxCommandPageIdentity) => canvases.current.get(page.pageId) ?? null, []);
  const prepareGeometry = useCallback((frame: Parameters<UsePeerlessEditSessionOptions['prepareGeometry']>[0]) => {
    const geometry = options.prepareGeometry(frame);
    if (geometry) {
      for (const { page } of frame.payload.stagedPresentation?.pages ?? []) {
        if (geometry.pages.some((candidate) => candidate.page.pageId === page.pageId && candidate.page.pageIndex === page.pageIndex)) {
          canvasForPage(page);
        }
      }
    }
    return geometry;
  }, [options.prepareGeometry, canvasForPage]);
  const session = usePeerlessEditSession(useMemo(() => ({ ...options, rasterScale, canvasForPage: residentCanvasForPage, prepareGeometry }), [
    options.source, options.openConnection, options.recovery, prepareGeometry, options.onError, options.onRecoverableError,
    residentCanvasForPage, rasterScale,
  ]));
  useLayoutEffect(() => {
    if (session.status === 'ready') onReady?.(session.binding);
  }, [session, onReady]);
  useLayoutEffect(() => {
    if (session.status !== 'ready') return;
    let installedGeometry: PeerlessFrameGeometry | null = null;
    const evict = () => {
      const geometry = session.binding.viewStore.getSnapshot().geometry;
      if (!geometry || geometry === installedGeometry) return;
      installedGeometry = geometry;
      evictCanvases(geometry);
    };
    evict();
    return session.binding.viewStore.subscribe(evict);
  }, [session, evictCanvases]);
  useLayoutEffect(() => {
    if (session.status !== 'ready') return;
    const geometry = session.binding.viewStore.getSnapshot().geometry;
    if (geometry) evictCanvases(geometry);
  }, [session, pageWindow, evictCanvases]);
  useLayoutEffect(() => () => {
    for (const canvas of canvases.current.values()) {
      canvas.remove();
      canvas.width = 0;
      canvas.height = 0;
    }
    canvases.current.clear();
  }, [sessionOptions.source, session.status === 'failed']);
  return <div className={className} data-peerless-docx="true" aria-busy={session.status === 'opening'}>
    {session.status === 'ready' && <PeerlessReady binding={session.binding} canvasForPage={canvasForPage} zoom={pageZoom} autoFocus={autoFocus}
      onRecoverableText={retainText} displayWindow={pageWindow} onDisplayWindowChange={changeWindow} onClipboard={onClipboard} onPasteFragment={onPasteFragment} highlights={highlights} onNavigateCell={onNavigateCell} onPartPoint={onPartPoint} pageOverlays={pageOverlays} pageMirror={pageMirror} />}
    {session.status === 'failed' && <div role="alert">Unable to open document.</div>}
    {recoverableText.map((text, index) => <div key={index} role="status">
      This text could not be applied. Copy it to keep your draft, then dismiss it.
      <textarea aria-label="Recoverable document text" value={text} readOnly />
      <button type="button" onClick={() => setRecoverableText((drafts) => drafts.filter((_, draft) => draft !== index))}>Dismiss draft</button>
    </div>)}
  </div>;
}
