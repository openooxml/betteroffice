import { memo, useCallback, useEffect, useInsertionEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type PointerEvent, type ReactElement, type ReactNode, type RefObject } from 'react';
import type { DocxEchoRect } from '@betteroffice/docx/worker';
import type { DocxCommandPageIdentity } from '@betteroffice/docx/worker';
import { displayPageNoteAnchorRevision, displayPageRevision, loadedRustDisplayListQueryEngine, loadRustDisplayListQueryEngine, type DisplayPage } from '@betteroffice/docx/layout/render';
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
import { beginPictureDrag, floatingPictureAt, inlinePictureAt, movePictureDrag, pictureAt, pictureDropIntent, pictureInSelection, type PictureDrag } from './pictureDrag';
import { beginPictureResize, movePictureResize, pictureResizeIntent, type PictureResize, type PictureResizeHandle } from './pictureResize';
import type { PeerlessPropertyIntent } from './inputEvents';
import { CanvasPageMirror } from '../CanvasPageMirror';
import { createPageChromeHold, type PageChromeHold } from '../usePageChrome';
import { CanvasA11yLiveRegion } from '../CanvasA11yLiveRegion';
import { DEFAULT_BLINK_INTERVAL } from '../overlays/SelectionOverlay';
import { createPresentedPageQueries } from './presentedPageQueries';
import { useTranslation } from '../../../i18n';

type PageGeometry = PeerlessFrameGeometry['pages'][number];
const emptyPages: readonly PageGeometry[] = [];
const SCROLL_SETTLE_MS = 150;
const TOUCH_TAP_SLOP_PX = 8;
const TOUCH_LONG_PRESS_MS = 500;
const pictureHandles = [
  { handle: 'nw', x: 0, y: 0, cursor: 'nw-resize' },
  { handle: 'n', x: 0.5, y: 0, cursor: 'ns-resize' },
  { handle: 'ne', x: 1, y: 0, cursor: 'ne-resize' },
  { handle: 'e', x: 1, y: 0.5, cursor: 'ew-resize' },
  { handle: 'se', x: 1, y: 1, cursor: 'se-resize' },
  { handle: 's', x: 0.5, y: 1, cursor: 'ns-resize' },
  { handle: 'sw', x: 0, y: 1, cursor: 'sw-resize' },
  { handle: 'w', x: 0, y: 0.5, cursor: 'ew-resize' },
] as const;
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

function provisionalTextHitAtPoint(page: PageGeometry, x: number, y: number, cache: ReturnType<typeof createPresentedPageQueries>): { caret: DocxEchoRect | null } | null {
  const engine = loadedRustDisplayListQueryEngine();
  if (!engine || !page.displayPage || page.displayPage.unbuilt) return null;
  const queries = cache.get(page.displayPage);
  if (!queries) return null;
  try {
    const hit = queries.hitTestRegions(0, x, y);
    if (hit?.region !== 'body' || hit.pos === null) return null;
    const rect = queries.caretRect(hit.pos);
    return { caret: rect && rect.pageIndex === 0 && rect.height > 0
      ? { x: page.bounds.x + rect.x, y: page.bounds.y + rect.y, width: 1, height: rect.height } : null };
  } finally {}
}

function TableGridlines({ page }: { page: DisplayPage }): ReactElement | null {
  if (!page.tableGridlines?.length) return null;
  return <svg aria-hidden data-table-gridlines="" viewBox={`0 0 ${page.width} ${page.height}`} preserveAspectRatio="none"
    style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', pointerEvents: 'none', overflow: 'visible' }}>
    {page.tableGridlines.map((line, index) => <line key={index} x1={line.x1} y1={line.y1} x2={line.x2} y2={line.y2}
      stroke="#a6a6a6" strokeWidth={1} strokeDasharray="2 2" vectorEffect="non-scaling-stroke" shapeRendering="crispEdges" />)}
  </svg>;
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
    style={{ position: 'absolute', left: geometry.bounds.x, top: geometry.bounds.y, width: geometry.bounds.width, height: geometry.bounds.height, background: 'white', touchAction: 'pan-x pan-y', cursor: 'text' }}
  >{geometry.displayPage && <CanvasPageMirror page={geometry.displayPage} active={visible && fullMirror} visible={visible}
    noteAnchorRevision={noteAnchorRevision} hold={hold} />}
    {geometry.displayPage && <TableGridlines page={geometry.displayPage} />}</div>;
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

function PeerlessReady({ binding, canvasForPage, zoom, autoFocus, onRecoverableText, displayWindow, onDisplayWindowChange, onClipboard, onPasteFragment, highlights, onNavigateCell, onPartPoint, pageOverlays, pageMirror, overlay, pictures, queryEngineReady }: {
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
  onPartPoint?(point: { page: PageGeometry; x: number; y: number; clickCount: number; intent: import('@betteroffice/docx/yrs').DocxSelectionPointIntent }): boolean | { bodyY: number };
  pageOverlays?: ReactNode;
  pageMirror?: 'full' | 'text';
  overlay?: ReactNode;
  pictures?: PeerlessPictureMoves;
  queryEngineReady: boolean;
}): ReactElement {
  const snapshot = useSyncExternalStore(binding.viewStore.subscribe, binding.viewStore.getSnapshot, binding.viewStore.getServerSnapshot);
  const pageQueries = useMemo(createPresentedPageQueries, []);
  useEffect(() => () => pageQueries.dispose(), [pageQueries]);
  const notify = useRef<(() => void) | null>(null);
  const getSelectionContext = useCallback(() => binding.viewStore.getSnapshot().editState?.toolbar, [binding]);
  useLayoutEffect(() => { notify.current?.(); }, [snapshot.editState]);
  const [composition, setComposition] = useState<PeerlessCompositionState | null>(null);
  const onCompositionChange = useCallback((state: PeerlessCompositionState | null) => {
    setComposition((current) => state && current && state.id !== current.id && state.phase !== 'active' ? current : state);
  }, []);
  const input = useRef<HTMLTextAreaElement>(null);
  const click = useRef<{ time: number; x: number; y: number; pageId: string; count: number; press: number } | null>(null);
  const touchPointer = useRef<{ pointerId: number; x: number; y: number } | null>(null);
  const touchLongPress = useRef<ReturnType<typeof setTimeout> | null>(null);
  const touchSelection = useRef<HTMLDivElement | null>(null);
  const cancelTouchPress = useCallback(() => {
    touchPointer.current = null;
    if (touchLongPress.current !== null) clearTimeout(touchLongPress.current);
    touchLongPress.current = null;
  }, []);
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
    const host = column.current?.querySelector<HTMLDivElement>('[data-docx-pages-host]');
    if (!host) return;
    const owner = host.ownerDocument;
    const press = () => { presses.current += 1; cancelTouchPress(); };
    const move = (event: TouchEvent) => { if (touchSelection.current) event.preventDefault(); };
    owner.addEventListener('pointerdown', press, true);
    host.addEventListener('touchmove', move, { passive: false });
    return () => {
      cancelTouchPress();
      owner.removeEventListener('pointerdown', press, true);
      host.removeEventListener('touchmove', move);
    };
  }, [cancelTouchPress]);
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
    // Later changes measure in a frame: a window set while committing chains synchronous renders during a scroll.
    if (measuredWindow.current) schedule(); else measure();
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
  useEffect(() => {
    const presented = visiblePages.flatMap(({ displayPage }) => displayPage && !displayPage.unbuilt ? [displayPage] : []);
    pageQueries.retain(presented);
    let index = 0;
    let cancel = () => {};
    const warm = () => {
      const page = presented[index++];
      if (!page) return;
      pageQueries.get(page)?.hitTestRegions(0, 0, 0);
      if (index < presented.length) schedule();
    };
    const schedule = () => {
      if (typeof requestIdleCallback === 'function') {
        const id = requestIdleCallback(warm);
        cancel = () => cancelIdleCallback(id);
      } else {
        const id = setTimeout(warm, 32);
        cancel = () => clearTimeout(id);
      }
    };
    if (presented.length && queryEngineReady) schedule();
    return () => cancel();
  }, [pageQueries, visiblePages, geometry?.layoutRevision, queryEngineReady]);
  const dragPointer = useRef<number | null>(null);
  const picture = useRef<{ pointerId: number; host: HTMLDivElement; start: { x: number; y: number }; drag: PictureDrag } | null>(null);
  const [pictureDragging, setPictureDragging] = useState(false);
  const [picturePreview, setPicturePreview] = useState<{ x: number; y: number; width: number; height: number } | null>(null);
  const [pictureDropCaret, setPictureDropCaret] = useState<DocxEchoRect | null>(null);
  const canEditPictures = pictures?.canMove() ?? false;
  // Handles follow the worker selection, so a picture selected by click or keyboard shows them and any other selection hides them.
  const selected = geometry && canEditPictures ? pictureInSelection(geometry) : null;
  const canResizePicture = selected != null && (selected.rotation ?? 0) % 360 === 0;
  const selectedPage = selected && pages.find(page => page.page.pageIndex === selected.pageIndex);
  const resize = useRef<{ pointerId: number; host: HTMLDivElement; pageHost: HTMLDivElement; gesture: PictureResize } | null>(null);
  const [resizePreview, setResizePreview] = useState<PictureResize | null>(null);
  const endResize = () => {
    const active = resize.current;
    resize.current = null;
    setResizePreview(null);
    if (active?.host.hasPointerCapture(active.pointerId)) active.host.releasePointerCapture(active.pointerId);
  };
  useLayoutEffect(() => {
    const active = resize.current?.gesture;
    if (active && (!selected || active.target.version !== selected.version || active.target.pageIndex !== selected.pageIndex ||
      active.target.rect.x !== selected.rect.x || active.target.rect.y !== selected.rect.y ||
      active.target.rect.width !== selected.rect.width || active.target.rect.height !== selected.rect.height)) endResize();
  }, [geometry, canEditPictures]);
  const armPicture = (event: PointerEvent<HTMLDivElement>) => {
    picture.current = null;
    const current = binding.viewStore.getSnapshot().geometry;
    // Touch and modified presses (Shift extends a selection) keep the ordinary selection path.
    if (!current || !pictures?.canMove() || event.pointerType === 'touch' || event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) return false;
    const bounds = event.currentTarget.getBoundingClientRect();
    const scale = bounds.width / width;
    const point = clientPointToPage(event.clientX, event.clientY, bounds, scale, current.pages);
    const hit = point && pictureAt(current, point.page.page.pageIndex, point.x, point.y);
    if (!hit || !point) return false;
    // The press first places the caret as any press does, so an earlier selection never survives it, even when the worker
    // refuses the picture selection below as stale. The worker then selects the body picture its handle names.
    if (!onPoint(event, false)) return true;
    if (current.objects?.objects.find(object => object.handle === hit.handle)?.story === 'body') {
      binding.commands.submit({ kind: 'select', payload: { object: { handle: hit.handle, expectVersion: hit.version } } }, { source: 'selection', fence: true });
    }
    const target = floatingPictureAt(current, point.page.page.pageIndex, point.x, point.y) ??
      (pictures.canMoveInline() ? inlinePictureAt(current, point.page.page.pageIndex, point.x, point.y) : null);
    if (target) picture.current = { pointerId: event.pointerId, host: event.currentTarget, start: { x: point.x, y: point.y }, drag: beginPictureDrag(target, event.clientX, event.clientY) };
    return true;
  };
  const picturePoint = (host: HTMLDivElement, pageIndex: number, clientX: number, clientY: number) => {
    const page = binding.viewStore.getSnapshot().geometry?.pages.find(page => page.page.pageIndex === pageIndex);
    const bounds = host.getBoundingClientRect();
    const scale = bounds.width / width;
    if (!page || !Number.isFinite(scale) || scale <= 0) return null;
    return { x: (clientX - bounds.left) / scale - page.bounds.x, y: (clientY - bounds.top) / scale - page.bounds.y, scale };
  };
  const movePicture = (pressed: NonNullable<typeof picture.current>, clientX: number, clientY: number): PictureDrag | null => {
    const point = picturePoint(pressed.host, pressed.drag.target.pageIndex, clientX, clientY);
    return point ? { ...movePictureDrag(pressed.drag, clientX, clientY, point.scale), dx: point.x - pressed.start.x, dy: point.y - pressed.start.y } : null;
  };
  const inlineDropAtPoint = (host: HTMLDivElement, clientX: number, clientY: number) => {
    const current = binding.viewStore.getSnapshot().geometry;
    if (!current) return null;
    const bounds = host.getBoundingClientRect();
    const point = clientPointToPage(clientX, clientY, bounds, bounds.width / width, current.pages);
    if (!point) return null;
    const hit = provisionalTextHitAtPoint(point.page, point.x, point.y, pageQueries);
    if (!hit) return null;
    const intent = captureSelectionIntent({ kind: 'point', point: {
      page: point.page.page, x: point.x, y: point.y,
      frameEpoch: current.frameEpoch, layoutRevision: current.layoutRevision, clickCount: 1,
      modifiers: { shift: false, alt: false, ctrl: false, meta: false },
    } }, current);
    return intent.kind === 'selectPoint' ? { caret: hit.caret, drop: intent.payload } : null;
  };
  const resizePicture = (active: NonNullable<typeof resize.current>, clientX: number, clientY: number, shift: boolean): PictureResize | null => {
    const point = picturePoint(active.pageHost, active.gesture.target.pageIndex, clientX, clientY);
    return point ? movePictureResize(active.gesture, point.x, point.y, 1, shift) : null;
  };
  const endPicture = () => {
    picture.current = null;
    setPictureDragging(false);
    setPicturePreview(null);
    setPictureDropCaret(null);
  };
  const previewingPicture = pictureDragging || picturePreview !== null || pictureDropCaret !== null;
  const selectedHandle = selected?.handle ?? null;
  useEffect(() => {
    if (!selectedHandle && !previewingPicture) return;
    const cancel = (event: KeyboardEvent) => {
      const active = picture.current;
      // With no gesture running, only Escape typed in the editor leaves the picture; menus and dialogs keep theirs.
      if (event.key !== 'Escape' || (!active && !resize.current && event.target !== input.current)) return;
      event.preventDefault();
      event.stopPropagation();
      if (resize.current) { endResize(); return; }
      if (!active) {
        // Escape leaves the picture: the selection collapses to its end, after the picture.
        binding.commands.submit({ kind: 'navigate', payload: { unit: 'character', direction: 'forward', extend: false } }, { source: 'selection', fence: true });
        return;
      }
      endPicture();
      dragPointer.current = null;
      if (active.host.hasPointerCapture(active.pointerId)) active.host.releasePointerCapture(active.pointerId);
    };
    document.addEventListener('keydown', cancel, true);
    return () => document.removeEventListener('keydown', cancel, true);
  }, [selectedHandle, previewingPicture]);
  const startResize = (event: PointerEvent<HTMLDivElement>, handle: PictureResizeHandle) => {
    event.stopPropagation();
    event.preventDefault();
    if (event.button !== 0 || event.isPrimary === false || resize.current || picture.current || !selected || !canResizePicture || !pictures?.canMove()) return;
    const pageHost = event.currentTarget.closest<HTMLDivElement>('[data-docx-pages-host]');
    const point = pageHost && picturePoint(pageHost, selected.pageIndex, event.clientX, event.clientY);
    if (!pageHost || !point) return;
    const gesture = beginPictureResize(selected, handle, point.x, point.y);
    resize.current = { pointerId: event.pointerId, host: event.currentTarget, pageHost, gesture };
    setResizePreview(gesture);
    event.currentTarget.setPointerCapture(event.pointerId);
    input.current?.focus({ preventScroll: true });
  };
  const onPoint = (event: PointerEvent<HTMLDivElement>, drag: boolean, host = event.currentTarget) => {
    event.preventDefault();
    const currentSnapshot = binding.viewStore.getSnapshot();
    const current = currentSnapshot.geometry;
    if (!current) return;
    const bounds = host.getBoundingClientRect();
    if (bounds.width === 0 || bounds.height === 0) return;
    const point = clientPointToPage(event.clientX, event.clientY, bounds, zoom, current.pages);
    if (!point) return;
    const { page, x } = point;
    let { y } = point;
    if (!drag && (x < 0 || x > page.bounds.width || y < 0 || y > page.bounds.height)) return;
    let clickCount = 1;
    if (!drag) {
      const previous = click.current;
      clickCount = previous && previous.press === presses.current - 1 && event.timeStamp - previous.time < 500 && previous.pageId === page.page.pageId &&
        Math.hypot(event.clientX - previous.x, event.clientY - previous.y) < (event.pointerType === 'touch' ? TOUCH_TAP_SLOP_PX : 5) ? previous.count % 3 + 1 : 1;
      click.current = { time: event.timeStamp, x: event.clientX, y: event.clientY, pageId: page.page.pageId, count: clickCount, press: presses.current };
    } else click.current = null;
    const capture = () => captureSelectionIntent({
      kind: drag ? 'drag' : 'point',
      point: {
        page: page.page, x, y,
        frameEpoch: current.frameEpoch, layoutRevision: current.layoutRevision,
        clickCount,
        modifiers: { shift: event.shiftKey, alt: event.altKey, ctrl: event.ctrlKey, meta: event.metaKey },
      },
    }, current);
    let intent = capture();
    if (!drag && intent.kind === 'selectPoint') {
      const part = onPartPoint?.({ page, x, y, clickCount, intent: intent.payload });
      if (part === true) return true;
      if (part) {
        y = part.bodyY;
        intent = capture();
      }
    }
    const admission = binding.admission.getSnapshot();
    const tailSeq = admission.tail?.seq;
    const provisional = !drag && event.type === 'pointerdown' && event.pointerType !== 'touch' && event.button === 0 &&
      event.isPrimary !== false && (event.detail === 0 || event.detail === 1) && clickCount === 1 && intent.kind === 'selectPoint' &&
      !event.shiftKey && !event.altKey && !event.ctrlKey && !event.metaKey && composition?.phase !== 'active' &&
      admission.compositionSlot === null && !admission.readOnly &&
      (!admission.tail || tailSeq !== undefined && tailSeq <= current.coversThroughSeq) &&
      current.coversThroughSeq === currentSnapshot.throughSeq ? provisionalTextHitAtPoint(page, x, y, pageQueries)?.caret ?? null : null;
    const result = binding.commands.submit(intent, { source: 'selection', fence: true });
    if (provisional && result.status === 'submitted') {
      const caret = { get throughSeq() { return result.ticket.seq; }, pageIndex: page.page.pageIndex, rect: provisional };
      binding.viewStore.setProvisionalCaret(caret);
      void result.ticket.presented.then(() => {
        if (binding.viewStore.getSnapshot().provisionalCaret === caret) binding.viewStore.setProvisionalCaret(null);
      });
    }
    input.current?.focus({ preventScroll: true });
    return true;
  };
  const canDisplaySelection = composition?.phase !== 'active' && !snapshot.provisionalCaret && geometry?.coversThroughSeq === snapshot.throughSeq;
  const caret = composition?.phase === 'active' ? null : snapshot.provisionalCaret?.rect ?? (canDisplaySelection ? geometry?.caret : null);
  const hideCaret = !snapshot.provisionalCaret && snapshot.editState?.tableContext?.hasMultiCellSelection === true;
  const releasePointer = (event: PointerEvent<HTMLDivElement>) => {
    if (resize.current?.pointerId === event.pointerId) endResize();
    if (touchPointer.current?.pointerId === event.pointerId) cancelTouchPress();
    if (picture.current?.pointerId === event.pointerId) endPicture();
    if (dragPointer.current !== event.pointerId) return;
    dragPointer.current = null;
    if (touchSelection.current) {
      touchSelection.current.style.touchAction = 'pan-x pan-y';
      touchSelection.current = null;
    }
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
  };
  return <>
    <CanvasA11yLiveRegion active getSelectionContext={getSelectionContext} notifyRef={notify} />
    <PeerlessInput admission={binding.admission} viewStore={binding.viewStore} textareaRef={input} autoFocus={autoFocus}
      onCompositionChange={onCompositionChange} onRecoverableText={onRecoverableText} onClipboard={onClipboard} onPasteFragment={onPasteFragment}
      caretHostRef={column} zoom={zoom} onNavigateCell={onNavigateCell} />
    <div ref={column} style={{ position: 'relative', width: width * zoom, height: height * zoom, marginLeft: 'auto', marginRight: 'auto' }}>
      <div data-docx-pages-host style={{ position: 'relative', width, height, transform: `scale(${zoom})`, transformOrigin: 'top left', touchAction: 'pan-x pan-y' }}
        onMouseDown={(event) => {
          if (event.button === 0 && !(event.target as HTMLElement).closest('button, input, select, [data-docx-escape-layer]')) event.preventDefault();
        }}
        onPointerDown={(event) => {
          if (event.button !== 0 || (event.pointerType === 'touch' && event.isPrimary === false) || dragPointer.current !== null || (event.target as HTMLElement).closest('button, input, select, [data-docx-escape-layer]')) return;
          if (armPicture(event)) {
            input.current?.focus({ preventScroll: true });
            if (picture.current) {
              event.currentTarget.setPointerCapture(event.pointerId);
              dragPointer.current = event.pointerId;
            }
            return;
          }
          if (event.pointerType === 'touch') {
            touchPointer.current = { pointerId: event.pointerId, x: event.clientX, y: event.clientY };
            const host = event.currentTarget;
            touchLongPress.current = setTimeout(() => {
              cancelTouchPress();
              if (!onPoint(event, false, host)) return;
              host.setPointerCapture(event.pointerId);
              dragPointer.current = event.pointerId;
              touchSelection.current = host;
              host.style.touchAction = 'none';
            }, TOUCH_LONG_PRESS_MS);
            return;
          }
          if (!onPoint(event, false)) { picture.current = null; return; }
          event.currentTarget.setPointerCapture(event.pointerId);
          dragPointer.current = event.pointerId;
        }}
        onPointerMove={(event) => {
          const resizing = resize.current;
          if (resizing?.pointerId === event.pointerId) {
            if (event.buttons !== 1) return;
            const gesture = resizePicture(resizing, event.clientX, event.clientY, event.shiftKey);
            if (!gesture) { endResize(); return; }
            resizing.gesture = gesture;
            setResizePreview(resizing.gesture);
            return;
          }
          const touch = touchPointer.current;
          if (touch?.pointerId === event.pointerId && Math.hypot(event.clientX - touch.x, event.clientY - touch.y) > TOUCH_TAP_SLOP_PX) cancelTouchPress();
          const pressed = picture.current;
          if (pressed?.pointerId === event.pointerId && dragPointer.current === event.pointerId) {
            if (event.buttons !== 1) return;
            const drag = movePicture(pressed, event.clientX, event.clientY);
            if (!drag || drag.target.floating === false && !pictures?.canMoveInline()) { endPicture(); return; }
            pressed.drag = drag;
            setPictureDragging(drag.dragging);
            if (drag.target.floating === false) {
              setPictureDropCaret(drag.dragging ? inlineDropAtPoint(pressed.host, event.clientX, event.clientY)?.caret ?? null : null);
              return;
            }
            const page = drag.dragging ? binding.viewStore.getSnapshot().geometry?.pages.find(page => page.page.pageIndex === drag.target.pageIndex) : null;
            if (page) setPicturePreview({ x: page.bounds.x + drag.target.rect.x + drag.dx, y: page.bounds.y + drag.target.rect.y + drag.dy,
              width: drag.target.rect.width, height: drag.target.rect.height });
            return;
          }
          if (dragPointer.current === event.pointerId && (event.buttons === 1 || touchSelection.current)) onPoint(event, true);
        }}
        onPointerUp={(event) => {
          const resizing = resize.current;
          if (resizing?.pointerId === event.pointerId) {
            const gesture = resizePicture(resizing, event.clientX, event.clientY, event.shiftKey);
            const intent = gesture && pictureResizeIntent(gesture);
            endResize();
            if (intent && pictures?.canMove()) pictures.resize(intent);
            return;
          }
          const touch = touchPointer.current;
          if (touch?.pointerId === event.pointerId) {
            cancelTouchPress();
            if (Math.hypot(event.clientX - touch.x, event.clientY - touch.y) <= TOUCH_TAP_SLOP_PX) onPoint(event, false);
          }
          const pressed = picture.current;
          if (pressed?.pointerId === event.pointerId) {
            const drag = movePicture(pressed, event.clientX, event.clientY);
            const drop = drag?.dragging && drag.target.floating === false ? inlineDropAtPoint(pressed.host, event.clientX, event.clientY) : null;
            const intent = drag && pictureDropIntent(drag, drop?.drop);
            endPicture();
            if (intent && pictures?.canMove() && (intent.kind !== 'moveEmbed' || pictures.canMoveInline())) pictures.move(intent);
          }
          releasePointer(event);
        }}
        onPointerCancel={releasePointer}
        onLostPointerCapture={(event) => {
          if (event.target === event.currentTarget) releasePointer(event);
        }}>
        {pageOverlays}
        <PeerlessPages pages={pages} visiblePages={visiblePages} canvasForPage={canvasForPage} hold={chromeHold}
          textMirror={pageMirror === 'text'} />
        {overlay}
        {highlights?.map((rect, index) => <div key={`find-${index}`} aria-hidden style={{ position: 'absolute', left: rect.x, top: rect.y, width: rect.width, height: rect.height, background: rect.current ? 'rgba(255, 150, 0, 0.4)' : 'rgba(255, 220, 0, 0.3)', pointerEvents: 'none' }} />)}
        {canDisplaySelection && geometry?.selectionRects.map((rect, index) => <div key={index} aria-hidden style={{ position: 'absolute', left: rect.x, top: rect.y, width: rect.width, height: rect.height, background: 'rgba(59, 130, 246, 0.25)', pointerEvents: 'none' }} />)}
        {caret && !hideCaret && <PeerlessCaret caret={caret} input={input} />}
        {picturePreview && <div className="paged-editor__picture-drag-preview" aria-hidden style={{ position: 'absolute', left: picturePreview.x, top: picturePreview.y,
          width: picturePreview.width, height: picturePreview.height, outline: '1px dashed rgb(59, 130, 246)', background: 'rgba(59, 130, 246, 0.12)', pointerEvents: 'none', cursor: 'move' }} />}
        {pictureDropCaret && <div className="paged-editor__picture-drop-caret" aria-hidden style={{ position: 'absolute', left: pictureDropCaret.x, top: pictureDropCaret.y,
          width: 1, height: pictureDropCaret.height, background: '#2563eb', pointerEvents: 'none', zIndex: 16 }} />}
        {selected && selectedPage && <div className="canvas-image-selection-overlay" aria-hidden style={{ position: 'absolute',
          left: selectedPage.bounds.x + selected.rect.x, top: selectedPage.bounds.y + selected.rect.y,
          width: selected.rect.width, height: selected.rect.height, outline: '2px solid #2563eb', pointerEvents: 'none', zIndex: 15 }}>
          {canResizePicture && pictureHandles.map(({ handle, x, y, cursor }) => <div key={handle} data-handle={handle}
            onPointerDown={event => startResize(event, handle)} onLostPointerCapture={event => { if (resize.current?.host === event.currentTarget) endResize(); }}
            style={{ position: 'absolute', left: `${x * 100}%`, top: `${y * 100}%`, width: 10, height: 10,
              transform: 'translate(-50%, -50%)', background: 'white', border: '1.5px solid #2563eb', borderRadius: '50%',
              boxSizing: 'border-box', boxShadow: '0 1px 2.5px #0006', pointerEvents: 'auto', touchAction: 'none', cursor }} />)}
        </div>}
        {resizePreview && selectedPage && <div className="paged-editor__picture-resize-preview" aria-hidden style={{ position: 'absolute',
          left: selectedPage.bounds.x + resizePreview.rect.x, top: selectedPage.bounds.y + resizePreview.rect.y,
          width: resizePreview.rect.width, height: resizePreview.rect.height, outline: '1px dashed rgb(59, 130, 246)',
          background: 'rgba(59, 130, 246, 0.12)', pointerEvents: 'none', zIndex: 16 }} />}
      </div>
    </div>
  </>;
}

/** @internal */
export interface PeerlessPagedEditorProps {
  overlay?: ReactNode;
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
  onPartPoint?(point: { page: PageGeometry; x: number; y: number; clickCount: number; intent: import('@betteroffice/docx/yrs').DocxSelectionPointIntent }): boolean | { bodyY: number };
  pageOverlays?: ReactNode;
  pageMirror?: 'full' | 'text';
  /** Picture selection, floating drags and resizing. */
  pictures?: PeerlessPictureMoves;
}

/** @internal */
export interface PeerlessPictureMoves {
  canMove(): boolean;
  canMoveInline(): boolean;
  move(intent: PeerlessPropertyIntent): void;
  resize(intent: PeerlessPropertyIntent): void;
}

/** @internal */
export function PeerlessPagedEditor({ sessionOptions, zoom = 1, className, autoFocus, onReady, displayWindow, onDisplayWindowChange, onClipboard, onPasteFragment, highlights, onNavigateCell, onPartPoint, pageOverlays, pageMirror, overlay, pictures }: PeerlessPagedEditorProps): ReactElement | null {
  const { t } = useTranslation();
  const [queryEngineReady, setQueryEngineReady] = useState(() => !!loadedRustDisplayListQueryEngine());
  useEffect(() => {
    let active = true;
    if (!loadedRustDisplayListQueryEngine()) void loadRustDisplayListQueryEngine().then(() => {
      if (active) setQueryEngineReady(true);
    }).catch(() => {});
    else setQueryEngineReady(true);
    return () => { active = false; };
  }, []);
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
  const windowChanged = useRef(onDisplayWindowChange);
  useInsertionEffect(() => { windowChanged.current = onDisplayWindowChange; }, [onDisplayWindowChange]);
  // Stable, so a new window does not re-run the window measurement.
  const changeWindow = useCallback((start: number, end: number) => {
    setMeasuredWindow((previous) => previous[0] === start && previous[1] === end ? previous : [start, end]);
    const [shownStart, shownEnd] = canvasWindow.current;
    if (shownStart !== start || shownEnd !== end) windowChanged.current?.([start, end]);
  }, []);
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
  return <div className={className} data-docx-worker-editor="true" aria-busy={session.status === 'opening'}>
    {session.status === 'ready' && <PeerlessReady binding={session.binding} canvasForPage={canvasForPage} zoom={pageZoom} autoFocus={autoFocus}
      onRecoverableText={retainText} displayWindow={pageWindow} onDisplayWindowChange={changeWindow} onClipboard={onClipboard} onPasteFragment={onPasteFragment} highlights={highlights} onNavigateCell={onNavigateCell} onPartPoint={onPartPoint} pageOverlays={pageOverlays} pageMirror={pageMirror} overlay={overlay} pictures={pictures} queryEngineReady={queryEngineReady} />}
    {session.status === 'failed' && <div role="alert">{t('errors.failedToLoad')}</div>}
    {recoverableText.map((text, index) => <div key={index} role="status">
      {t('editor.recoverableTextNotice')}
      <textarea aria-label={t('editor.recoverableText')} value={text} readOnly />
      <button type="button" onClick={() => setRecoverableText((drafts) => drafts.filter((_, draft) => draft !== index))}>{t('editor.dismissDraft')}</button>
    </div>)}
  </div>;
}
