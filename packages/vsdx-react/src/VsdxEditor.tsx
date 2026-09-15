import { createT, deepMerge, diagnosticMessage, en } from '@betteroffice/vsdx-i18n';
import type { Translations } from '@betteroffice/vsdx-i18n';
import { canvasPointToModel, initWasm, openDiagram, paintPage, sizeCanvasForPage } from '@betteroffice/vsdx';
import type { Affine, PageLayer, PagePrimitive, CollaborationReplica, DiagramHandle, DiagramSnapshot, HitTestResult, ModelPoint, PageDisplayList, PageSnapshot, ShapeSnapshot, TextDiagnostic, VsdxFontFace, VsdxPresence } from '@betteroffice/vsdx';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, DragEvent, FocusEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode } from 'react';
import { Ribbon } from './components/ribbon/Ribbon';
import { ShapeContextMenu } from './components/ribbon/ShapeContextMenu';
import { RibbonCommandsProvider, findShapePlacement, isHandleResizeBlocked, numericCellValue, useRibbonCommands } from './components/ribbon/commands';
import type { RibbonCommands } from './components/ribbon/commands';
import { STENCIL_DRAG_MIME, ShapesPanel } from './components/shapes/ShapesPanel';
import { LayersPanel } from './components/layers/LayersPanel';
import { standardShapeById, standardShapes } from './components/shapes/shapeLibrary';
import type { StandardShape } from './components/shapes/shapeLibrary';
import { StatusBar, clampZoom } from './components/statusbar';
import { paintDragPreview, paintSelectionFrame, passedDragThreshold, previewOutline, hitTestSelection, isPrintableEntryKey, resolveDragGeometry, resolveNudgeGeometry, resolveRotationAngle, resizeCursor, canvasKeyboardIntent, textEditOverlay, withoutTextBox } from './interactions';
import type { DragStart, ResizeHandle } from './interactions';
export { resolveDragGeometry };
export type { DragStart };

export interface VsdxShapeSelection { pageId: string; shapeId: string; hit: HitTestResult; }
export interface VsdxEditorApi { handle: DiagramHandle; refresh: () => void; }
/** Save edits before changing a session identity or seed, or remount for a new session. */
export interface VsdxEditorCollaborationOptions {
  clientId: number;
  initialUpdate?: Uint8Array;
  onReplica?: (replica: CollaborationReplica | null) => void;
  presence?: VsdxPresence;
}
export interface VsdxEditorProps {
  file?: Uint8Array;
  fonts: ReadonlyArray<VsdxFontFace>;
  clientId?: number;
  collaboration?: VsdxEditorCollaborationOptions;
  i18n?: Translations;
  className?: string;
  onReady?: (api: VsdxEditorApi) => void;
  onChange?: () => void;
  onError?: (error: Error) => void;
  leftPanel?: ReactNode;
  statusBar?: ReactNode;
}

interface EditorModel { snapshot: DiagramSnapshot | null; pageIndex: number; frame: PageDisplayList | null; layers: PageLayer[]; }

export function VsdxEditor({ file, fonts, clientId, collaboration, i18n, className, onReady, onChange, onError, leftPanel, statusBar }: VsdxEditorProps) {
  const strings = useMemo(() => deepMerge(en, i18n) as typeof en, [i18n]);
  const t = useMemo(() => createT(strings), [strings]);
  const handleRef = useRef<DiagramHandle | null>(null);
  const onReadyRef = useRef(onReady);
  const onChangeRef = useRef(onChange);
  const onErrorRef = useRef(onError);
  const collaborationRef = useRef(collaboration);
  const attachedCollaborationRef = useRef<VsdxEditorCollaborationOptions | undefined>(undefined);
  const mainCanvasRef = useRef<HTMLCanvasElement>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement>(null);
  const workspaceRef = useRef<HTMLElement>(null);
  const imageCache = useRef(new Map<string, Promise<CanvasImageSource | null>>());
  const stableFonts = useStableFontFaces(fonts);
  const fontsRef = useRef(stableFonts);
  const registeredFontsRef = useRef<ReadonlyArray<VsdxFontFace>>([]);
  const browserFontsRef = useRef(new Map<string, FontFace>());
  fontsRef.current = stableFonts;
  const requestedClientId = collaboration?.clientId ?? clientId;
  const requestedInitialUpdate = useStableInitialUpdate(collaboration?.initialUpdate);
  const sessionRef = useRef({ file, clientId: requestedClientId, initialUpdate: requestedInitialUpdate });
  const [model, setModel] = useState<EditorModel>({ snapshot: null, pageIndex: 0, frame: null, layers: [] });
  const modelRef = useRef(model);
  const [selection, setSelection] = useState<VsdxShapeSelection | null>(null);
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  const [editing, setEditing] = useState<{ pageId: string; shapeId: string; selectAll: boolean } | null>(null);
  const [draft, setDraft] = useState('');
  const editingRef = useRef(editing);
  const draftRef = useRef(draft);
  const committedTextRef = useRef('');
  const editWrapRef = useRef<HTMLDivElement>(null);
  const editBoxRef = useRef<HTMLTextAreaElement>(null);
  editingRef.current = editing;
  draftRef.current = draft;
  const [dirty, setDirty] = useState(false);
  const sessionSwitchBlocked = dirty && sessionRef.current.file === file &&
    (sessionRef.current.clientId !== requestedClientId || sessionRef.current.initialUpdate !== requestedInitialUpdate);
  const sessionClientId = sessionSwitchBlocked ? sessionRef.current.clientId : requestedClientId;
  const initialUpdate = sessionSwitchBlocked ? sessionRef.current.initialUpdate : requestedInitialUpdate;
  const [zoom, setZoom] = useState(1);
  const [shapesCollapsed, setShapesCollapsed] = useState(false);
  const [layersCollapsed, setLayersCollapsed] = useState(false);
  const [diagnostics, setDiagnostics] = useState<TextDiagnostic[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [contextMenu, setContextMenu] = useState<{ top: number; left: number } | null>(null);
  const pointerRef = useRef<DragStart | null>(null);
  const dragPreviewRef = useRef<ModelPoint | null>(null);
  const dragSnapRef = useRef(false);
  const previewFrameRef = useRef<number | null>(null);
  const insertCascadeRef = useRef<Map<string, number>>(new Map());
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const [loading, setLoading] = useState(Boolean(file));
  onReadyRef.current = onReady;
  onChangeRef.current = onChange;
  onErrorRef.current = onError;
  collaborationRef.current = collaboration;

  const reportError = useCallback((value: unknown) => { const next = value instanceof Error ? value : new Error(String(value)); setError(next.message); onErrorRef.current?.(next); }, []);
  const enterTextEdit = useCallback((pageId: string, shapeId: string, override?: string) => {
    const handle = handleRef.current;
    if (!handle) return;
    let committed = '';
    try { committed = handle.shapeText(pageId, shapeId); }
    catch { committed = ''; }
    const value = override ?? committed;
    committedTextRef.current = committed;
    setDraft(value);
    draftRef.current = value;
    const next = { pageId, shapeId, selectAll: override === undefined };
    setEditing(next);
    editingRef.current = next;
  }, []);
  const refresh = useCallback((requestedPage?: number, notify = false) => {
    const handle = handleRef.current;
    if (!handle) return;
    if (notify) setDirty(true);
    try {
      if (notify) onChangeRef.current?.();
      const current = handle.snapshot();
      const previous = modelRef.current;
      const activeId = previous.snapshot?.pages[previous.pageIndex]?.id;
      const retainedIndex = current.pages.findIndex((page) => page.id === activeId);
      const pageIndex = Math.max(0, Math.min(requestedPage ?? (retainedIndex >= 0 ? retainedIndex : previous.pageIndex), Math.max(0, current.pages.length - 1)));
      const frame = current.pages.length ? handle.layoutPage(pageIndex) : null;
      const layers = current.pages.length ? readPageLayers(handle, pageIndex) : [];
      modelRef.current = { snapshot: current, pageIndex, frame, layers };
      setModel(modelRef.current);
      setDiagnostics(frame ? collectDiagnostics(frame) : []);
      setSelection((existing) => existing && stillSelectable(current, pageIndex, existing, layers) ? existing : null);
      const open = editingRef.current;
      if (open) {
        const committed = handle.shapeText(open.pageId, open.shapeId);
        if (committed !== committedTextRef.current) {
          if (draftRef.current === committedTextRef.current) { setDraft(committed); draftRef.current = committed; }
          committedTextRef.current = committed;
        }
      }
    } catch (value) { reportError(value); }
  }, [reportError]);
  const commitTextEdit = useCallback(() => {
    const current = editingRef.current;
    const handle = handleRef.current;
    if (!current || !handle) return;
    const text = draftRef.current;
    setEditing(null);
    editingRef.current = null;
    if (text === committedTextRef.current) return;
    try { handle.setShapeText(current.pageId, current.shapeId, text); }
    catch (value) { reportError(value); return; }
    refresh(undefined, true);
  }, [refresh, reportError]);

  useEffect(() => {
    sessionRef.current = { file, clientId: sessionClientId, initialUpdate };
    let disposed = false;
    let handle: DiagramHandle | null = null;
    let stopUpdates = () => {};
    let stopResync = () => {};
    handleRef.current?.dispose(); handleRef.current = null; imageCache.current.clear(); setSelection(null); setEditing(null); setDraft(''); modelRef.current = { snapshot: null, pageIndex: 0, frame: null, layers: [] }; setModel(modelRef.current); setError(null); setDirty(false);
    if (!file) { setLoading(false); return; }
    setLoading(true);
    const openingFonts = fontsRef.current;
    void Promise.all([initWasm(), loadFonts(openingFonts)]).then(([, loadedFonts]) => {
      if (disposed) return;
      try {
        const activeCollaboration = collaborationRef.current;
        handle = openDiagram(file, { clientId: sessionClientId, fonts: openingFonts, initialUpdate }); registeredFontsRef.current = openingFonts; installFonts(loadedFonts, browserFontsRef.current); handleRef.current = handle; activeCollaboration?.onReplica?.(handle); attachedCollaborationRef.current = activeCollaboration;
        stopUpdates = handle.onUpdate(() => refresh(undefined, true));
        stopResync = handle.onResync(() => refresh(undefined, true));
        refresh(0); setLoading(false); onReadyRef.current?.({ handle, refresh: () => refresh(undefined, false) });
      } catch (value) { setLoading(false); reportError(value); }
    }, (value: unknown) => { if (!disposed) { setLoading(false); reportError(value); } });
    return () => {
      disposed = true;
      try { attachedCollaborationRef.current?.onReplica?.(null); }
      finally {
        attachedCollaborationRef.current = undefined;
        stopUpdates(); stopResync(); handle?.dispose();
        if (handleRef.current === handle) handleRef.current = null;
        for (const face of browserFontsRef.current.values()) document.fonts.delete?.(face);
        browserFontsRef.current.clear();
      }
    };
  }, [sessionClientId, initialUpdate, file, refresh, reportError]);

  useEffect(() => {
    const handle = handleRef.current;
    const attached = attachedCollaborationRef.current;
    if (!handle || sessionSwitchBlocked || attached?.onReplica === collaboration?.onReplica) return;
    attached?.onReplica?.(null);
    collaboration?.onReplica?.(handle);
    attachedCollaborationRef.current = collaboration;
  }, [collaboration, sessionSwitchBlocked]);

  useEffect(() => {
    if (sessionSwitchBlocked) reportError(new Error('Save your changes before switching collaboration sessions.'));
  }, [sessionSwitchBlocked, reportError]);

  const hasDocument = model.snapshot !== null;
  useEffect(() => {
    const handle = handleRef.current;
    if (!handle) return;
    const additions = stableFonts.filter((face) => !registeredFontsRef.current.some((registered) => fontFaceEqual(face, registered)));
    if (!additions.length) return;
    let disposed = false;
    void loadFonts(additions).then((loadedFonts) => {
      if (disposed || handleRef.current !== handle) return;
      try {
        for (const [index, face] of additions.entries()) {
          handle.registerFont(face);
          if (loadedFonts[index]) installFonts([loadedFonts[index]], browserFontsRef.current);
          registeredFontsRef.current = [...registeredFontsRef.current.filter((registered) => !fontFaceKeyEqual(face, registered)), face];
        }
        refresh();
      } catch (value) { reportError(value); }
    }, (value: unknown) => { if (!disposed) reportError(value); });
    return () => { disposed = true; };
  }, [stableFonts, hasDocument, refresh, reportError]);

  const editedTextId = editing && model.snapshot ? textPrimitiveId(model.snapshot, model.pageIndex, editing.pageId, editing.shapeId) : null;
  const paintFrame = useMemo(
    () => (model.frame && editedTextId ? { ...model.frame, primitives: withoutTextBox(model.frame.primitives, editedTextId) } : model.frame),
    [model.frame, editedTextId],
  );

  useEffect(() => {
    const canvas = mainCanvasRef.current; const frame = paintFrame;
    if (!canvas || !frame) return;
    const context = canvas.getContext('2d'); if (!context) return;
    const controller = new AbortController();
    const originHandle = handleRef.current;
    const dpr = window.devicePixelRatio || 1; sizeCanvasForPage(canvas, frame, dpr, zoom);
    void paintPage(context, frame, dpr, zoom, {
      signal: controller.signal,
      resolveImage: async (assetId) => {
        if (controller.signal.aborted) return null;
        try { return await resolveImage(assetId, originHandle, imageCache, t('errors.decodePageImage')); }
        catch (value) { if (!controller.signal.aborted) reportError(value); return null; }
      },
    }).catch((value) => { if (!controller.signal.aborted) reportError(value); });
    return () => controller.abort();
  }, [paintFrame, reportError, t, zoom]);

  useEffect(() => {
    const canvas = overlayCanvasRef.current; const frame = model.frame;
    if (!canvas || !frame) return;
    const context = canvas.getContext('2d'); if (!context) return;
    const dpr = window.devicePixelRatio || 1; sizeCanvasForPage(canvas, frame, dpr, zoom); context.clearRect(0, 0, canvas.width, canvas.height);
    const snapshot = model.snapshot;
    const page = snapshot?.pages[model.pageIndex];
    if (selection && page && !selectionHiddenByLayers(page, model.layers, selection)) {
      try {
        const corners = selectionCorners(page, frame, selection);
        const placement = findShapePlacement(page.shapes, selection.shapeId);
        const blocked = placement ? isHandleResizeBlocked(placement.shape) : false;
        if (corners) paintSelectionFrame(context, corners, dpr, zoom, blocked ? [] : undefined);
      } catch { void 0; }
    }
    const start = pointerRef.current; const release = dragPreviewRef.current;
    if (start && release) {
      try { paintDragPreview(context, previewOutline(start, release, frame.paintTransform), dpr, zoom); } catch { void 0; }
    }
  }, [model.frame, model.snapshot, model.pageIndex, selection, zoom]);

  useEffect(() => {
    if (!editing) return;
    if (!selection || selection.pageId !== editing.pageId || selection.shapeId !== editing.shapeId) setEditing(null);
  }, [editing, selection]);

  useEffect(() => {
    const box = editBoxRef.current;
    if (!editing || !box) return;
    box.focus();
    if (editing.selectAll) box.select();
    else box.setSelectionRange(box.value.length, box.value.length);
  }, [editing]);

  useEffect(() => {
    const box = editBoxRef.current;
    if (!editing || !box) return;
    box.style.height = 'auto';
    box.style.height = `${box.scrollHeight}px`;
  }, [editing, draft, zoom, model.frame]);

  useEffect(() => {
    if (!editing) return;
    const onDown = (event: globalThis.PointerEvent) => {
      if (editWrapRef.current?.contains(event.target as Node)) return;
      commitTextEdit();
    };
    document.addEventListener('pointerdown', onDown, true);
    return () => document.removeEventListener('pointerdown', onDown, true);
  }, [editing, commitTextEdit]);

  useEffect(() => () => { if (previewFrameRef.current !== null) cancelAnimationFrame(previewFrameRef.current); }, []);

  const clearDragPreview = () => {
    if (previewFrameRef.current !== null) { cancelAnimationFrame(previewFrameRef.current); previewFrameRef.current = null; }
    dragPreviewRef.current = null;
    repaintOverlaySelection();
  };

  const repaintOverlaySelection = () => {
    const overlay = overlayCanvasRef.current; const current = modelRef.current; const frame = current.frame;
    if (!overlay || !frame) return;
    const context = overlay.getContext('2d'); if (!context) return;
    const currentSelection = selectionRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    context.clearRect(0, 0, overlay.width, overlay.height);
    if (currentSelection && page && !selectionHiddenByLayers(page, current.layers, currentSelection)) {
      try {
        const corners = selectionCorners(page, frame, currentSelection);
        const placement = findShapePlacement(page.shapes, currentSelection.shapeId);
        const blocked = placement ? isHandleResizeBlocked(placement.shape) : false;
        if (corners) paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current, blocked ? [] : undefined);
      } catch { void 0; }
    }
  };

  const onPointerDown = (event: PointerEvent<HTMLCanvasElement>) => {
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    if (event.button === 2) return;
    if (pointerRef.current) return;
    pointerRef.current = null; dragPreviewRef.current = null;
    try {
      const point = canvasPointerPosition(event, frame);
      const active = selectionRef.current;
      if (active && active.pageId === page.id && !selectionHiddenByLayers(page, modelRef.current.layers, active)) {
        try {
          const corners = selectionCorners(page, frame, active);
          if (corners) {
            const target = hitTestSelection(point.canvas, corners, zoomRef.current);
            if (target) {
              const placement = findShapePlacement(page.shapes, active.shapeId);
              if (placement) {
                if (target !== 'rotate' && isHandleResizeBlocked(placement.shape)) {
                  reportError(new Error(t('errors.resizeLocked')));
                  return;
                }
                const base = shapeDragStart(page, placement.shape, frame, handle);
                pointerRef.current = {
                  ...point,
                  ...base,
                  pointerId: event.pointerId,
                  startX: event.clientX,
                  startY: event.clientY,
                  resize: false,
                  handle: target === 'rotate' ? undefined : (target as ResizeHandle),
                  rotate: target === 'rotate',
                };
                event.currentTarget.setPointerCapture(event.pointerId);
                return;
              }
            }
          }
        } catch { void 0; }
      }
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const next = hit ? selectionForHit(page, hit) : null;
      setSelection(next);
      const placement = next ? findShapePlacement(page.shapes, next.shapeId) : null;
      pointerRef.current = next && placement ? {
        ...point,
        ...shapeDragStart(page, placement.shape, frame, handle),
        pointerId: event.pointerId,
        startX: event.clientX,
        startY: event.clientY,
        resize: event.shiftKey,
      } : null;
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch (value) { reportError(value); }
  };
  const onPointerMove = (event: PointerEvent<HTMLCanvasElement>) => {
    const start = pointerRef.current;
    if (!start) {
      try {
        const current = modelRef.current; const frame = current.frame;
        const page = current.snapshot?.pages[current.pageIndex];
        const active = selectionRef.current;
        if (!frame || !page || !active || active.pageId !== page.id || selectionHiddenByLayers(page, current.layers, active)) { event.currentTarget.style.cursor = ''; return; }
        const corners = selectionCorners(page, frame, active);
        if (!corners) { event.currentTarget.style.cursor = ''; return; }
        const point = canvasPointerPosition(event, frame);
        const target = hitTestSelection(point.canvas, corners, zoomRef.current);
        if (target !== 'rotate' && target) {
          const placement = findShapePlacement(page.shapes, active.shapeId);
          if (placement && isHandleResizeBlocked(placement.shape)) { event.currentTarget.style.cursor = ''; return; }
        }
        event.currentTarget.style.cursor = target === 'rotate' ? 'grab' : target ? resizeCursor(target) : '';
      } catch { void 0; }
      return;
    }
    if (start.pointerId !== undefined && start.pointerId !== event.pointerId) return;
    if (start.startX !== undefined && start.startY !== undefined && !start.thresholdPassed) {
      if (!passedDragThreshold(start.startX, start.startY, event.clientX, event.clientY)) return;
      start.thresholdPassed = true;
    }
    const frame = modelRef.current.frame;
    if (!frame) return;
    try {
      const point = canvasPointerPosition(event, frame);
      dragPreviewRef.current = point.model;
      dragSnapRef.current = event.shiftKey;
      if (previewFrameRef.current !== null) return;
      previewFrameRef.current = requestAnimationFrame(() => {
        previewFrameRef.current = null;
        const liveFrame = modelRef.current.frame; const liveStart = pointerRef.current; const release = dragPreviewRef.current; const overlay = overlayCanvasRef.current;
        if (!liveFrame || !liveStart || !release || !overlay) return;
        const context = overlay.getContext('2d'); if (!context) return;
        try {
          const corners = previewOutline(liveStart, release, liveFrame.paintTransform, dragSnapRef.current);
          context.clearRect(0, 0, overlay.width, overlay.height);
          paintDragPreview(context, corners, window.devicePixelRatio || 1, zoomRef.current);
          paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current);
        } catch (value) { reportError(value); }
      });
    } catch (value) { reportError(value); }
  };
  const onPointerUp = (event: PointerEvent<HTMLCanvasElement>) => {
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null;
    const hadPreview = dragPreviewRef.current !== null;
    clearDragPreview();
    const handle = handleRef.current; const selected = selection; const frame = model.frame;
    if (!handle || !selected || !frame) return;
    try {
      const point = canvasPointerPosition(event, frame);
      if (!pointer.thresholdPassed && !hadPreview && pointer.startX !== undefined && pointer.startY !== undefined && !passedDragThreshold(pointer.startX, pointer.startY, event.clientX, event.clientY)) return;
      if (!pointer.thresholdPassed && !hadPreview && Math.abs(point.canvas.x - pointer.canvas.x) < 0.01 && Math.abs(point.canvas.y - pointer.canvas.y) < 0.01) return;
      if (pointer.rotate) {
        handle.setCellFormula(selected.pageId, selected.shapeId, { cellName: 'Angle' }, String(resolveRotationAngle(pointer, point.model, event.shiftKey)));
        refresh(undefined, true);
        return;
      }
      const geometry = resolveDragGeometry(pointer, point.model);
      if (pointer.handle) {
        const livePage = handle.snapshot().pages.find((page) => page.id === selected.pageId);
        const livePlacement = livePage ? findShapePlacement(livePage.shapes, selected.shapeId) : null;
        if (livePlacement && isHandleResizeBlocked(livePlacement.shape)) throw new Error(t('errors.resizeLocked'));
        handle.setShapeBounds(selected.pageId, selected.shapeId, inchFormula(geometry.x), inchFormula(geometry.y), inchFormula(geometry.width), inchFormula(geometry.height));
      }
      else if (pointer.resize) handle.resizeShape(selected.pageId, selected.shapeId, inchFormula(geometry.width), inchFormula(geometry.height));
      else handle.moveShape(selected.pageId, selected.shapeId, inchFormula(geometry.x), inchFormula(geometry.y));
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  };
  const onPointerCancel = (event: PointerEvent<HTMLCanvasElement>) => {
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null; clearDragPreview();
  };
  const onLostPointerCapture = (event: PointerEvent<HTMLCanvasElement>) => {
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null; clearDragPreview();
  };
  const closeContextMenu = () => setContextMenu(null);
  const closeContextMenuAndFocus = () => { setContextMenu(null); mainCanvasRef.current?.focus(); };
  const onCanvasDoubleClick = (event: MouseEvent<HTMLCanvasElement>) => {
    if (editingRef.current) return;
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    try {
      const point = canvasPointerPosition(event, frame);
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const next = hit ? selectionForHit(page, hit) : null;
      if (!next) return;
      setSelection(next);
      enterTextEdit(next.pageId, next.shapeId);
    } catch (value) { reportError(value); }
  };
  const onCanvasContextMenu = (event: MouseEvent<HTMLCanvasElement>) => {
    event.preventDefault();
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    if (pointerRef.current) return;
    try {
      const point = canvasPointerPosition(event, frame);
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const next = hit ? selectionForHit(page, hit) : null;
      if (!next) {
        closeContextMenu();
        setSelection(null);
        return;
      }
      const active = selectionRef.current;
      if (!active || active.pageId !== next.pageId || active.shapeId !== next.shapeId) setSelection(next);
      setContextMenu({ top: event.clientY, left: event.clientX });
    } catch (value) { reportError(value); }
  };
  const commandsRef = useRef<RibbonCommands | null>(null);
  const cancelActiveDrag = () => {
    const pointer = pointerRef.current;
    if (!pointer) return false;
    pointerRef.current = null;
    clearDragPreview();
    const canvas = mainCanvasRef.current;
    if (canvas && pointer.pointerId !== undefined) {
      try {
        if (typeof canvas.hasPointerCapture !== 'function' || canvas.hasPointerCapture(pointer.pointerId)) canvas.releasePointerCapture(pointer.pointerId);
      } catch { void 0; }
    }
    return true;
  };
  const nudgeSelection = (dx: number, dy: number) => {
    const handle = handleRef.current; const selected = selectionRef.current;
    if (!handle || !selected) return;
    try {
      const snapshot = handle.snapshot();
      const page = snapshot.pages.find((entry) => entry.id === selected.pageId);
      if (!page) return;
      const placement = findShapePlacement(page.shapes, selected.shapeId);
      if (!placement) return;
      const frame = modelRef.current.frame;
      if (!frame) return;
      const base = shapeDragStart(page, placement.shape, frame);
      const geometry = resolveNudgeGeometry({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, ...base }, dx, dy);
      handle.moveShape(selected.pageId, selected.shapeId, inchFormula(geometry.x), inchFormula(geometry.y));
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  };
  const onCanvasKeyDown = (event: KeyboardEvent<HTMLCanvasElement>) => {
    const selected = selectionRef.current;
    if (selected && !editingRef.current) {
      if (event.key === 'Enter') { event.preventDefault(); enterTextEdit(selected.pageId, selected.shapeId); return; }
      if (isPrintableEntryKey(event)) { event.preventDefault(); enterTextEdit(selected.pageId, selected.shapeId, event.key); return; }
    }
    const intent = canvasKeyboardIntent(event, zoomRef.current);
    if (!intent) return;
    event.preventDefault();
    const commands = commandsRef.current;
    if (intent.kind === 'undo') { if (commands?.undo.enabled) commands.undo.run(); return; }
    if (intent.kind === 'redo') { if (commands?.redo.enabled) commands.redo.run(); return; }
    if (intent.kind === 'delete') { if (commands?.delete.enabled) commands.delete.run(); return; }
    if (intent.kind === 'escape') { cancelActiveDrag(); closeContextMenu(); setSelection(null); return; }
    nudgeSelection(intent.dx, intent.dy);
  };
  const onCanvasFocus = (event: FocusEvent<HTMLCanvasElement>) => { event.currentTarget.style.outline = '2px solid #0f6cbd'; event.currentTarget.style.outlineOffset = '2px'; };
  const onCanvasBlur = (event: FocusEvent<HTMLCanvasElement>) => { event.currentTarget.style.outline = ''; event.currentTarget.style.outlineOffset = ''; };
  const insertShapeAt = useCallback((shape: StandardShape, point: ModelPoint) => {
    const handle = handleRef.current; const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !page) return;
    try {
      const receipt = handle.addShape(page.id, shape.draft(point.x, point.y, shape.defaultSize.width, shape.defaultSize.height));
      refresh(undefined, true);
      setSelection({ pageId: page.id, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } });
    } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const insertShape = useCallback((shape: StandardShape) => {
    const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!frame || !page) return;
    const cascade = insertCascadeRef.current.get(page.id) ?? 0;
    insertCascadeRef.current.set(page.id, cascade + 1);
    insertShapeAt(shape, centreInsertPoint(canvasPointToModel(frame.paintTransform, frame.width / 2, frame.height / 2), cascade));
  }, [insertShapeAt]);
  const onCanvasDragOver = (event: DragEvent<HTMLDivElement>) => {
    if (!handleRef.current || !modelRef.current.frame) return;
    if (!event.dataTransfer.types.includes(STENCIL_DRAG_MIME)) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  };
  const onCanvasDrop = (event: DragEvent<HTMLDivElement>) => {
    const frame = modelRef.current.frame; const canvas = mainCanvasRef.current;
    if (!frame || !canvas) return;
    const shape = standardShapeById(event.dataTransfer.getData(STENCIL_DRAG_MIME).trim());
    if (!shape) return;
    event.preventDefault();
    insertShapeAt(shape, clientPointToModel(frame, canvas.getBoundingClientRect(), event.clientX, event.clientY).model);
    canvas.focus();
  };
  const reorderPage = useCallback((pageId: string, toIndex: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    try { handle.reorderPage(pageId, toIndex); refresh(toIndex, true); } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const toggleLayerVisible = useCallback((index: number, visible: boolean) => {
    const handle = handleRef.current; const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !page) return;
    try { handle.setLayerVisible(page.sourcePartPath, index, visible); refresh(); } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const fitToWindow = useCallback(() => {
    const frame = modelRef.current.frame; const workspace = workspaceRef.current;
    if (!frame || !workspace || frame.width <= 0 || frame.height <= 0) return;
    const rect = workspace.getBoundingClientRect();
    setZoom(clampZoom(Math.min((rect.width - WORKSPACE_MARGIN) / frame.width, (rect.height - WORKSPACE_MARGIN) / frame.height)));
  }, []);
  const download = useCallback((bytes: Uint8Array) => { const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer; const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.visio' })); const anchor = document.createElement('a'); anchor.href = url; anchor.download = 'diagram.vsdx'; anchor.click(); URL.revokeObjectURL(url); setDirty(false); }, []);
  const editOverlay = editing && model.frame && editedTextId ? textEditOverlay(model.frame, editedTextId, zoom) : null;
  const integrity = diagnostics.filter((diagnostic) => diagnostic.category === 'integrity');
  const fidelity = diagnostics.filter((diagnostic) => diagnostic.category === 'fidelity');
  return <div className={className} style={styles.root} aria-label={t('editor.appLabel')}>
    <header style={styles.titleBar}><strong>{t('ribbon.documentName')}</strong><span style={{ color: dirty ? '#a16207' : '#526273' }}>{dirty ? t('ribbon.dirty') : t('ribbon.saved')}</span></header>
    <RibbonCommandsProvider handle={handleRef.current} snapshot={model.snapshot} pageId={model.snapshot?.pages[model.pageIndex]?.id} selection={selection} onMutation={() => refresh(undefined, true)} onError={reportError} onDownload={download}>
    <RibbonCommandsBridge target={commandsRef} />
    <Ribbon t={t} />
    <div style={styles.contentRow}>
    {leftPanel === undefined ? (
      <div style={styles.leftColumn}>
        <div style={styles.shapesWrap}>
          <ShapesPanel shapes={standardShapes} collapsed={shapesCollapsed} onToggleCollapsed={() => setShapesCollapsed((value) => !value)} onInsert={insertShape} t={t} />
        </div>
        <LayersPanel layers={model.layers} collapsed={layersCollapsed} onToggleCollapsed={() => setLayersCollapsed((value) => !value)} onToggleLayer={toggleLayerVisible} t={t} />
      </div>
    ) : leftPanel}
    <main ref={workspaceRef} style={styles.workspace}>
      {loading && <span>{t('editor.opening')}</span>}
      {!loading && !model.frame && <span>{file ? t('editor.noPages') : t('editor.openPrompt')}</span>}
      <div style={styles.canvasFrame} onDragOver={onCanvasDragOver} onDrop={onCanvasDrop}>
        <canvas ref={mainCanvasRef} tabIndex={0} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerCancel} onLostPointerCapture={onLostPointerCapture} onDoubleClick={onCanvasDoubleClick} onContextMenu={onCanvasContextMenu} onKeyDown={onCanvasKeyDown} onFocus={onCanvasFocus} onBlur={onCanvasBlur} aria-label={selection ? t('pages.canvasLabelWithSelection', { current: model.pageIndex + 1, total: model.snapshot?.pages.length ?? 0, name: selection.shapeId }) : t('pages.canvasLabel', { current: model.pageIndex + 1, total: model.snapshot?.pages.length ?? 0 })} style={styles.canvas} />
        <canvas ref={overlayCanvasRef} aria-hidden="true" style={styles.overlay} />
        {editing && editOverlay && (
          <div ref={editWrapRef} style={{ ...styles.textEditWrap, width: editOverlay.width, height: editOverlay.height, transform: `matrix(${editOverlay.matrix.a}, ${editOverlay.matrix.b}, ${editOverlay.matrix.c}, ${editOverlay.matrix.d}, ${editOverlay.matrix.e}, ${editOverlay.matrix.f})` }}>
            <textarea
              ref={editBoxRef}
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); commitTextEdit(); mainCanvasRef.current?.focus(); } else event.stopPropagation(); }}
              aria-label={t('shapes.editingText', { name: editing.shapeId })}
              rows={1}
              style={{
                ...styles.textEditBox,
                fontFamily: `"${editOverlay.font.family}", sans-serif`,
                fontSize: editOverlay.font.sizePx,
                fontWeight: editOverlay.font.bold ? 700 : 400,
                fontStyle: editOverlay.font.italic ? 'italic' : 'normal',
                color: editOverlay.font.color,
              }}
            />
          </div>
        )}
      </div>
      {contextMenu && selection && <ShapeContextMenu t={t} position={contextMenu} onClose={closeContextMenu} onCloseAndFocus={closeContextMenuAndFocus} />}
      {integrity.length > 0 && <section role="alert" style={styles.integrity}><strong>{t('diagnostics.integrityHeading')}</strong>{integrity.map((item, index) => <div key={`${item.code}-${index}`}>{diagnosticMessage(t, item.category, item.code)}</div>)}</section>}
      {fidelity.length > 0 && <details style={styles.fidelity}><summary>{t('diagnostics.fidelityHeading')}</summary>{fidelity.map((item, index) => <div key={`${item.code}-${index}`}>{diagnosticMessage(t, item.category, item.code)}</div>)}</details>}
      {error && <div role="alert" style={styles.error}>{error}</div>}
    </main>
    </div>
    {statusBar === undefined ? <StatusBar pages={model.snapshot?.pages ?? []} activeIndex={model.pageIndex} onSelectPage={(index) => refresh(index)} onReorderPage={reorderPage} zoom={zoom} onZoomChange={setZoom} onFitToWindow={fitToWindow} t={t} /> : statusBar}
    </RibbonCommandsProvider>
  </div>;
}

const WORKSPACE_MARGIN = 32;

/** Page layers for the panel; empty when the handle predates layer support. */
function readPageLayers(handle: DiagramHandle, pageIndex: number): PageLayer[] {
  try {
    return typeof handle.pageLayers === 'function' ? handle.pageLayers(pageIndex) : [];
  } catch {
    return [];
  }
}

/** Cascade step for repeated centre inserts, in inches. */
export const CENTRE_INSERT_STEP_IN = 0.25;
/** Cascade length before a centre insert wraps back. */
export const CENTRE_INSERT_CASCADE = 8;

/** Offset a page-centre insert so repeated clicks cascade instead of stacking. */
export function centreInsertPoint(centre: ModelPoint, count: number): ModelPoint {
  const step = count % CENTRE_INSERT_CASCADE;
  return { x: centre.x + step * CENTRE_INSERT_STEP_IN, y: centre.y - step * CENTRE_INSERT_STEP_IN };
}

interface ClientRectLike { left: number; top: number; width: number; height: number; }

/** Map a client point onto canvas pixels and Y-up inches, dividing out zoom once through the rendered rect. */
export function clientPointToModel(frame: PageDisplayList, rect: ClientRectLike, clientX: number, clientY: number): { canvas: ModelPoint; model: ModelPoint } {
  const canvas = {
    x: (clientX - rect.left) * frame.width / Math.max(rect.width, 1),
    y: (clientY - rect.top) * frame.height / Math.max(rect.height, 1),
  };
  return { canvas, model: canvasPointToModel(frame.paintTransform, canvas.x, canvas.y) };
}

export function canvasPointerPosition(event: PointerEvent<HTMLCanvasElement> | MouseEvent<HTMLCanvasElement>, frame: PageDisplayList): { canvas: ModelPoint; model: ModelPoint } {
  return clientPointToModel(frame, event.currentTarget.getBoundingClientRect(), event.clientX, event.clientY);
}

export function inchFormula(value: number): string {
  if (!Number.isFinite(value)) throw new Error('Shape geometry must be finite.');
  const rounded = Number(value.toFixed(6));
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

export function shapeParentTransforms(primitives: readonly PagePrimitive[], id: string, depth = 0): Affine[] | null {
  if (depth >= 256) return null;
  for (const primitive of primitives) {
    if (primitive.id === id) return [];
    if (primitive.kind === 'group') {
      const nested = shapeParentTransforms(primitive.primitives, id, depth + 1);
      if (nested) return primitive.transform ? [primitive.transform, ...nested] : nested;
    }
  }
  return null;
}

/** Visio selects the outermost shape a hit falls in; only a top-level shape carries page-space bounds. */
function selectionForHit(page: PageSnapshot, hit: HitTestResult): VsdxShapeSelection | null {
  const top = page.shapes.find((shape) => shape.id === hit.shapeId || findShapePlacement(shape.children, hit.shapeId) !== null);
  return top ? { pageId: page.id, shapeId: top.id, hit } : null;
}

function textPrimitiveId(snapshot: DiagramSnapshot, pageIndex: number, pageId: string, shapeId: string): string | null {
  const page = snapshot.pages[pageIndex];
  if (!page || page.id !== pageId) return null;
  const placement = findShapePlacement(page.shapes, shapeId);
  return placement ? `${page.sourcePartPath}:${placement.shape.sourceId}` : null;
}

export function stillSelectable(snapshot: DiagramSnapshot, pageIndex: number, selection: VsdxShapeSelection, layers?: readonly PageLayer[]): boolean {
  const page = snapshot.pages[pageIndex];
  if (!page || page.id !== selection.pageId || !findShapePlacement(page.shapes, selection.shapeId)) return false;
  return !layers || !selectionHiddenByLayers(page, layers, selection);
}

export function selectionHiddenByLayers(page: PageSnapshot, layers: readonly PageLayer[], selection: VsdxShapeSelection): boolean {
  return shapeSubtreeHidden(page.shapes, layers, selection.shapeId, false) ?? false;
}

function shapeSubtreeHidden(shapes: readonly ShapeSnapshot[], layers: readonly PageLayer[], shapeId: string, ancestorHidden: boolean): boolean | null {
  for (const shape of shapes) {
    const hiddenHere = ancestorHidden || shapeHiddenByLayers(shape, layers);
    if (shape.id === shapeId) return hiddenHere;
    const nested = shapeSubtreeHidden(shape.children, layers, shapeId, hiddenHere);
    if (nested !== null) return nested;
  }
  return null;
}

function shapeHiddenByLayers(shape: ShapeSnapshot, layers: readonly PageLayer[]): boolean {
  const indices = layerMemberIndices(shape);
  return indices.length > 0 && indices.every((index) => layers.some((layer) => layer.index === index && !layer.visible));
}

function layerMemberIndices(shape: ShapeSnapshot): number[] {
  const cell = shape.cells.find((entry) => entry.name === 'LayerMember');
  const raw = cell?.value ?? cell?.formula ?? '';
  const indices = raw.split(';').map((part) => part.trim()).filter((part) => /^\+?\d+$/.test(part)).map((part) => Number(part)).filter((index) => index <= 4294967295);
  return [...new Set(indices)].sort((left, right) => left - right);
}

function shapeDragStart(page: { id: string; shapes: readonly ShapeSnapshot[]; sourcePartPath: string }, shape: ShapeSnapshot, frame: PageDisplayList, handle?: DiagramHandle): Omit<DragStart, 'canvas' | 'model' | 'resize' | 'pointerId' | 'startX' | 'startY'> {
  const width = numericCellValue(shape, 'Width');
  const height = numericCellValue(shape, 'Height');
  return {
    parentTransforms: shapeParentTransforms(frame.primitives, `${page.sourcePartPath}:${shape.sourceId}`) ?? [],
    angle: numericCellValue(shape, 'Angle', 0),
    flipX: numericCellValue(shape, 'FlipX', 0) === 1,
    flipY: numericCellValue(shape, 'FlipY', 0) === 1,
    pin: { x: numericCellValue(shape, 'PinX'), y: numericCellValue(shape, 'PinY') },
    locPin: { x: numericCellValue(shape, 'LocPinX', width / 2), y: numericCellValue(shape, 'LocPinY', height / 2) },
    locPinAtSize: handle ? (nextWidth, nextHeight) => handle.resizeLocPin(page.id, shape.id, nextWidth, nextHeight) : undefined,
    size: { width, height },
  };
}

export function selectionCorners(page: PageSnapshot, frame: PageDisplayList, selection: VsdxShapeSelection): ModelPoint[] | null {
  const placement = findShapePlacement(page.shapes, selection.shapeId);
  if (!placement) return null;
  const start: DragStart = {
    canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false,
    ...shapeDragStart(page, placement.shape, frame),
  };
  return previewOutline(start, { x: 0, y: 0 }, frame.paintTransform);
}

export function collectDiagnostics(frame: PageDisplayList): TextDiagnostic[] { const result: TextDiagnostic[] = []; const work = frame.primitives.map((primitive) => ({ primitive, depth: 0 })); while (work.length) { const current = work.pop(); if (!current || current.depth >= 256) continue; if (current.primitive.kind === 'shape') result.push(...(current.primitive.diagnostics ?? [])); if (current.primitive.kind === 'textBox') for (const paragraph of current.primitive.paragraphs) for (const run of paragraph.runs) result.push(...(run.diagnostics ?? [])); if (current.primitive.kind === 'group') for (const primitive of current.primitive.primitives) work.push({ primitive, depth: current.depth + 1 }); } return result; }
/** Latest ribbon commands for the canvas keyboard layer, which lives outside the provider. */
function RibbonCommandsBridge({ target }: { target: { current: RibbonCommands | null } }) {
  const commands = useRibbonCommands();
  useEffect(() => { target.current = commands; }, [commands, target]);
  target.current = commands;
  return null;
}
interface LoadedFont { key: string; face: FontFace; }
async function loadFonts(fonts: ReadonlyArray<VsdxFontFace>): Promise<LoadedFont[]> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return [];
  return Promise.all(fonts.map(async (font) => {
    const source = font.bytes.slice().buffer as ArrayBuffer;
    const face = await new FontFace(font.family, source, { style: font.italic ? 'italic' : 'normal', weight: font.bold ? '700' : '400' }).load();
    return { key: JSON.stringify([font.family, font.bold ?? false, font.italic ?? false]), face };
  }));
}
function installFonts(fonts: readonly LoadedFont[], installed: Map<string, FontFace>): void {
  for (const { key, face } of fonts) {
    const previous = installed.get(key);
    if (previous) document.fonts.delete?.(previous);
    document.fonts.add(face);
    installed.set(key, face);
  }
}
function useStableInitialUpdate(update: Uint8Array | undefined): Uint8Array | undefined { const stable = useRef(update); if (stable.current !== update && (!stable.current || !update || !bytesEqual(stable.current, update))) stable.current = update; return stable.current; }
function useStableFontFaces(fonts: ReadonlyArray<VsdxFontFace>): ReadonlyArray<VsdxFontFace> { const stable = useRef(fonts); if (!fontFacesEqual(stable.current, fonts)) stable.current = fonts; return stable.current; }
function fontFacesEqual(left: ReadonlyArray<VsdxFontFace>, right: ReadonlyArray<VsdxFontFace>): boolean { return left === right || (left.length === right.length && left.every((face, index) => fontFaceEqual(face, right[index]))); }
function fontFaceKeyEqual(left: VsdxFontFace, right: VsdxFontFace): boolean { return left.family === right.family && (left.bold ?? false) === (right.bold ?? false) && (left.italic ?? false) === (right.italic ?? false); }
function fontFaceEqual(left: VsdxFontFace, right: VsdxFontFace): boolean { return fontFaceKeyEqual(left, right) && bytesEqual(left.bytes, right.bytes); }
function bytesEqual(left: Uint8Array, right: Uint8Array): boolean { return left === right || (left.byteLength === right.byteLength && left.every((byte, index) => byte === right[index])); }
function resolveImage(assetId: string, handle: DiagramHandle | null, cache: { current: Map<string, Promise<CanvasImageSource | null>> }, message: string): Promise<CanvasImageSource | null> { const existing = cache.current.get(assetId); if (existing) return existing; const pending = decodeImage(handle?.mediaBytes(assetId), message); cache.current.set(assetId, pending); return pending; }
async function decodeImage(bytes: Uint8Array | undefined, message: string): Promise<CanvasImageSource | null> { if (!bytes) return null; const blob = new Blob([bytes.slice()]); if (typeof createImageBitmap === 'function') return createImageBitmap(blob); const url = URL.createObjectURL(blob); try { return await new Promise<HTMLImageElement>((resolve, reject) => { const image = new Image(); image.onload = () => resolve(image); image.onerror = () => reject(new Error(message)); image.src = url; }); } finally { URL.revokeObjectURL(url); } }
const styles: Record<string, CSSProperties> = { root: { display: 'flex', flexDirection: 'column', width: '100%', height: '100%', minHeight: 480, color: '#172033', background: '#f3f5f8', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }, titleBar: { display: 'flex', alignItems: 'center', gap: 12, minHeight: 32, padding: '0 14px', background: '#f8fafc', borderBottom: '1px solid #d8dee9', fontSize: 13 }, contentRow: { display: 'flex', flex: 1, minHeight: 0 }, leftColumn: { display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 }, shapesWrap: { display: 'flex', flex: '1 1 auto', minHeight: 0 }, workspace: { position: 'relative', display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', overflow: 'auto' }, canvasFrame: { position: 'relative', flex: '0 0 auto' }, canvas: { display: 'block', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' }, overlay: { position: 'absolute', inset: 0, pointerEvents: 'none' }, textEditWrap: { position: 'absolute', left: 0, top: 0, transformOrigin: '0 0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'visible', background: 'transparent', border: '1px dashed #1d4ed8', zIndex: 2 }, textEditBox: { width: '100%', background: 'transparent', border: 'none', outline: 'none', resize: 'none', overflow: 'visible', textAlign: 'center', whiteSpace: 'pre-wrap', overflowWrap: 'break-word', wordBreak: 'break-word', lineHeight: 1.2, padding: 0, margin: 0 }, integrity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 12, color: '#7f1d1d', background: '#fef2f2', border: '1px solid #fca5a5' }, fidelity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 8, color: '#475569', background: '#fff', fontSize: 12 }, error: { position: 'absolute', left: 14, right: 14, bottom: 14, padding: 10, color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0' } };
