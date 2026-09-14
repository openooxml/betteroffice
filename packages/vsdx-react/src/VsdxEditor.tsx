import { createT, deepMerge, diagnosticMessage, en } from '@betteroffice/vsdx-i18n';
import type { Translations } from '@betteroffice/vsdx-i18n';
import { canvasPointToModel, initWasm, openDiagram, paintPage, sizeCanvasForPage } from '@betteroffice/vsdx';
import type { Affine, PagePrimitive, CollaborationReplica, DiagramHandle, DiagramSnapshot, HitTestResult, ModelPoint, PageDisplayList, PageSnapshot, ShapeSnapshot, TextDiagnostic, VsdxFontFace, VsdxPresence } from '@betteroffice/vsdx';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, FocusEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode } from 'react';
import { Ribbon } from './components/ribbon/Ribbon';
import { ShapeContextMenu } from './components/ribbon/ShapeContextMenu';
import { RibbonCommandsProvider, findShapePlacement, isHandleResizeBlocked, numericCellValue, useRibbonCommands } from './components/ribbon/commands';
import type { RibbonCommands } from './components/ribbon/commands';
import { ShapesPanel } from './components/shapes/ShapesPanel';
import { standardShapes } from './components/shapes/shapeLibrary';
import type { StandardShape } from './components/shapes/shapeLibrary';
import { StatusBar, clampZoom } from './components/statusbar';
import { paintDragPreview, paintSelectionFrame, passedDragThreshold, previewOutline, hitTestSelection, resolveDragGeometry, resolveNudgeGeometry, resolveRotationAngle, resizeCursor, canvasKeyboardIntent } from './interactions';
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

interface EditorModel { snapshot: DiagramSnapshot | null; pageIndex: number; frame: PageDisplayList | null; }

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
  const [model, setModel] = useState<EditorModel>({ snapshot: null, pageIndex: 0, frame: null });
  const modelRef = useRef(model);
  const [selection, setSelection] = useState<VsdxShapeSelection | null>(null);
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  const [dirty, setDirty] = useState(false);
  const sessionSwitchBlocked = dirty && sessionRef.current.file === file &&
    (sessionRef.current.clientId !== requestedClientId || sessionRef.current.initialUpdate !== requestedInitialUpdate);
  const sessionClientId = sessionSwitchBlocked ? sessionRef.current.clientId : requestedClientId;
  const initialUpdate = sessionSwitchBlocked ? sessionRef.current.initialUpdate : requestedInitialUpdate;
  const [zoom, setZoom] = useState(1);
  const [shapesCollapsed, setShapesCollapsed] = useState(false);
  const [diagnostics, setDiagnostics] = useState<TextDiagnostic[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [contextMenu, setContextMenu] = useState<{ top: number; left: number } | null>(null);
  const pointerRef = useRef<DragStart | null>(null);
  const dragPreviewRef = useRef<ModelPoint | null>(null);
  const previewFrameRef = useRef<number | null>(null);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const [loading, setLoading] = useState(Boolean(file));
  onReadyRef.current = onReady;
  onChangeRef.current = onChange;
  onErrorRef.current = onError;
  collaborationRef.current = collaboration;

  const reportError = useCallback((value: unknown) => { const next = value instanceof Error ? value : new Error(String(value)); setError(next.message); onErrorRef.current?.(next); }, []);
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
      modelRef.current = { snapshot: current, pageIndex, frame };
      setModel(modelRef.current);
      setDiagnostics(frame ? collectDiagnostics(frame) : []);
      setSelection((existing) => existing && stillSelectable(current, pageIndex, existing) ? existing : null);
    } catch (value) { reportError(value); }
  }, [reportError]);

  useEffect(() => {
    sessionRef.current = { file, clientId: sessionClientId, initialUpdate };
    let disposed = false;
    let handle: DiagramHandle | null = null;
    let stopUpdates = () => {};
    let stopResync = () => {};
    handleRef.current?.dispose(); handleRef.current = null; imageCache.current.clear(); setSelection(null); modelRef.current = { snapshot: null, pageIndex: 0, frame: null }; setModel(modelRef.current); setError(null); setDirty(false);
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

  useEffect(() => {
    const canvas = mainCanvasRef.current; const frame = model.frame;
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
  }, [model.frame, reportError, t, zoom]);

  useEffect(() => {
    const canvas = overlayCanvasRef.current; const frame = model.frame;
    if (!canvas || !frame) return;
    const context = canvas.getContext('2d'); if (!context) return;
    const dpr = window.devicePixelRatio || 1; sizeCanvasForPage(canvas, frame, dpr, zoom); context.clearRect(0, 0, canvas.width, canvas.height);
    const snapshot = model.snapshot;
    const page = snapshot?.pages[model.pageIndex];
    if (selection && page) {
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
    if (currentSelection && page) {
      try {
        const corners = selectionCorners(page, frame, currentSelection);
        const placement = findShapePlacement(page.shapes, currentSelection.shapeId);
        const blocked = placement ? isHandleResizeBlocked(placement.shape) : false;
        if (corners) paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current, blocked ? [] : undefined);
      } catch { void 0; }
    }
  };

  const dragStartForPlacement = (page: { shapes: readonly ShapeSnapshot[]; sourcePartPath: string }, shape: ShapeSnapshot, frame: PageDisplayList): Omit<DragStart, 'canvas' | 'model' | 'resize' | 'pointerId' | 'startX' | 'startY'> => {
    const width = numericCellValue(shape, 'Width');
    const height = numericCellValue(shape, 'Height');
    return {
      parentTransforms: shapeParentTransforms(frame.primitives, `${page.sourcePartPath}:${shape.sourceId}`) ?? [],
      angle: numericCellValue(shape, 'Angle', 0),
      flipX: numericCellValue(shape, 'FlipX', 0) === 1,
      flipY: numericCellValue(shape, 'FlipY', 0) === 1,
      pin: { x: numericCellValue(shape, 'PinX'), y: numericCellValue(shape, 'PinY') },
      locPin: { x: numericCellValue(shape, 'LocPinX', width / 2), y: numericCellValue(shape, 'LocPinY', height / 2) },
      size: { width, height },
    };
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
      if (active && active.pageId === page.id) {
        try {
          const corners = selectionCorners(page, frame, active);
          if (corners) {
            const target = hitTestSelection(point.canvas, corners, zoomRef.current);
            if (target) {
              const placement = findShapePlacement(page.shapes, active.shapeId);
              if (placement) {
                if (target !== 'rotate' && isHandleResizeBlocked(placement.shape)) {
                  reportError(new Error('Shape is locked and cannot be resized with handles.'));
                  return;
                }
                const base = dragStartForPlacement(page, placement.shape, frame);
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
      setSelection(hit ? { pageId: page.id, shapeId: hit.shapeId, hit } : null);
      const placement = hit ? findShapePlacement(page.shapes, hit.shapeId) : null;
      pointerRef.current = hit && placement ? {
        ...point,
        ...dragStartForPlacement(page, placement.shape, frame),
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
        if (!frame || !page || !active || active.pageId !== page.id) { event.currentTarget.style.cursor = ''; return; }
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
      const snap = event.shiftKey;
      dragPreviewRef.current = point.model;
      if (previewFrameRef.current !== null) return;
      previewFrameRef.current = requestAnimationFrame(() => {
        previewFrameRef.current = null;
        const liveFrame = modelRef.current.frame; const liveStart = pointerRef.current; const release = dragPreviewRef.current; const overlay = overlayCanvasRef.current;
        if (!liveFrame || !liveStart || !release || !overlay) return;
        const context = overlay.getContext('2d'); if (!context) return;
        try {
          const corners = previewOutline(liveStart, release, liveFrame.paintTransform, snap);
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
        if (livePlacement && isHandleResizeBlocked(livePlacement.shape)) throw new Error('Shape is locked and cannot be resized with handles.');
        handle.transformShape(selected.pageId, selected.shapeId, inchFormula(geometry.x), inchFormula(geometry.y), inchFormula(geometry.width), inchFormula(geometry.height));
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
  const onCanvasContextMenu = (event: MouseEvent<HTMLCanvasElement>) => {
    event.preventDefault();
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    if (pointerRef.current) return;
    try {
      const point = canvasPointerPosition(event, frame);
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      if (!hit) return;
      const active = selectionRef.current;
      if (!active || active.pageId !== page.id || active.shapeId !== hit.shapeId) setSelection({ pageId: page.id, shapeId: hit.shapeId, hit });
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
      if (!frame) {
        handle.moveShape(selected.pageId, selected.shapeId, inchFormula(numericCellValue(placement.shape, 'PinX') + dx), inchFormula(numericCellValue(placement.shape, 'PinY') + dy));
        refresh(undefined, true);
        return;
      }
      const base = dragStartForPlacement(page, placement.shape, frame);
      const geometry = resolveNudgeGeometry({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, ...base }, dx, dy);
      handle.moveShape(selected.pageId, selected.shapeId, inchFormula(geometry.x), inchFormula(geometry.y));
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  };
  const onCanvasKeyDown = (event: KeyboardEvent<HTMLCanvasElement>) => {
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
  const insertShape = useCallback((shape: StandardShape) => {
    const handle = handleRef.current; const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !page || !frame) return;
    try {
      const centre = canvasPointToModel(frame.paintTransform, frame.width / 2, frame.height / 2);
      handle.addShape(page.id, shape.draft(centre.x, centre.y, 1, 1));
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const reorderPage = useCallback((pageId: string, toIndex: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    try { handle.reorderPage(pageId, toIndex); refresh(toIndex, true); } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const fitToWindow = useCallback(() => {
    const frame = modelRef.current.frame; const workspace = workspaceRef.current;
    if (!frame || !workspace || frame.width <= 0 || frame.height <= 0) return;
    const rect = workspace.getBoundingClientRect();
    setZoom(clampZoom(Math.min((rect.width - WORKSPACE_MARGIN) / frame.width, (rect.height - WORKSPACE_MARGIN) / frame.height)));
  }, []);
  const download = useCallback((bytes: Uint8Array) => { const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer; const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.visio' })); const anchor = document.createElement('a'); anchor.href = url; anchor.download = 'diagram.vsdx'; anchor.click(); URL.revokeObjectURL(url); setDirty(false); }, []);
  const integrity = diagnostics.filter((diagnostic) => diagnostic.category === 'integrity');
  const fidelity = diagnostics.filter((diagnostic) => diagnostic.category === 'fidelity');
  return <div className={className} style={styles.root} aria-label={t('editor.appLabel')}>
    <header style={styles.titleBar}><strong>{t('ribbon.documentName')}</strong><span style={{ color: dirty ? '#a16207' : '#526273' }}>{dirty ? t('ribbon.dirty') : t('ribbon.saved')}</span></header>
    <RibbonCommandsProvider handle={handleRef.current} snapshot={model.snapshot} pageId={model.snapshot?.pages[model.pageIndex]?.id} selection={selection} onMutation={() => refresh(undefined, true)} onError={reportError} onDownload={download}>
    <RibbonCommandsBridge target={commandsRef} />
    <Ribbon t={t} />
    <div style={styles.contentRow}>
    {leftPanel === undefined ? <ShapesPanel shapes={standardShapes} collapsed={shapesCollapsed} onToggleCollapsed={() => setShapesCollapsed((value) => !value)} onInsert={insertShape} t={t} /> : leftPanel}
    <main ref={workspaceRef} style={styles.workspace}>
      {loading && <span>{t('editor.opening')}</span>}
      {!loading && !model.frame && <span>{file ? t('editor.noPages') : t('editor.openPrompt')}</span>}
      <div style={styles.canvasFrame}>
        <canvas ref={mainCanvasRef} tabIndex={0} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerCancel} onLostPointerCapture={onLostPointerCapture} onContextMenu={onCanvasContextMenu} onKeyDown={onCanvasKeyDown} onFocus={onCanvasFocus} onBlur={onCanvasBlur} aria-label={selection ? t('pages.canvasLabelWithSelection', { current: model.pageIndex + 1, total: model.snapshot?.pages.length ?? 0, name: selection.shapeId }) : t('pages.canvasLabel', { current: model.pageIndex + 1, total: model.snapshot?.pages.length ?? 0 })} style={styles.canvas} />
        <canvas ref={overlayCanvasRef} aria-hidden="true" style={styles.overlay} />
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

export function canvasPointerPosition(event: PointerEvent<HTMLCanvasElement> | MouseEvent<HTMLCanvasElement>, frame: PageDisplayList): { canvas: ModelPoint; model: ModelPoint } {
  const rect = event.currentTarget.getBoundingClientRect();
  const canvas = {
    x: (event.clientX - rect.left) * frame.width / Math.max(rect.width, 1),
    y: (event.clientY - rect.top) * frame.height / Math.max(rect.height, 1),
  };
  return { canvas, model: canvasPointToModel(frame.paintTransform, canvas.x, canvas.y) };
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

export function stillSelectable(snapshot: DiagramSnapshot, pageIndex: number, selection: VsdxShapeSelection): boolean {
  const page = snapshot.pages[pageIndex];
  return Boolean(page && page.id === selection.pageId && findShapePlacement(page.shapes, selection.shapeId));
}

/** Current selection corners in scale-1 canvas coordinates for overlay paint and hit tests. */
export function selectionCorners(page: PageSnapshot, frame: PageDisplayList, selection: VsdxShapeSelection): ModelPoint[] | null {
  const placement = findShapePlacement(page.shapes, selection.shapeId);
  if (!placement) return null;
  const shape: ShapeSnapshot = placement.shape;
  const width = numericCellValue(shape, 'Width');
  const height = numericCellValue(shape, 'Height');
  const start: DragStart = {
    canvas: { x: 0, y: 0 },
    model: { x: 0, y: 0 },
    resize: false,
    pin: { x: numericCellValue(shape, 'PinX'), y: numericCellValue(shape, 'PinY') },
    locPin: { x: numericCellValue(shape, 'LocPinX', width / 2), y: numericCellValue(shape, 'LocPinY', height / 2) },
    size: { width, height },
    parentTransforms: shapeParentTransforms(frame.primitives, `${page.sourcePartPath}:${shape.sourceId}`) ?? [],
    angle: numericCellValue(shape, 'Angle', 0),
    flipX: numericCellValue(shape, 'FlipX', 0) === 1,
    flipY: numericCellValue(shape, 'FlipY', 0) === 1,
  };
  return previewOutline(start, { x: 0, y: 0 }, frame.paintTransform);
}

export function collectDiagnostics(frame: PageDisplayList): TextDiagnostic[] { const result: TextDiagnostic[] = []; const work = frame.primitives.map((primitive) => ({ primitive, depth: 0 })); while (work.length) { const current = work.pop(); if (!current || current.depth >= 256) continue; if (current.primitive.kind === 'textBox') for (const paragraph of current.primitive.paragraphs) for (const run of paragraph.runs) result.push(...run.diagnostics); if (current.primitive.kind === 'group') for (const primitive of current.primitive.primitives) work.push({ primitive, depth: current.depth + 1 }); } return result; }
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
const styles: Record<string, CSSProperties> = { root: { display: 'flex', flexDirection: 'column', width: '100%', height: '100%', minHeight: 480, color: '#172033', background: '#f3f5f8', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }, titleBar: { display: 'flex', alignItems: 'center', gap: 12, minHeight: 32, padding: '0 14px', background: '#f8fafc', borderBottom: '1px solid #d8dee9', fontSize: 13 }, contentRow: { display: 'flex', flex: 1, minHeight: 0 }, workspace: { position: 'relative', display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', overflow: 'auto' }, canvasFrame: { position: 'relative', flex: '0 0 auto' }, canvas: { display: 'block', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' }, overlay: { position: 'absolute', inset: 0, pointerEvents: 'none' }, integrity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 12, color: '#7f1d1d', background: '#fef2f2', border: '1px solid #fca5a5' }, fidelity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 8, color: '#475569', background: '#fff', fontSize: 12 }, error: { position: 'absolute', left: 14, right: 14, bottom: 14, padding: 10, color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0' } };
