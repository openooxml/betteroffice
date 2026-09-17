import { createT, deepMerge, diagnosticMessage, en } from '@betteroffice/vsdx-i18n';
import type { TFunction, Translations } from '@betteroffice/vsdx-i18n';
import { canvasPointToModel, modelPointToCanvas, initWasm, openDiagram, paintPage, sizeCanvasForPage } from '@betteroffice/vsdx';
import type { Affine, CellLocator, DocumentMaster, PageLayer, PagePrimitive, CollaborationReplica, DiagramHandle, DiagramSnapshot, HitTestResult, ModelPoint, PageDisplayList, PageSnapshot, ShapeDataRow, ShapeMove, ShapeSnapshot, TextDiagnostic, VsdxFontFace, VsdxPresence } from '@betteroffice/vsdx';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, DragEvent, FocusEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode } from 'react';
import { AUTO_CONNECT_FADE_MS, HOVER_FREE_DRAG_INCHES, HOVER_PROXIMITY_PX, QUICK_SHAPE_IDS, autoConnectArrowAt, autoConnectArrowCss, autoConnectArrowsForShape, autoConnectHaloHit, connectionPointsForShape, connectorDraft, connectorEndpointGlue, connectorGlue, connectorRouteFromFrame, dropTargetForPoint, hoverPointAt, isConnectorShape, nearestConnectionPointAnywhere, paintAutoConnectOverlay, paintConnectorOverlay, quickShapePlacement, reroutePreviewForMove, routeConnector } from './connector';
import type { AutoConnectSide, ConnectionPoint, ConnectorDragEndpoint, ConnectorOverlayRoute, ConnectorOverlayScene } from './connector';
import { Ribbon } from './components/ribbon/Ribbon';
import type { ViewToggleKey } from './components/ribbon/Ribbon';
import { CanvasContextMenu } from './components/ribbon/CanvasContextMenu';
import { ShapeContextMenu } from './components/ribbon/ShapeContextMenu';
import { RibbonCommandsProvider, copySelection, duplicateEntry, findShapePlacement, isCellWriteBlocked, isHandleResizeBlocked, isRotateBlocked, numericCellValue, pasteEntry, useRibbonCommands } from './components/ribbon/commands';
import type { RibbonCommands } from './components/ribbon/commands';
import { dragSegmentRoute, hitSegmentDot, paintConnectorChrome, previewChrome, selectedConnectorChrome } from './connectorChrome';
import type { ChromePoint } from './connectorChrome';
import type { VsdxClipboardEntry } from './components/ribbon/clipboard';
import { DUPLICATE_OFFSET, PASTE_OFFSET } from './components/ribbon/clipboard';
import { STENCIL_DRAG_MIME, ShapesPanel } from './components/shapes/ShapesPanel';
import { RulerLeft, RulerTop, RULER_SIZE } from './components/ruler/Rulers';
import { LayersPanel } from './components/layers/LayersPanel';
import { ShapeDataPanel } from './components/shapeData/ShapeDataPanel';
import { DrawingExplorer } from './components/explorer/DrawingExplorer';
import { shapeStencils, stencilShapeById } from './components/shapes/shapeLibrary';
import type { ShapeStencil, StandardShape } from './components/shapes/shapeLibrary';
import { documentStencilEntries } from './components/shapes/documentStencil';
import { StatusBar, clampZoom } from './components/statusbar';
import { paintDragPreview, paintMarquee, paintSelectionFrame, passedDragThreshold, previewOutline, hitTestSelection, isOwnedBrowserShortcut, hitTestControlHandles, controlCellWriteBlocked, controlHandleCanvasPositions, controlHandlesForShape, isPrintableEntryKey, marqueeEnclosesQuad, normalizeMarquee, paintControlHandles, resolveControlDrag, resolveDragGeometry, resolveNudgeGeometry, resolveRotationAngle, resizeCursor, canvasKeyboardIntent, shapeLocalToPage, textEditOverlay, withoutTextBox } from './interactions';
import type { CanvasKeyboardIntent, ControlDrag, DragStart, MarqueeRect, ResizeHandle } from './interactions';
import { collectSnapTargets, paintGrid, paintSmartGuides, snapRelease, GRID_SPACING_IN } from './snap';
import type { SnapTargets } from './snap';
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
  rightPanel?: ReactNode;
  statusBar?: ReactNode;
}

interface EditorModel { snapshot: DiagramSnapshot | null; pageIndex: number; frame: PageDisplayList | null; layers: PageLayer[]; }

export function VsdxEditor({ file, fonts, clientId, collaboration, i18n, className, onReady, onChange, onError, leftPanel, rightPanel, statusBar }: VsdxEditorProps) {
  const strings = useMemo(() => deepMerge(en, i18n) as typeof en, [i18n]);
  const t = useMemo(() => createT(strings), [strings]);
  const handleRef = useRef<DiagramHandle | null>(null);
  const onReadyRef = useRef(onReady);
  const onChangeRef = useRef(onChange);
  const onErrorRef = useRef(onError);
  const collaborationRef = useRef(collaboration);
  const attachedCollaborationRef = useRef<VsdxEditorCollaborationOptions | undefined>(undefined);
  const editorRootRef = useRef<HTMLDivElement>(null);
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
  const [selection, setSelection] = useState<VsdxShapeSelection[]>([]);
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  const [clipboard, setClipboard] = useState<VsdxClipboardEntry | null>(null);
  const clipboardRef = useRef(clipboard);
  clipboardRef.current = clipboard;
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
  const [showGrid, setShowGrid] = useState(true);
  const [snapEnabled, setSnapEnabled] = useState(true);
  const [showRulers, setShowRulers] = useState(true);
  const [rulerMark, setRulerMark] = useState<ModelPoint | null>(null);
  const rulerMarkRef = useRef<ModelPoint | null>(null);
  const rulerFrameRef = useRef<number | null>(null);
  const snappedReleaseRef = useRef<{ raw: ModelPoint; point: ModelPoint } | null>(null);
  const [shapesCollapsed, setShapesCollapsed] = useState(false);
  const [documentStencil, setDocumentStencil] = useState<{ loaded: boolean; masters: readonly DocumentMaster[] }>({ loaded: false, masters: [] });
  const [layersCollapsed, setLayersCollapsed] = useState(false);
  const [explorerCollapsed, setExplorerCollapsed] = useState(false);
  const [showPageBreaks, setShowPageBreaks] = useState(false);
  const pageBreakToggle = useMemo(() => ({ shown: showPageBreaks, toggle: () => setShowPageBreaks((value) => !value) }), [showPageBreaks]);
  const [activeStencilId, setActiveStencilId] = useState(shapeStencils[0].id);
  const [diagnostics, setDiagnostics] = useState<TextDiagnostic[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [contextMenu, setContextMenu] = useState<{ top: number; left: number; kind: 'shape' | 'canvas' } | null>(null);
  const contextMenuRef = useRef(contextMenu);
  contextMenuRef.current = contextMenu;
  const pointerRef = useRef<DragStart | null>(null);
  const dragShapeRef = useRef<VsdxShapeSelection | null>(null);
  const dragPreviewRef = useRef<ModelPoint | null>(null);
  const dragSnapRef = useRef(false);
  const guidesRef = useRef<SnapTargets>({ x: [], y: [] });
  const previewFrameRef = useRef<number | null>(null);
  const segmentDragRef = useRef<SegmentDrag | null>(null);
  const insertCascadeRef = useRef<Map<string, number>>(new Map());
  const marqueeRef = useRef<{ startCanvas: ModelPoint; currentCanvas: ModelPoint; pointerId?: number; startX: number; startY: number; thresholdPassed: boolean; additive: boolean } | null>(null);
  const marqueeFrameRef = useRef<number | null>(null);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const [connectorMode, setConnectorMode] = useState(false);
  const connectorModeRef = useRef(connectorMode);
  connectorModeRef.current = connectorMode;
  const hoverShapeRef = useRef<string | null>(null);
  const connectorDragRef = useRef<{ pageId: string; from: ConnectorDragEndpoint; current: ModelPoint; snap: ConnectorDragEndpoint | null } | null>(null);
  const reroutePreviewRef = useRef<ReadonlyArray<readonly ModelPoint[]>>([]);
  const connectorFrameRef = useRef<number | null>(null);
  const autoHoverRef = useRef<string | null>(null);
  const pointHoverRef = useRef<string | null>(null);
  const pointCursorRef = useRef(false);
  const autoArrowRef = useRef<{ shapeId: string; side: AutoConnectSide } | null>(null);
  const autoAlphaRef = useRef(0);
  const autoFadeStartRef = useRef(0);
  const autoFrameRef = useRef<number | null>(null);
  const autoMoveRef = useRef<{ canvas: ModelPoint; model: ModelPoint } | null>(null);
  const [quickMenu, setQuickMenu] = useState<{ shapeId: string; side: AutoConnectSide; x: number; y: number } | null>(null);
  const quickMenuRef = useRef(quickMenu);
  quickMenuRef.current = quickMenu;
  const quickMenuNodeRef = useRef<HTMLDivElement | null>(null);
  const showGridRef = useRef(showGrid);
  showGridRef.current = showGrid;
  const snapEnabledRef = useRef(snapEnabled);
  snapEnabledRef.current = snapEnabled;
  const showRulersRef = useRef(showRulers);
  showRulersRef.current = showRulers;
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
      if (pageIndex !== previous.pageIndex) setContextMenu(null);
      modelRef.current = { snapshot: current, pageIndex, frame, layers };
      setModel(modelRef.current);
      setDiagnostics(frame ? collectDiagnostics(frame) : []);
      setError(null);
      setSelection((existing) => { const kept = existing.filter((item) => stillSelectable(current, pageIndex, item, layers)); return kept.length === existing.length ? existing : kept; });
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
  const stencils = useMemo<readonly ShapeStencil[]>(() => [
    ...shapeStencils,
    { id: 'document', nameKey: 'shapesPanel.documentShapes', shapes: documentStencilEntries(documentStencil.masters) },
  ], [documentStencil]);
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
    handleRef.current?.dispose(); handleRef.current = null; imageCache.current.clear(); setSelection([]); setContextMenu(null); setClipboard(null); setEditing(null); setDraft(''); modelRef.current = { snapshot: null, pageIndex: 0, frame: null, layers: [] }; setModel(modelRef.current); setError(null); setDirty(false); setDocumentStencil({ loaded: false, masters: [] }); setActiveStencilId('standard');
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

  useLayoutEffect(() => {
    const handle = handleRef.current;
    if (!handle || loading || documentStencil.loaded || activeStencilId !== 'document') return;
    try { setDocumentStencil({ loaded: true, masters: handle.masters() }); }
    catch (value) { setDocumentStencil({ loaded: true, masters: [] }); reportError(value); }
  }, [activeStencilId, documentStencil.loaded, loading, reportError]);

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

  useEffect(() => {
    if ((!dirty && editing === null) || typeof window === 'undefined') return;
    const onBeforeUnload = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ''; };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [dirty, editing]);

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
    if (showGrid) {
      try { paintGrid(context, frame, dpr, zoom); } catch { void 0; }
    }
    const snapshot = model.snapshot;
    const page = snapshot?.pages[model.pageIndex];
    paintConnectorLayer(context, frame);
    if (page) {
      for (const item of selection) {
        if (selectionHiddenByLayers(page, model.layers, item)) continue;
        try {
          const corners = selectionCorners(page, frame, item);
          const placement = findShapePlacement(page.shapes, item.shapeId);
          const blocked = placement ? isHandleResizeBlocked(placement.shape) : false;
          const rotationBlocked = placement ? isRotateBlocked(placement.shape) : false;
          if (corners) paintSelectionFrame(context, corners, dpr, zoom, blocked ? [] : undefined, !rotationBlocked);
          if (placement && selection.length === 1) paintControlHandles(context, controlHandleCanvasPositions(placement.shape, shapeDragStart(page, placement.shape, frame), frame.paintTransform), dpr, zoom);
        } catch { void 0; }
        if (selection.length === 1) paintSelectedConnector(context, frame, page, item, segmentDragRef.current, dpr, zoom);
      }
    }
    const start = pointerRef.current; const release = dragPreviewRef.current;
    if (start && release) {
      try { paintDragPreview(context, previewOutline(start, release, frame.paintTransform, dragSnapRef.current), dpr, zoom); } catch { void 0; }
    }
    const marquee = marqueeRef.current;
    if (marquee && marquee.thresholdPassed) {
      try { paintMarquee(context, normalizeMarquee(marquee.startCanvas, marquee.currentCanvas), dpr, zoom); } catch { void 0; }
    }
    try { paintSmartGuides(context, frame, dpr, zoom, guidesRef.current); } catch { void 0; }
  }, [model.frame, model.snapshot, model.pageIndex, model.layers, selection, zoom, connectorMode, showGrid]);

  useEffect(() => {
    if (!editing) return;
    if (selection.length !== 1 || selection[0].pageId !== editing.pageId || selection[0].shapeId !== editing.shapeId) setEditing(null);
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

  useEffect(() => () => {
    if (previewFrameRef.current !== null) cancelAnimationFrame(previewFrameRef.current);
    if (connectorFrameRef.current !== null) cancelAnimationFrame(connectorFrameRef.current);
    if (autoFrameRef.current !== null) cancelAnimationFrame(autoFrameRef.current);
    if (marqueeFrameRef.current !== null) cancelAnimationFrame(marqueeFrameRef.current);
  }, []);

  const clearDragPreview = () => {
    if (previewFrameRef.current !== null) { cancelAnimationFrame(previewFrameRef.current); previewFrameRef.current = null; }
    dragPreviewRef.current = null;
    guidesRef.current = { x: [], y: [] };
    snappedReleaseRef.current = null;
    repaintOverlaySelection();
  };

  const paintMarqueeState = () => {
    const marquee = marqueeRef.current; const overlay = overlayCanvasRef.current;
    if (!marquee || !overlay || !marquee.thresholdPassed) return;
    const context = overlay.getContext('2d'); if (!context) return;
    try { paintMarquee(context, normalizeMarquee(marquee.startCanvas, marquee.currentCanvas), window.devicePixelRatio || 1, zoomRef.current); } catch { void 0; }
  };

  const repaintOverlaySelection = () => {
    const overlay = overlayCanvasRef.current; const current = modelRef.current; const frame = current.frame;
    if (!overlay || !frame) return;
    const context = overlay.getContext('2d'); if (!context) return;
    const currentSelection = selectionRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    const dpr = window.devicePixelRatio || 1;
    const zoom = zoomRef.current;
    context.clearRect(0, 0, overlay.width, overlay.height);
    if (showGridRef.current) {
      try { paintGrid(context, frame, dpr, zoom); } catch { void 0; }
    }
    paintConnectorLayer(context, frame);
    if (page) {
      for (const item of currentSelection) {
        if (selectionHiddenByLayers(page, current.layers, item)) continue;
        try {
          const corners = selectionCorners(page, frame, item);
          const placement = findShapePlacement(page.shapes, item.shapeId);
          const blocked = placement ? isHandleResizeBlocked(placement.shape) : false;
          const rotationBlocked = placement ? isRotateBlocked(placement.shape) : false;
          if (corners) paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current, blocked ? [] : undefined, !rotationBlocked);
          if (placement && currentSelection.length === 1) paintControlHandles(context, controlHandleCanvasPositions(placement.shape, shapeDragStart(page, placement.shape, frame), frame.paintTransform), window.devicePixelRatio || 1, zoomRef.current);
        } catch { void 0; }
        if (currentSelection.length === 1) paintSelectedConnector(context, frame, page, item, segmentDragRef.current, window.devicePixelRatio || 1, zoomRef.current);
      }
    }
    try { paintSmartGuides(context, frame, dpr, zoom, guidesRef.current); } catch { void 0; }
    paintMarqueeState();
  };

  /** Coalesces ruler marker updates onto one animation frame. */
  const markRulerPointer = (point: ModelPoint | null) => {
    rulerMarkRef.current = point;
    if (rulerFrameRef.current !== null) return;
    rulerFrameRef.current = requestAnimationFrame(() => { rulerFrameRef.current = null; setRulerMark(rulerMarkRef.current); });
  };

  /** Only a free move snaps: a handle resize, a rotate and a control drag keep the raw point. */
  const snappableDrag = (start: DragStart): boolean => !start.rotate && !start.control && !start.handle;

  /** The point the last preview used, so the commit cannot disagree with what was drawn. */
  const releaseForCommit = (start: DragStart, raw: ModelPoint, suppressed: boolean, previewed: { raw: ModelPoint; point: ModelPoint } | null): ModelPoint => {
    if (!snappableDrag(start)) return raw;
    if (previewed && Math.abs(previewed.raw.x - raw.x) < 1e-9 && Math.abs(previewed.raw.y - raw.y) < 1e-9) return previewed.point;
    return snapDragPoint(start, raw, suppressed).point;
  };

  /** Snaps a drag release the same way for preview and commit; Alt suppresses. */
  const snapDragPoint = (start: DragStart, raw: ModelPoint, suppressed: boolean): { point: ModelPoint; guides: SnapTargets } => {
    const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    const selected = dragShapeRef.current;
    const grouped = (start.parentTransforms?.length ?? 0) > 0;
    const targets = page && selected && !grouped ? collectSnapTargets(page.shapes, selected.shapeId) : { x: [], y: [] };
    return snapRelease(start, raw, {
      zoom: zoomRef.current,
      gridSpacing: GRID_SPACING_IN,
      snapEnabled: snapEnabledRef.current && (start.parentTransforms?.length ?? 0) === 0,
      suppressed,
      isRotate: Boolean(start.rotate),
      targets,
    });
  };

  const connectorScene = (frame: PageDisplayList): ConnectorOverlayScene => {
    const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    const drag = connectorDragRef.current;
    const connectors: ConnectorOverlayRoute[] = [];
    let hoverPoints: readonly ConnectionPoint[] = [];
    if (page) {
      const selectedIds = new Set(selectionRef.current.filter((item) => item.pageId === page.id).map((item) => item.shapeId));
      for (const shape of page.shapes) {
        if (!isConnectorShape(shape)) continue;
        const route = connectorRouteFromFrame(frame, page.sourcePartPath, shape.sourceId);
        if (!route) continue;
        const [beginGlue, endGlue] = connectorEndpointGlue(route, page.shapes);
        connectors.push({ route, selected: selectedIds.has(shape.id), beginGlue, endGlue });
      }
      const hoverId = drag?.from.shapeId ?? (connectorModeRef.current ? hoverShapeRef.current : pointHoverRef.current);
      const hovered = hoverId ? page.shapes.find((shape) => shape.id === hoverId) : undefined;
      if (hovered) hoverPoints = connectionPointsForShape(hovered);
    }
    return {
      hoverPoints,
      snapPoint: drag?.snap?.point ?? null,
      previewRoute: drag ? routeConnector(drag.from.point, drag.snap?.point ?? drag.current) : null,
      reroutePreview: reroutePreviewRef.current,
      connectors,
    };
  };

  const paintConnectorLayer = (context: CanvasRenderingContext2D, frame: PageDisplayList) => {
    const dpr = window.devicePixelRatio || 1;
    try { paintConnectorOverlay(context, frame, dpr, zoomRef.current, connectorScene(frame)); } catch { void 0; }
    if (connectorModeRef.current || connectorDragRef.current || !autoHoverRef.current) return;
    const page = modelRef.current.snapshot?.pages[modelRef.current.pageIndex];
    const placement = page ? findShapePlacement(page.shapes, autoHoverRef.current) : null;
    if (!page || !placement || placement.siblings !== page.shapes || isConnectorShape(placement.shape)) return;
    try {
      paintAutoConnectOverlay(context, frame, dpr, zoomRef.current, {
        arrows: autoConnectArrowsForShape(placement.shape),
        hovered: autoArrowRef.current?.shapeId === placement.shape.id ? autoArrowRef.current.side : null,
        alpha: autoAlphaRef.current,
      });
    } catch { void 0; }
  };

  const cancelAutoFade = () => {
    if (autoFrameRef.current !== null) cancelAnimationFrame(autoFrameRef.current);
    autoFrameRef.current = null;
  };

  const startAutoFade = () => {
    cancelAutoFade();
    autoFadeStartRef.current = performance.now();
    autoAlphaRef.current = 0;
    const tick = () => {
      autoAlphaRef.current = Math.min(1, (performance.now() - autoFadeStartRef.current) / AUTO_CONNECT_FADE_MS);
      repaintOverlaySelection();
      autoFrameRef.current = autoAlphaRef.current < 1 ? requestAnimationFrame(tick) : null;
    };
    autoFrameRef.current = requestAnimationFrame(tick);
  };

  const setPointCursor = (active: boolean) => {
    const canvas = mainCanvasRef.current;
    if (!canvas || pointCursorRef.current === active) return;
    pointCursorRef.current = active;
    canvas.style.cursor = active ? 'crosshair' : '';
  };

  /** One direct hit-test plus eight probes on a screen-pixel ring, so nearing an edge reveals its points. */
  const pointHoverShapeAt = (shapes: readonly ShapeSnapshot[], canvas: ModelPoint): string | null => {
    const handle = handleRef.current;
    if (!handle) return null;
    const probe = (x: number, y: number): string | null => {
      let hit: HitTestResult | null = null;
      try { hit = handle.hitTest(x, y); } catch { hit = null; }
      const placement = hit ? findShapePlacement(shapes, hit.shapeId) : null;
      return placement && placement.siblings === shapes && !isConnectorShape(placement.shape) ? placement.shape.id : null;
    };
    const direct = probe(canvas.x, canvas.y);
    if (direct) return direct;
    const zoom = Number.isFinite(zoomRef.current) && zoomRef.current > 0 ? zoomRef.current : 1;
    const radius = HOVER_PROXIMITY_PX / zoom;
    for (const [dx, dy] of HOVER_PROBE_DIRS) {
      const found = probe(canvas.x + dx * radius, canvas.y + dy * radius);
      if (found) return found;
    }
    return null;
  };

  const clearAutoConnect = () => {
    cancelAutoFade();
    autoHoverRef.current = null;
    autoArrowRef.current = null;
    autoAlphaRef.current = 0;
    autoMoveRef.current = null;
    pointHoverRef.current = null;
    setPointCursor(false);
    if (quickMenuRef.current) setQuickMenu(null);
  };

  const hideAutoConnect = () => { clearAutoConnect(); repaintOverlaySelection(); };

  const scheduleConnectorRepaint = () => {
    if (connectorFrameRef.current !== null) return;
    connectorFrameRef.current = requestAnimationFrame(() => { connectorFrameRef.current = null; repaintOverlaySelection(); });
  };

  /** Routes for connectors glued to the shape being dragged, recomputed at its preview geometry. */
  const gluedReroutePreview = (start: DragStart, release: ModelPoint): ReadonlyArray<readonly ModelPoint[]> => {
    const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    const active = dragShapeRef.current;
    if (!frame || !page || !active || active.pageId !== page.id) return [];
    const placement = findShapePlacement(page.shapes, active.shapeId);
    if (!placement || placement.siblings !== page.shapes) return [];
    try { return reroutePreviewForMove(page.shapes, frame, page.sourcePartPath, placement.shape, resolveDragGeometry(start, release)); }
    catch { return []; }
  };

  /** Connector mode owns the pointer: a live drag updates its snap, otherwise only a hover change repaints. */
  const onConnectorPointerMove = (event: PointerEvent<HTMLCanvasElement>) => {
    const handle = handleRef.current; const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !frame || !page) return;
    try {
      const point = canvasPointerPosition(event, frame);
      const drag = connectorDragRef.current;
      if (drag) {
        drag.current = point.model;
        drag.snap = connectorTargetForPoint(page.shapes, handle, point.canvas, point.model);
        scheduleConnectorRepaint();
        return;
      }
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const placement = hit ? findShapePlacement(page.shapes, hit.shapeId) : null;
      const hovered = placement && placement.siblings === page.shapes && !isConnectorShape(placement.shape) ? placement.shape.id : null;
      if (hovered !== hoverShapeRef.current) { hoverShapeRef.current = hovered; repaintOverlaySelection(); }
    } catch (value) { reportError(value); }
  };

  /** Coalesces AutoConnect hover to one hit-test per frame; a pointer move alone never calls wasm. */
  const scheduleAutoConnectHover = (event: PointerEvent<HTMLCanvasElement>, frame: PageDisplayList) => {
    autoMoveRef.current = canvasPointerPosition(event, frame);
    if (connectorFrameRef.current !== null) return;
    connectorFrameRef.current = requestAnimationFrame(() => {
      connectorFrameRef.current = null;
      const point = autoMoveRef.current;
      if (point) resolveAutoConnectHover(point);
    });
  };

  const resolveAutoConnectHover = (point: { canvas: ModelPoint; model: ModelPoint }) => {
    const handle = handleRef.current; const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !frame || !page || connectorModeRef.current || connectorDragRef.current || pointerRef.current) return;
    const arrowsFor = (shapeId: string | null) => {
      const placement = shapeId ? findShapePlacement(page.shapes, shapeId) : null;
      return placement && placement.siblings === page.shapes && !isConnectorShape(placement.shape) ? autoConnectArrowsForShape(placement.shape) : [];
    };
    try {
      if (autoHoverRef.current) {
        const arrow = autoConnectArrowAt(arrowsFor(autoHoverRef.current), frame, zoomRef.current, point.canvas);
        const next = arrow ? { shapeId: autoHoverRef.current, side: arrow.side } : null;
        const previous = autoArrowRef.current;
        if (next?.shapeId !== previous?.shapeId || next?.side !== previous?.side) {
          autoArrowRef.current = next;
          if (next && arrow) setQuickMenu({ shapeId: next.shapeId, side: next.side, ...autoConnectArrowCss(arrow, frame, zoomRef.current) });
          else if (quickMenuRef.current) setQuickMenu(null);
          repaintOverlaySelection();
        }
        if (arrow) { setPointCursor(false); return; }
      }
      let placement = pointHoverRef.current ? findShapePlacement(page.shapes, pointHoverRef.current) : null;
      if (placement && (placement.siblings !== page.shapes || isConnectorShape(placement.shape))) placement = null;
      if (placement && !autoConnectHaloHit(placement.shape, frame, zoomRef.current, point.canvas)) placement = null;
      if (!placement) {
        const probed = pointHoverShapeAt(page.shapes, point.canvas);
        placement = probed ? findShapePlacement(page.shapes, probed) : null;
        if (placement && (placement.siblings !== page.shapes || isConnectorShape(placement.shape))) placement = null;
      }
      const hovered = placement ? placement.shape.id : null;
      if (hovered !== autoHoverRef.current || hovered !== pointHoverRef.current) {
        autoHoverRef.current = hovered;
        pointHoverRef.current = hovered;
        autoArrowRef.current = null;
        if (quickMenuRef.current) setQuickMenu(null);
        if (hovered) startAutoFade();
        else { cancelAutoFade(); autoAlphaRef.current = 0; repaintOverlaySelection(); }
      }
      setPointCursor(placement !== null && hoverPointAt(connectionPointsForShape(placement.shape), frame, zoomRef.current, point.model) !== null);
    } catch (value) { reportError(value); }
  };

  const onPointerLeave = (event: PointerEvent<HTMLCanvasElement>) => {
    markRulerPointer(null);
    if (!connectorModeRef.current && !connectorDragRef.current) {
      const next = event.relatedTarget as Node | null;
      if (next && quickMenuNodeRef.current?.contains(next)) return;
      setPointCursor(false);
      if (autoHoverRef.current || autoArrowRef.current || quickMenuRef.current || pointHoverRef.current) hideAutoConnect();
      return;
    }
    if (!connectorModeRef.current || connectorDragRef.current) return;
    if (hoverShapeRef.current) { hoverShapeRef.current = null; repaintOverlaySelection(); }
  };

  const onPointerDown = (event: PointerEvent<HTMLCanvasElement>) => {
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    if (event.button === 2) return;
    if (pointerRef.current || marqueeRef.current) return;
    pointerRef.current = null; dragPreviewRef.current = null; dragShapeRef.current = null; segmentDragRef.current = null; marqueeRef.current = null;
    if (connectorModeRef.current) {
      try {
        const point = canvasPointerPosition(event, frame);
        const target = connectorTargetForPoint(page.shapes, handle, point.canvas, point.model);
        if (target) {
          connectorDragRef.current = { pageId: page.id, from: target, current: point.model, snap: target };
          setSelection([{ pageId: page.id, shapeId: target.shapeId, hit: { kind: 'shape', shapeId: target.shapeId } }]);
          capturePointer(event);
        } else {
          const hit = handle.hitTest(point.canvas.x, point.canvas.y);
          const connectorHit = hit ? selectionForHit(page, hit) : null;
          setSelection(connectorHit ? [connectorHit] : []);
        }
        repaintOverlaySelection();
      } catch (value) { reportError(value); }
      return;
    }
    try {
      const point = canvasPointerPosition(event, frame);
      for (const active of selectionRef.current) {
        if (active.pageId !== page.id || selectionHiddenByLayers(page, modelRef.current.layers, active)) continue;
        try {
          const placement = selectionRef.current.length === 1 ? findShapePlacement(page.shapes, active.shapeId) : null;
          if (placement) {
            const base = shapeDragStart(page, placement.shape, frame, handle);
            const controls = controlHandleCanvasPositions(placement.shape, base, frame.paintTransform);
            const row = hitTestControlHandles(point.canvas, controls, zoomRef.current);
            const hit = row ? controls.find((entry) => entry.row === row) : undefined;
            const drag = hit && row ? controlDragStart(placement.shape, row, hit) : null;
            if (hit && drag && !(drag.lockedX && drag.lockedY)) {
              pointerRef.current = {
                ...point,
                ...base,
                pointerId: event.pointerId,
                startX: event.clientX,
                startY: event.clientY,
                resize: false,
                control: drag,
              };
              dragShapeRef.current = active;
              event.currentTarget.setPointerCapture(event.pointerId);
              return;
            }
          }
        } catch { void 0; }
        try {
          const corners = selectionCorners(page, frame, active);
          if (!corners) continue;
          const target = hitTestSelection(point.canvas, corners, zoomRef.current);
          if (!target) continue;
          const handlePlacement = findShapePlacement(page.shapes, active.shapeId);
          if (!handlePlacement) continue;
          if (target === 'rotate' && isRotateBlocked(handlePlacement.shape)) {
            reportError(new Error(rotateBlockedMessage(t, handlePlacement.shape)));
            return;
          }
          if (target !== 'rotate' && isHandleResizeBlocked(handlePlacement.shape)) {
            reportError(new Error(t('errors.resizeLocked')));
            return;
          }
          const base = shapeDragStart(page, handlePlacement.shape, frame, handle);
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
          dragShapeRef.current = active;
          setSelection([active]);
          event.currentTarget.setPointerCapture(event.pointerId);
          return;
        } catch { void 0; }
      }
      if (!connectorModeRef.current) {
        const grabbedId = pointHoverRef.current ?? pointHoverShapeAt(page.shapes, point.canvas);
        const grabbedPlacement = grabbedId ? findShapePlacement(page.shapes, grabbedId) : null;
        const grabbedShape = grabbedPlacement && grabbedPlacement.siblings === page.shapes && !isConnectorShape(grabbedPlacement.shape) ? grabbedPlacement.shape : null;
        const grabbed = grabbedShape ? hoverPointAt(connectionPointsForShape(grabbedShape), frame, zoomRef.current, point.model) : null;
        if (grabbed && grabbedShape) {
          clearAutoConnect();
          pointHoverRef.current = grabbedShape.id;
          connectorDragRef.current = { pageId: page.id, from: { shapeId: grabbedShape.id, point: grabbed }, current: point.model, snap: null };
          setSelection([{ pageId: page.id, shapeId: grabbedShape.id, hit: { kind: 'shape', shapeId: grabbedShape.id } }]);
          setPointCursor(true);
          capturePointer(event);
          repaintOverlaySelection();
          return;
        }
      }
      if (autoHoverRef.current) {
        const hovered = findShapePlacement(page.shapes, autoHoverRef.current);
        const arrows = hovered && hovered.siblings === page.shapes && !isConnectorShape(hovered.shape) ? autoConnectArrowsForShape(hovered.shape) : [];
        const arrow = autoConnectArrowAt(arrows, frame, zoomRef.current, point.canvas);
        if (arrow) {
          const sourceId = autoHoverRef.current;
          connectorDragRef.current = { pageId: page.id, from: { shapeId: sourceId, point: arrow.point }, current: point.model, snap: null };
          clearAutoConnect();
          setSelection([{ pageId: page.id, shapeId: sourceId, hit: { kind: 'shape', shapeId: sourceId } }]);
          capturePointer(event);
          repaintOverlaySelection();
          return;
        }
      }
      if (autoHoverRef.current || autoArrowRef.current || quickMenuRef.current) clearAutoConnect();
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const next = hit ? selectionForHit(page, hit) : null;
      if (!next) {
        const additive = event.shiftKey;
        if (!additive) setSelection([]);
        marqueeRef.current = { startCanvas: point.canvas, currentCanvas: point.canvas, pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, thresholdPassed: false, additive };
        event.currentTarget.setPointerCapture(event.pointerId);
        return;
      }
      setSelection([next]);
      if (!selectionHiddenByLayers(page, modelRef.current.layers, next)) {
        const chrome = selectedConnectorChrome(frame, page, next);
        const segment = chrome?.draggable ? hitSegmentDot(chrome, point.model, grabTolerance(event.currentTarget, frame)) : null;
        if (chrome?.draggable && segment) {
          segmentDragRef.current = {
            selection: next,
            vertices: chrome.draggable,
            segment: segment.index,
            anchor: { ...segment.mid },
            grab: { x: point.model.x - segment.mid.x, y: point.model.y - segment.mid.y },
            pointerId: event.pointerId,
            preview: null,
          };
          event.currentTarget.setPointerCapture(event.pointerId);
          return;
        }
      }
      const placement = findShapePlacement(page.shapes, next.shapeId);
      pointerRef.current = placement ? {
        ...point,
        ...shapeDragStart(page, placement.shape, frame, handle),
        pointerId: event.pointerId,
        startX: event.clientX,
        startY: event.clientY,
        resize: event.shiftKey,
      } : null;
      dragShapeRef.current = placement ? next : null;
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch (value) { reportError(value); }
  };
  const onPointerMove = (event: PointerEvent<HTMLCanvasElement>) => {
    const marquee = marqueeRef.current;
    if (marquee) {
      if (marquee.pointerId !== undefined && marquee.pointerId !== event.pointerId) return;
      if (!marquee.thresholdPassed) {
        if (!passedDragThreshold(marquee.startX, marquee.startY, event.clientX, event.clientY)) return;
        marquee.thresholdPassed = true;
      }
      const marqueeFrame = modelRef.current.frame;
      if (!marqueeFrame) return;
      try {
        marquee.currentCanvas = canvasPointerPosition(event, marqueeFrame).canvas;
        if (marqueeFrameRef.current !== null) return;
        marqueeFrameRef.current = requestAnimationFrame(() => {
          marqueeFrameRef.current = null;
          repaintOverlaySelection();
        });
      } catch (value) { reportError(value); }
      return;
    }
    const drag = segmentDragRef.current;
    if (drag) {
      if (drag.pointerId !== event.pointerId) return;
      try {
        const liveFrame = modelRef.current.frame;
        if (!liveFrame) return;
        const point = canvasPointerPosition(event, liveFrame);
        drag.preview = dragSegmentRoute(drag.vertices, drag.segment, { x: drag.anchor.x + drag.grab.x, y: drag.anchor.y + drag.grab.y }, point.model);
        event.currentTarget.style.cursor = 'grabbing';
        repaintOverlaySelection();
      } catch (value) { reportError(value); }
      return;
    }
    const start = pointerRef.current;
    if (connectorModeRef.current) { onConnectorPointerMove(event); return; }
    if (start) {
      if (autoHoverRef.current || autoArrowRef.current || quickMenuRef.current) clearAutoConnect();
    } else {
      const hoverFrame = modelRef.current.frame;
      if (hoverFrame) { try { scheduleAutoConnectHover(event, hoverFrame); } catch { void 0; } }
    }
    if (!start) {
      if (showRulersRef.current) {
        try {
          const hoverFrame = modelRef.current.frame;
          if (hoverFrame) markRulerPointer(canvasPointerPosition(event, hoverFrame).canvas);
        } catch { void 0; }
      }
      try {
        const current = modelRef.current; const frame = current.frame;
        const page = current.snapshot?.pages[current.pageIndex];
        const selected = selectionRef.current;
        if (!frame || !page || selected.length === 0) { event.currentTarget.style.cursor = ''; return; }
        const point = canvasPointerPosition(event, frame);
        for (const active of selected) {
          if (active.pageId !== page.id || selectionHiddenByLayers(page, current.layers, active)) continue;
          const placement = selected.length === 1 ? findShapePlacement(page.shapes, active.shapeId) : null;
          if (placement) {
            try {
              const controls = controlHandleCanvasPositions(placement.shape, shapeDragStart(page, placement.shape, frame), frame.paintTransform);
              const row = hitTestControlHandles(point.canvas, controls, zoomRef.current);
              const hit = row ? controls.find((entry) => entry.row === row) : undefined;
              const drag = hit && row ? controlDragStart(placement.shape, row, hit) : null;
              if (drag && !(drag.lockedX && drag.lockedY)) { event.currentTarget.style.cursor = 'move'; return; }
            } catch { void 0; }
          }
          const corners = selectionCorners(page, frame, active);
          if (!corners) continue;
          const target = hitTestSelection(point.canvas, corners, zoomRef.current);
          const shape = findShapePlacement(page.shapes, active.shapeId)?.shape;
          if (target && shape && (target === 'rotate' ? isRotateBlocked(shape) : isHandleResizeBlocked(shape))) continue;
          if (target) { event.currentTarget.style.cursor = target === 'rotate' ? 'grab' : resizeCursor(target); return; }
          if (selected.length === 1) {
            const chrome = selectedConnectorChrome(frame, page, active);
            const hover = chrome?.draggable ? hitSegmentDot(chrome, point.model, grabTolerance(event.currentTarget, frame)) : null;
            if (hover) { event.currentTarget.style.cursor = 'grab'; return; }
          }
        }
        event.currentTarget.style.cursor = '';
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
      if (showRulersRef.current) markRulerPointer(point.canvas);
      if (snappableDrag(start)) {
        const snapped = snapDragPoint(start, point.model, event.altKey);
        dragPreviewRef.current = snapped.point;
        guidesRef.current = snapped.guides;
        snappedReleaseRef.current = { raw: point.model, point: snapped.point };
      } else {
        dragPreviewRef.current = point.model;
        guidesRef.current = { x: [], y: [] };
        snappedReleaseRef.current = null;
      }
      dragSnapRef.current = event.shiftKey;
      if (previewFrameRef.current !== null) return;
      previewFrameRef.current = requestAnimationFrame(() => {
        previewFrameRef.current = null;
        const liveFrame = modelRef.current.frame; const liveStart = pointerRef.current; const release = dragPreviewRef.current; const overlay = overlayCanvasRef.current;
        if (!liveFrame || !liveStart || !release || !overlay) return;
        const context = overlay.getContext('2d'); if (!context) return;
        try {
          if (liveStart.control) {
            const livePage = modelRef.current.snapshot?.pages[modelRef.current.pageIndex];
            const liveSelection = selectionRef.current[0];
            if (!livePage || !liveSelection) return;
            const livePlacement = findShapePlacement(livePage.shapes, liveSelection.shapeId);
            const corners = selectionCorners(livePage, liveFrame, liveSelection);
            context.clearRect(0, 0, overlay.width, overlay.height);
            if (corners) paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current);
            if (livePlacement) {
              const positions = controlHandleCanvasPositions(livePlacement.shape, liveStart, liveFrame.paintTransform).map((position) => {
                if (position.row !== liveStart.control!.row) return position;
                const next = resolveControlDrag(liveStart, liveStart.control!.startLocal, release, liveStart.control!.lockedX, liveStart.control!.lockedY);
                const page = shapeLocalToPage(liveStart, next);
                return { ...position, canvas: modelPointToCanvas(liveFrame.paintTransform, page.x, page.y) };
              });
              paintControlHandles(context, positions, window.devicePixelRatio || 1, zoomRef.current);
            }
            return;
          }
          const corners = previewOutline(liveStart, release, liveFrame.paintTransform, dragSnapRef.current);
          reroutePreviewRef.current = gluedReroutePreview(liveStart, release);
          context.clearRect(0, 0, overlay.width, overlay.height);
          if (showGridRef.current) {
            try { paintGrid(context, liveFrame, window.devicePixelRatio || 1, zoomRef.current); } catch { void 0; }
          }
          paintConnectorLayer(context, liveFrame);
          paintDragPreview(context, corners, window.devicePixelRatio || 1, zoomRef.current);
          paintSelectionFrame(context, corners, window.devicePixelRatio || 1, zoomRef.current);
          try { paintSmartGuides(context, liveFrame, window.devicePixelRatio || 1, zoomRef.current, guidesRef.current); } catch { void 0; }
        } catch (value) { reportError(value); }
      });
    } catch (value) { reportError(value); }
  };
  const commitConnectorDrag = (event: PointerEvent<HTMLCanvasElement>) => {
    const drag = connectorDragRef.current;
    connectorDragRef.current = null;
    const handle = handleRef.current; const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    try {
      if (!drag || !handle || !frame || !page) return;
      let end = drag.snap;
      try {
        const point = canvasPointerPosition(event, frame);
        end = connectorTargetForPoint(page.shapes, handle, point.canvas, point.model) ?? drag.snap;
      } catch { end = drag.snap; }
      if (!end && !connectorModeRef.current) {
        let drop: ModelPoint | null = null;
        try { drop = canvasPointerPosition(event, frame).model; } catch { drop = null; }
        if (!drop || Math.hypot(drop.x - drag.from.point.x, drop.y - drag.from.point.y) < HOVER_FREE_DRAG_INCHES) return;
        const receipt = handle.addFreeConnector(drag.pageId, connectorDraft(drag.from.point, drop), connectorGlue(drag.from.shapeId, drag.from.point));
        refresh(undefined, true);
        setSelection([{ pageId: drag.pageId, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } }]);
        pointHoverRef.current = drag.from.shapeId;
        return;
      }
      if (!end || (end.shapeId === drag.from.shapeId && end.point.side === drag.from.point.side)) return;
      const live = handle.snapshot().pages.find((item) => item.id === drag.pageId);
      if (!live || !findShapePlacement(live.shapes, drag.from.shapeId) || !findShapePlacement(live.shapes, end.shapeId)) return;
      const receipt = handle.addConnector(drag.pageId, connectorDraft(drag.from.point, end.point), connectorGlue(drag.from.shapeId, drag.from.point), connectorGlue(end.shapeId, end.point));
      refresh(undefined, true);
      setSelection([{ pageId: drag.pageId, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } }]);
      pointHoverRef.current = end.shapeId;
    } catch (value) { reportError(value); }
    finally { setPointCursor(false); repaintOverlaySelection(); }
  };

  const onPointerUp = (event: PointerEvent<HTMLCanvasElement>) => {
    if (connectorDragRef.current) { commitConnectorDrag(event); return; }
    const marquee = marqueeRef.current;
    if (marquee) {
      if (marquee.pointerId !== undefined && marquee.pointerId !== event.pointerId) return;
      marqueeRef.current = null;
      if (marqueeFrameRef.current !== null) { cancelAnimationFrame(marqueeFrameRef.current); marqueeFrameRef.current = null; }
      repaintOverlaySelection();
      if (!marquee.thresholdPassed) return;
      try {
        const current = modelRef.current;
        const marqueePage = current.snapshot?.pages[current.pageIndex];
        const marqueeFrame = current.frame;
        if (!marqueePage || !marqueeFrame) return;
        const enclosed = marqueeEnclosedShapes(marqueePage, marqueeFrame, normalizeMarquee(marquee.startCanvas, marquee.currentCanvas), current.layers);
        if (marquee.additive) {
          const known = new Set(selectionRef.current.map((item) => `${item.pageId}:${item.shapeId}`));
          setSelection([...selectionRef.current, ...enclosed.filter((item) => !known.has(`${item.pageId}:${item.shapeId}`))]);
        } else setSelection(enclosed);
      } catch (value) { reportError(value); }
      return;
    }
    const drag = segmentDragRef.current;
    if (drag && drag.pointerId === event.pointerId) {
      segmentDragRef.current = null;
      const routeHandle = handleRef.current;
      try {
        if (routeHandle && drag.preview && !sameRoutePoints(drag.preview, drag.vertices)) {
          routeHandle.setConnectorRoute(drag.selection.pageId, drag.selection.shapeId, drag.preview);
          refresh(undefined, true);
          return;
        }
      } catch (value) { reportError(value); }
      repaintOverlaySelection();
      return;
    }
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null;
    const hadPreview = dragPreviewRef.current !== null;
    reroutePreviewRef.current = [];
    const previewedRelease = snappedReleaseRef.current;
    clearDragPreview();
    const handle = handleRef.current; const selected = dragShapeRef.current; const frame = model.frame;
    dragShapeRef.current = null;
    if (!handle || !selected || !frame) return;
    try {
      const point = canvasPointerPosition(event, frame);
      if (!pointer.thresholdPassed && !hadPreview && pointer.startX !== undefined && pointer.startY !== undefined && !passedDragThreshold(pointer.startX, pointer.startY, event.clientX, event.clientY)) return;
      if (!pointer.thresholdPassed && !hadPreview && Math.abs(point.canvas.x - pointer.canvas.x) < 0.01 && Math.abs(point.canvas.y - pointer.canvas.y) < 0.01) return;
      if (pointer.control) {
        const livePage = handle.snapshot().pages.find((page) => page.id === selected.pageId);
        const liveShape = livePage ? findShapePlacement(livePage.shapes, selected.shapeId)?.shape : undefined;
        const lockedX = pointer.control.lockedX || (liveShape ? controlCellWriteBlocked(liveShape, pointer.control.row, 'X') : false);
        const lockedY = pointer.control.lockedY || (liveShape ? controlCellWriteBlocked(liveShape, pointer.control.row, 'Y') : false);
        if (lockedX && lockedY) return;
        const next = resolveControlDrag(pointer, pointer.control.startLocal, point.model, lockedX, lockedY);
        handle.setControlHandle(selected.pageId, selected.shapeId, pointer.control.row, lockedX ? null : inchFormula(next.x), lockedY ? null : inchFormula(next.y));
        refresh(undefined, true);
        return;
      }
      if (pointer.rotate) {
        const livePage = handle.snapshot().pages.find((page) => page.id === selected.pageId);
        const livePlacement = livePage ? findShapePlacement(livePage.shapes, selected.shapeId) : null;
        if (livePlacement && isRotateBlocked(livePlacement.shape)) throw new Error(rotateBlockedMessage(t, livePlacement.shape));
        handle.setCellFormula(selected.pageId, selected.shapeId, { cellName: 'Angle' }, String(resolveRotationAngle(pointer, point.model, event.shiftKey)));
        refresh(undefined, true);
        return;
      }
      const release = releaseForCommit(pointer, point.model, event.altKey, previewedRelease);
      const geometry = resolveDragGeometry(pointer, release);
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
  const clearMarquee = (event: PointerEvent<HTMLCanvasElement>): boolean => {
    const marquee = marqueeRef.current;
    if (!marquee) return false;
    if (marquee.pointerId !== undefined && marquee.pointerId !== event.pointerId) return false;
    marqueeRef.current = null;
    if (marqueeFrameRef.current !== null) { cancelAnimationFrame(marqueeFrameRef.current); marqueeFrameRef.current = null; }
    repaintOverlaySelection();
    return true;
  };
  const abandonConnectorDrag = () => {
    if (!connectorDragRef.current) return false;
    connectorDragRef.current = null;
    repaintOverlaySelection();
    return true;
  };
  const cancelSegmentDrag = (event: PointerEvent<HTMLCanvasElement>): boolean => {
    const drag = segmentDragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return false;
    segmentDragRef.current = null;
    repaintOverlaySelection();
    return true;
  };
  const onPointerCancel = (event: PointerEvent<HTMLCanvasElement>) => {
    if (abandonConnectorDrag()) return;
    if (cancelSegmentDrag(event)) return;
    if (clearMarquee(event)) return;
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null; reroutePreviewRef.current = []; clearDragPreview();
  };
  const onLostPointerCapture = (event: PointerEvent<HTMLCanvasElement>) => {
    if (abandonConnectorDrag()) return;
    if (cancelSegmentDrag(event)) return;
    if (clearMarquee(event)) return;
    const pointer = pointerRef.current;
    if (!pointer) return;
    if (pointer.pointerId !== undefined && pointer.pointerId !== event.pointerId) return;
    pointerRef.current = null; reroutePreviewRef.current = []; clearDragPreview();
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
      setSelection([next]);
      enterTextEdit(next.pageId, next.shapeId);
    } catch (value) { reportError(value); }
  };
  const onCanvasContextMenu = (event: MouseEvent<HTMLCanvasElement>) => {
    event.preventDefault();
    const handle = handleRef.current; const frame = model.frame; const page = model.snapshot?.pages[model.pageIndex];
    if (!handle || !frame || !page) return;
    if (pointerRef.current || marqueeRef.current) return;
    try {
      const point = canvasPointerPosition(event, frame);
      handle.layoutPage(model.pageIndex);
      const hit = handle.hitTest(point.canvas.x, point.canvas.y);
      const next = hit ? selectionForHit(page, hit) : null;
      if (!next) {
        setSelection([]);
        setContextMenu({ top: event.clientY, left: event.clientX, kind: 'canvas' });
        return;
      }
      const active = selectionRef.current;
      if (!active.some((item) => item.pageId === next.pageId && item.shapeId === next.shapeId)) setSelection([next]);
      setContextMenu({ top: event.clientY, left: event.clientX, kind: 'shape' });
    } catch (value) { reportError(value); }
  };
  const commandsRef = useRef<RibbonCommands | null>(null);
  const cancelActiveDrag = () => {
    const pointer = pointerRef.current;
    const marquee = marqueeRef.current;
    if (!pointer && !marquee) return false;
    pointerRef.current = null;
    marqueeRef.current = null;
    dragShapeRef.current = null;
    clearDragPreview();
    if (marqueeFrameRef.current !== null) { cancelAnimationFrame(marqueeFrameRef.current); marqueeFrameRef.current = null; }
    repaintOverlaySelection();
    const canvas = mainCanvasRef.current;
    const captured = pointer?.pointerId ?? marquee?.pointerId;
    if (canvas && captured !== undefined) {
      try {
        if (typeof canvas.hasPointerCapture !== 'function' || canvas.hasPointerCapture(captured)) canvas.releasePointerCapture(captured);
      } catch { void 0; }
    }
    return true;
  };
  const nudgeSelection = (dx: number, dy: number) => {
    const handle = handleRef.current; const selected = selectionRef.current;
    if (!handle || selected.length === 0) return;
    try {
      const snapshot = handle.snapshot();
      const frame = modelRef.current.frame;
      if (!frame) return;
      const moves: ShapeMove[] = [];
      for (const item of selected) {
        const page = snapshot.pages.find((entry) => entry.id === item.pageId);
        if (!page) continue;
        const placement = findShapePlacement(page.shapes, item.shapeId);
        if (!placement) continue;
        const base = shapeDragStart(page, placement.shape, frame);
        const geometry = resolveNudgeGeometry({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, ...base }, dx, dy);
        moves.push({ pageId: item.pageId, shapeId: item.shapeId, xFormula: inchFormula(geometry.x), yFormula: inchFormula(geometry.y) });
      }
      if (moves.length > 0) handle.moveShapes(moves);
      refresh(undefined, true);
    } catch (value) {
      try { refresh(undefined, false); } catch { void 0; }
      reportError(value);
    }
  };
  const selectAllShapes = () => {
    const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!page || page.shapes.length === 0) return;
    setSelection(page.shapes
      .map((shape) => ({ pageId: page.id, shapeId: shape.id, hit: { kind: 'shape', shapeId: shape.id } as const }))
      .filter((item) => !selectionHiddenByLayers(page, current.layers, item)));
  };
  const toggleConnector = useCallback(() => {
    connectorDragRef.current = null;
    hoverShapeRef.current = null;
    clearAutoConnect();
    setConnectorMode((value) => !value);
  }, []);
  const runKeyboardIntent = (intent: CanvasKeyboardIntent) => {
    const commands = commandsRef.current;
    if (intent.kind === 'undo') { if (commands?.undo.enabled) commands.undo.run(); return; }
    if (intent.kind === 'redo') { if (commands?.redo.enabled) commands.redo.run(); return; }
    if (intent.kind === 'delete') { if (commands?.delete.enabled) commands.delete.run(); return; }
    if (intent.kind === 'save') { if (commands?.download.enabled) commands.download.run(); return; }
    if (intent.kind === 'escape') { cancelActiveDrag(); hideAutoConnect(); closeContextMenu(); setSelection([]); return; }
    if (intent.kind === 'selectAll') { selectAllShapes(); return; }
    nudgeSelection(intent.dx, intent.dy);
  };
  const copySelected = useCallback(() => {
    const handle = handleRef.current; const single = selection.length === 1 ? selection[0] : null;
    if (!handle || !single) return;
    try { setClipboard(copySelection(handle, single)); } catch (value) { reportError(value); }
  }, [selection, reportError]);
  const cutSelected = useCallback(() => {
    const handle = handleRef.current; const single = selection.length === 1 ? selection[0] : null;
    if (!handle || !single) return;
    try {
      setClipboard(copySelection(handle, single));
      handle.deleteShape(single.pageId, single.shapeId);
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  }, [selection, refresh, reportError]);
  const pasteClipboardEntry = useCallback(() => {
    const handle = handleRef.current;
    const entry = clipboardRef.current;
    if (!handle || !entry) return;
    try {
      const target = modelRef.current.snapshot?.pages[modelRef.current.pageIndex]?.id ?? selection[0]?.pageId ?? entry.pageId;
      const step = entry.pasteCount + 1;
      const { receipt, entry: next } = pasteEntry(handle, target, entry, PASTE_OFFSET.x * step, PASTE_OFFSET.y * step);
      setClipboard(next);
      setSelection([{ pageId: target, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } }]);
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  }, [selection, refresh, reportError]);
  const duplicateSelected = useCallback(() => {
    const handle = handleRef.current; const single = selection.length === 1 ? selection[0] : null;
    if (!handle || !single) return;
    try {
      const entry = copySelection(handle, single);
      const receipt = duplicateEntry(handle, single.pageId, entry, DUPLICATE_OFFSET.x, DUPLICATE_OFFSET.y);
      setSelection([{ pageId: single.pageId, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } }]);
      refresh(undefined, true);
    } catch (value) { reportError(value); }
  }, [selection, refresh, reportError]);
  const onCanvasKeyDown = (event: KeyboardEvent<HTMLCanvasElement>) => {
    if (event.altKey && event.key === '3') { event.preventDefault(); toggleConnector(); return; }
    const selected = selectionRef.current.length === 1 ? selectionRef.current[0] : null;
    if (!editingRef.current && (event.ctrlKey || event.metaKey) && !event.altKey) {
      const key = event.key.toLowerCase();
      if (key === 'c' && selected) { event.preventDefault(); copySelected(); return; }
      if (key === 'x' && selected) { event.preventDefault(); cutSelected(); return; }
      if (key === 'v' && clipboardRef.current) { event.preventDefault(); pasteClipboardEntry(); return; }
      if (key === 'd' && selected) { event.preventDefault(); duplicateSelected(); return; }
    }
    if (selected && !editingRef.current) {
      if (event.key === 'Enter') { event.preventDefault(); enterTextEdit(selected.pageId, selected.shapeId); return; }
      if (isPrintableEntryKey(event)) { event.preventDefault(); enterTextEdit(selected.pageId, selected.shapeId, event.key); return; }
    }
    const intent = canvasKeyboardIntent(event, zoomRef.current);
    if (!intent) return;
    if (editingRef.current) return;
    event.preventDefault();
    runKeyboardIntent(intent);
  };
  /** An open menu owns its own keys; the editor chrome owns the rest of the canvas shortcuts. */
  const chromeOwnsIntent = () => !contextMenuRef.current
    && !editorRootRef.current?.querySelector('[role="menu"], [role="dialog"], [role="listbox"], dialog[open], [aria-modal="true"]');
  const onEditorKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.defaultPrevented) return;
    const intent = canvasKeyboardIntent(event, zoomRef.current);
    if (intent) {
      if (intent.kind !== 'save' && intent.kind !== 'undo' && intent.kind !== 'redo' && !chromeOwnsIntent()) return;
      event.preventDefault();
      if (editingRef.current) return;
      runKeyboardIntent(intent);
      return;
    }
    if (isOwnedBrowserShortcut(event)) event.preventDefault();
  };
  const onCanvasFocus = (event: FocusEvent<HTMLCanvasElement>) => { event.currentTarget.style.outline = '2px solid #0f6cbd'; event.currentTarget.style.outlineOffset = '2px'; };
  const onCanvasBlur = (event: FocusEvent<HTMLCanvasElement>) => { event.currentTarget.style.outline = ''; event.currentTarget.style.outlineOffset = ''; };
  const toggleView = useCallback((key: ViewToggleKey) => {
    if (key === 'grid') setShowGrid((value) => !value);
    else if (key === 'snap') setSnapEnabled((value) => !value);
    else setShowRulers((value) => !value);
  }, []);
  const insertShapeAt = useCallback((shape: StandardShape, point: ModelPoint) => {
    const handle = handleRef.current; const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !page) return;
    try {
      const receipt = handle.addShape(page.id, shape.draft(point.x, point.y, shape.defaultSize.width, shape.defaultSize.height));
      refresh(undefined, true);
      setSelection([{ pageId: page.id, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } }]);
    } catch (value) { reportError(value); }
  }, [refresh, reportError]);
  const insertQuickShape = useCallback((shape: StandardShape, sourceId: string, side: AutoConnectSide) => {
    const handle = handleRef.current; const current = modelRef.current;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!handle || !page) return;
    try {
      const live = handle.snapshot().pages.find((item) => item.id === page.id);
      const placement = live ? findShapePlacement(live.shapes, sourceId) : null;
      if (!live || !placement || placement.siblings !== live.shapes || isConnectorShape(placement.shape)) return;
      const source = placement.shape;
      const layout = quickShapePlacement(source, side, Math.max(0.25, numericCellValue(source, 'Width', 1)), Math.max(0.25, numericCellValue(source, 'Height', 1)));
      if (!layout) return;
      const receipt = handle.addConnectedShape(page.id, shape.draft(layout.x, layout.y, layout.width, layout.height), connectorDraft(layout.from, layout.to), connectorGlue(source.id, layout.from), layout.to.toCell);
      refresh(undefined, true);
      setSelection([{ pageId: page.id, shapeId: receipt.shape.shapeId, hit: { kind: 'shape', shapeId: receipt.shape.shapeId } }]);
    } catch (value) { reportError(value); }
    hideAutoConnect();
  }, [refresh, reportError]);
  const quickMenuShapes = useMemo(() => QUICK_SHAPE_IDS.map((id) => stencilShapeById(id)).filter((shape): shape is StandardShape => Boolean(shape)), []);
  const insertShape = useCallback((shape: StandardShape) => {
    const current = modelRef.current; const frame = current.frame;
    const page = current.snapshot?.pages[current.pageIndex];
    if (!frame || !page) return;
    const cascade = insertCascadeRef.current.get(page.id) ?? 0;
    insertCascadeRef.current.set(page.id, cascade + 1);
    insertShapeAt(shape, centreInsertPoint(canvasPointToModel(frame.paintTransform, frame.width / 2, frame.height / 2), cascade));
  }, [insertShapeAt]);
  const droppedShape = useCallback((id: string) => stencils.flatMap((stencil) => stencil.shapes).find((shape) => shape.id === id), [stencils]);
  const onCanvasDragOver = (event: DragEvent<HTMLDivElement>) => {
    if (!handleRef.current || !modelRef.current.frame) return;
    if (!event.dataTransfer.types.includes(STENCIL_DRAG_MIME)) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  };
  const onCanvasDrop = (event: DragEvent<HTMLDivElement>) => {
    const frame = modelRef.current.frame; const canvas = mainCanvasRef.current;
    if (!frame || !canvas) return;
    const shape = droppedShape(event.dataTransfer.getData(STENCIL_DRAG_MIME).trim());
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
  const commitShapeData = useCallback((row: ShapeDataRow, formula: string) => {
    const handle = handleRef.current; const selected = selectionRef.current.length === 1 ? selectionRef.current[0] : null;
    if (!handle || !selected) return;
    const locator: CellLocator = {
      cellName: 'Value',
      section: 'Property',
      ...(row.sectionIndex !== undefined ? { sectionIndex: row.sectionIndex } : {}),
      ...(row.rowName !== null ? { rowName: row.rowName } : { rowIndex: row.rowIndex ?? 0 }),
    };
    handle.setCellFormula(selected.pageId, selected.shapeId, locator, formula);
    refresh(undefined, true);
  }, [refresh]);
  const selectFromExplorer = useCallback((next: VsdxShapeSelection | null) => {
    const { snapshot, pageIndex, layers } = modelRef.current;
    if (next && !(snapshot && stillSelectable(snapshot, pageIndex, next, layers))) return;
    setSelection(next ? [next] : []);
  }, []);
  const selectedShape = (() => {
    const page = model.snapshot?.pages[model.pageIndex];
    if (!page || selection.length !== 1 || selection[0].pageId !== page.id) return null;
    return findShapePlacement(page.shapes, selection[0].shapeId)?.shape ?? null;
  })();
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
  return <div ref={editorRootRef} className={className} style={styles.root} aria-label={t('editor.appLabel')} onKeyDown={onEditorKeyDown}>
    <header style={styles.titleBar}><strong>{t('ribbon.documentName')}</strong><span style={{ color: dirty ? '#a16207' : '#526273' }}>{dirty ? t('ribbon.dirty') : t('ribbon.saved')}</span></header>
    <RibbonCommandsProvider handle={handleRef.current} snapshot={model.snapshot} pageId={model.snapshot?.pages[model.pageIndex]?.id} selection={selection} frame={model.frame} clipboard={clipboard} onClipboardChange={setClipboard} onSelectShape={(next) => setSelection([next])} pageBreaks={pageBreakToggle} onMutation={() => refresh(undefined, true)} onError={reportError} onDownload={download}>
    <RibbonCommandsBridge target={commandsRef} />
    <Ribbon t={t} hasSelection={selection.length > 0} connector={{ active: connectorMode, disabled: !model.frame, onToggle: toggleConnector }} view={{ grid: showGrid, snap: snapEnabled, rulers: showRulers }} onToggleView={toggleView} />
    <div style={styles.contentRow}>
    {leftPanel === undefined ? (
      <div style={styles.leftColumn}>
        <div style={styles.shapesWrap}>
          <ShapesPanel stencils={stencils} activeStencilId={activeStencilId} onSelectStencil={setActiveStencilId} collapsed={shapesCollapsed} onToggleCollapsed={() => setShapesCollapsed((value) => !value)} onInsert={insertShape} t={t} />
        </div>
        <LayersPanel layers={model.layers} collapsed={layersCollapsed} onToggleCollapsed={() => setLayersCollapsed((value) => !value)} onToggleLayer={toggleLayerVisible} t={t} />
      </div>
    ) : leftPanel}
    <main ref={workspaceRef} style={styles.workspace}>
      {loading && <span>{t('editor.opening')}</span>}
      {!loading && !model.frame && <span>{file ? t('editor.noPages') : t('editor.openPrompt')}</span>}
      <div style={showRulers ? styles.sheet : styles.sheetPlain}>
        {showRulers && <div style={styles.corner} />}
        {showRulers && <div style={styles.rulerTop}>{model.frame && <RulerTop frame={model.frame} zoom={zoom} marker={rulerMark} />}</div>}
        {showRulers && <div style={styles.rulerLeft}>{model.frame && <RulerLeft frame={model.frame} zoom={zoom} marker={rulerMark} />}</div>}
        <div style={styles.canvasFrame} onDragOver={onCanvasDragOver} onDrop={onCanvasDrop}>
          <canvas ref={mainCanvasRef} tabIndex={0} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerLeave={onPointerLeave} onPointerUp={onPointerUp} onPointerCancel={onPointerCancel} onLostPointerCapture={onLostPointerCapture} onDoubleClick={onCanvasDoubleClick} onContextMenu={onCanvasContextMenu} onKeyDown={onCanvasKeyDown} onFocus={onCanvasFocus} onBlur={onCanvasBlur} aria-label={canvasLabel(t, model.pageIndex, model.snapshot?.pages.length ?? 0, selection)} style={connectorMode ? { ...styles.canvas, cursor: 'crosshair' } : styles.canvas} />
          {showPageBreaks && model.frame && <PageBreakGrid frame={model.frame} zoom={zoom} />}
          <canvas ref={overlayCanvasRef} aria-hidden="true" style={styles.overlay} />
          {quickMenu && model.frame && (
            <div ref={quickMenuNodeRef} role="menu" aria-label={t('shapesPanel.quickShapes')} style={{ ...styles.quickMenu, left: Math.max(4, Math.min(quickMenu.x + 16, model.frame.width * zoom - 44)), top: Math.max(100, Math.min(quickMenu.y, model.frame.height * zoom - 100)) }} onMouseLeave={() => { autoArrowRef.current = null; setQuickMenu(null); repaintOverlaySelection(); }}>
              {quickMenuShapes.map((shape) => (
                <button key={shape.id} type="button" role="menuitem" aria-label={t(shape.nameKey)} title={t(shape.nameKey)} onClick={() => insertQuickShape(shape, quickMenu.shapeId, quickMenu.side)} style={styles.quickShape}>
                  <svg aria-hidden="true" viewBox="0 0 1 1" preserveAspectRatio="xMidYMid meet" style={styles.quickPreview}><path d={shape.preview} /></svg>
                </button>
              ))}
            </div>
          )}
          {editing && editOverlay && (
            <div ref={editWrapRef} style={{ ...styles.textEditWrap, width: editOverlay.width, height: editOverlay.height, transform: `matrix(${editOverlay.matrix.a}, ${editOverlay.matrix.b}, ${editOverlay.matrix.c}, ${editOverlay.matrix.d}, ${editOverlay.matrix.e}, ${editOverlay.matrix.f})` }}>
              <textarea
                ref={editBoxRef}
                value={draft}
                onChange={(event) => setDraft(event.target.value)}
                onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); commitTextEdit(); mainCanvasRef.current?.focus(); return; } if (isOwnedBrowserShortcut(event)) event.preventDefault(); event.stopPropagation(); }}
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
      </div>
      {contextMenu?.kind === 'shape' && selection.length > 0 && <ShapeContextMenu t={t} position={contextMenu} onClose={closeContextMenu} onCloseAndFocus={closeContextMenuAndFocus} />}
      {contextMenu?.kind === 'canvas' && <CanvasContextMenu t={t} position={contextMenu} onClose={closeContextMenu} onCloseAndFocus={closeContextMenuAndFocus} />}
      {integrity.length > 0 && <section role="alert" style={styles.integrity}><strong>{t('diagnostics.integrityHeading')}</strong>{integrity.map((item, index) => <div key={`${item.code}-${index}`}>{diagnosticMessage(t, item.category, item.code)}</div>)}</section>}
      {fidelity.length > 0 && <details style={styles.fidelity}><summary>{t('diagnostics.fidelityHeading')}</summary>{fidelity.map((item, index) => <div key={`${item.code}-${index}`}>{diagnosticMessage(t, item.category, item.code)}</div>)}</details>}
      {error && <div role="alert" style={styles.error}><span style={styles.errorText}>{error}</span><button type="button" aria-label={t('errors.dismiss')} onClick={() => setError(null)} style={styles.errorDismiss}>×</button></div>}
    </main>
    {rightPanel === undefined ? (
      <div style={styles.rightColumn}>
        <DrawingExplorer snapshot={model.snapshot} activePageIndex={model.pageIndex} selection={selection[0] ?? null} onSelectPage={(index) => refresh(index)} onSelectShape={selectFromExplorer} collapsed={explorerCollapsed} onToggleCollapsed={() => setExplorerCollapsed((value) => !value)} t={t} />
        <ShapeDataPanel shape={selectedShape} onCommit={commitShapeData} onError={reportError} t={t} />
      </div>
    ) : rightPanel}
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

/** Page extent and the printer-paper tile behind the page-break grid, in page pixels. */
export interface PageBreakFrame { width: number; height: number; printWidth: number; printHeight: number; }

/** Upper bound on page-break guides per axis; a denser tile draws none. */
export const MAX_PAGE_BREAK_LINES = 1000;

/** Page-break offsets in CSS pixels along one axis, excluding the page edges. */
function pageBreakAxis(extent: number, tile: number, zoom: number): number[] {
  const step = tile * zoom;
  const span = extent * zoom;
  if (!Number.isFinite(step) || step <= 0 || !Number.isFinite(span)) return [];
  if (span / step > MAX_PAGE_BREAK_LINES + 1) return [];
  const offsets: number[] = [];
  for (let k = 1; k * step < span - 1e-6; k += 1) offsets.push(k * step);
  return offsets;
}

/** Where printer-paper boundaries fall inside the page, in CSS pixels from the page origin. */
export function pageBreakLines(frame: PageBreakFrame, zoom: number): { vertical: number[]; horizontal: number[] } {
  return {
    vertical: pageBreakAxis(frame.width, frame.printWidth, zoom),
    horizontal: pageBreakAxis(frame.height, frame.printHeight, zoom),
  };
}

/** Printer-paper guides drawn over the page. */
export function PageBreakGrid({ frame, zoom }: { frame: PageBreakFrame; zoom: number }) {
  const lines = pageBreakLines(frame, zoom);
  return <div data-testid="vsdx-page-breaks" aria-hidden="true" style={styles.overlay}>
    {lines.vertical.map((x) => <div key={`v${x}`} style={{ ...styles.pageBreakLine, left: x, top: 0, width: 1, height: '100%' }} />)}
    {lines.horizontal.map((y) => <div key={`h${y}`} style={{ ...styles.pageBreakLine, left: 0, top: y, width: '100%', height: 1 }} />)}
  </div>;
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

const HOVER_PROBE_DIRS: ReadonlyArray<readonly [number, number]> = [[1, 0], [-1, 0], [0, 1], [0, -1], [1, 1], [1, -1], [-1, 1], [-1, -1]];

export function canvasPointerPosition(event: PointerEvent<HTMLCanvasElement> | MouseEvent<HTMLCanvasElement>, frame: PageDisplayList): { canvas: ModelPoint; model: ModelPoint } {
  return clientPointToModel(frame, event.currentTarget.getBoundingClientRect(), event.clientX, event.clientY);
}

export function inchFormula(value: number): string {
  if (!Number.isFinite(value)) throw new Error('Shape geometry must be finite.');
  const rounded = Number(value.toFixed(6));
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

export interface SegmentDrag { selection: VsdxShapeSelection; vertices: ChromePoint[]; segment: number; anchor: ChromePoint; grab: ChromePoint; pointerId: number; preview: ChromePoint[] | null; }

/** Pointer-grab radius for a segment dot, in model units. */
export function grabTolerance(canvas: HTMLCanvasElement, frame: PageDisplayList): number {
  const rect = canvas.getBoundingClientRect();
  const scale = Math.hypot(frame.paintTransform.a, frame.paintTransform.b);
  if (!Number.isFinite(scale) || scale <= 0) return 0;
  return 12 * frame.width / (scale * Math.max(rect.width, 1));
}

export function sameRoutePoints(left: readonly ChromePoint[], right: readonly ChromePoint[]): boolean {
  return left.length === right.length && left.every((point, index) => Math.hypot(point.x - right[index].x, point.y - right[index].y) < 1e-9);
}

function paintSelectedConnector(context: CanvasRenderingContext2D, frame: PageDisplayList, page: PageSnapshot, selection: VsdxShapeSelection, drag: SegmentDrag | null, dpr: number, zoom: number): void {
  const chrome = selectedConnectorChrome(frame, page, selection);
  if (!chrome) return;
  const active = drag && drag.selection.shapeId === selection.shapeId ? previewChrome(chrome, drag.preview ?? drag.vertices) : chrome;
  paintConnectorChrome(context, frame, active, dpr, zoom);
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
/** Nearest connection point of a known shape, in model inches. */
export function nearestPointOnShape(shapes: readonly ShapeSnapshot[], shapeId: string, at: ModelPoint): ConnectionPoint | null {
  const placement = findShapePlacement(shapes, shapeId);
  if (!placement || placement.siblings !== shapes || isConnectorShape(placement.shape)) return null;
  return nearestConnectionPointAnywhere(connectionPointsForShape(placement.shape), at);
}

/** Drop target forgiving of interior drops; falls back to the hit-tested shape. */
export function connectorTargetForPoint(shapes: readonly ShapeSnapshot[], handle: DiagramHandle, canvas: ModelPoint, at: ModelPoint): ConnectorDragEndpoint | null {
  const direct = dropTargetForPoint(shapes, at);
  if (direct) return direct;
  let shapeId: string | null = null;
  try { shapeId = handle.hitTest(canvas.x, canvas.y)?.shapeId ?? null; } catch { shapeId = null; }
  if (!shapeId) return null;
  const point = nearestPointOnShape(shapes, shapeId, at);
  return point ? { shapeId, point } : null;
}

/** Pointer capture is best-effort; synthetic pointers must not fail the gesture. */
function capturePointer(event: PointerEvent<HTMLCanvasElement>): void {
  try { event.currentTarget.setPointerCapture(event.pointerId); } catch { void 0; }
}

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

const rotateBlockedMessage = (t: TFunction, shape: ShapeSnapshot): string =>
  isCellWriteBlocked(shape, 'Angle') ? t('errors.rotationGuarded') : t('errors.rotateLocked');

export function canvasLabel(t: TFunction, pageIndex: number, total: number, selection: readonly VsdxShapeSelection[]): string {
  if (selection.length > 1) return t('pages.canvasLabelWithMultiSelection', { current: pageIndex + 1, total, count: selection.length });
  if (selection.length === 1) return t('pages.canvasLabelWithSelection', { current: pageIndex + 1, total, name: selection[0].shapeId });
  return t('pages.canvasLabel', { current: pageIndex + 1, total });
}

/** Top-level shapes fully enclosed by the marquee, matching click selection's rule and layer visibility. */
export function marqueeEnclosedShapes(page: PageSnapshot, frame: PageDisplayList, rect: MarqueeRect, layers: readonly PageLayer[] = []): VsdxShapeSelection[] {
  const enclosed: VsdxShapeSelection[] = [];
  for (const shape of page.shapes) {
    const item: VsdxShapeSelection = { pageId: page.id, shapeId: shape.id, hit: { kind: 'shape', shapeId: shape.id } };
    if (selectionHiddenByLayers(page, layers, item)) continue;
    try {
      const corners = selectionCorners(page, frame, item);
      if (corners && marqueeEnclosesQuad(corners, rect)) enclosed.push(item);
    } catch { void 0; }
  }
  return enclosed;
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

export function controlDragStart(shape: ShapeSnapshot, row: string, hit: { lockedX: boolean; lockedY: boolean }): ControlDrag | null {
  const handle = controlHandlesForShape(shape).find((entry) => entry.row === row);
  if (!handle) return null;
  return { row, startLocal: { x: handle.x, y: handle.y }, lockedX: hit.lockedX || controlCellWriteBlocked(shape, row, 'X'), lockedY: hit.lockedY || controlCellWriteBlocked(shape, row, 'Y') };
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
const styles: Record<string, CSSProperties> = { root: { display: 'flex', flexDirection: 'column', width: '100%', height: '100%', minHeight: 480, color: '#172033', background: '#f3f5f8', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }, titleBar: { display: 'flex', alignItems: 'center', gap: 12, minHeight: 32, padding: '0 14px', background: '#f8fafc', borderBottom: '1px solid #d8dee9', fontSize: 13 }, contentRow: { display: 'flex', flex: 1, minHeight: 0 }, leftColumn: { display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 }, rightColumn: { display: 'flex', height: '100%', minHeight: 0 }, shapesWrap: { display: 'flex', flex: '1 1 auto', minHeight: 0 }, workspace: { position: 'relative', display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', overflow: 'auto' }, canvasFrame: { position: 'relative', flex: '0 0 auto' }, canvas: { display: 'block', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' }, overlay: { position: 'absolute', inset: 0, pointerEvents: 'none' }, pageBreakLine: { position: 'absolute', pointerEvents: 'none', background: '#c3ccd9' }, quickMenu: { position: 'absolute', zIndex: 3, display: 'flex', flexDirection: 'column', gap: 4, padding: 4, background: '#fff', border: '1px solid #d8dee9', borderRadius: 6, boxShadow: '0 8px 24px rgba(27, 39, 61, 0.18)', transform: 'translateY(-50%)' }, quickShape: { appearance: 'none', display: 'grid', placeItems: 'center', width: 32, height: 32, padding: 3, border: '1px solid transparent', borderRadius: 4, background: 'transparent', cursor: 'pointer' }, quickPreview: { width: 24, height: 24, overflow: 'visible', fill: '#fff', stroke: '#172033', strokeWidth: 0.05 }, textEditWrap: { position: 'absolute', left: 0, top: 0, transformOrigin: '0 0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'visible', background: 'transparent', border: '1px dashed #1d4ed8', zIndex: 2 }, textEditBox: { width: '100%', background: 'transparent', border: 'none', outline: 'none', resize: 'none', overflow: 'visible', textAlign: 'center', whiteSpace: 'pre-wrap', overflowWrap: 'break-word', wordBreak: 'break-word', lineHeight: 1.2, padding: 0, margin: 0 }, integrity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 12, color: '#7f1d1d', background: '#fef2f2', border: '1px solid #fca5a5' }, fidelity: { position: 'absolute', right: 14, bottom: 14, maxWidth: 340, padding: 8, color: '#475569', background: '#fff', fontSize: 12 }, error: { position: 'absolute', left: 14, right: 14, bottom: 14, display: 'flex', alignItems: 'center', gap: 8, padding: 10, color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0' },
  errorText: { flex: '1 1 auto' },
  errorDismiss: { flex: '0 0 auto', appearance: 'none', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: 'inherit', fontSize: 16, lineHeight: '28px', cursor: 'pointer' },
  sheet: { display: 'grid', gridTemplateColumns: `${RULER_SIZE}px auto`, gridTemplateRows: `${RULER_SIZE}px auto`, width: 'max-content', margin: 'auto' },
  sheetPlain: { display: 'grid', width: 'max-content', margin: 'auto' },
  corner: { position: 'sticky', left: 0, top: 0, zIndex: 6, width: RULER_SIZE, height: RULER_SIZE, background: '#f3f5f8' },
  rulerTop: { position: 'sticky', top: 0, zIndex: 5, alignSelf: 'start', background: '#f3f5f8' },
  rulerLeft: { position: 'sticky', left: 0, zIndex: 5, justifySelf: 'start', background: '#f3f5f8' } };
