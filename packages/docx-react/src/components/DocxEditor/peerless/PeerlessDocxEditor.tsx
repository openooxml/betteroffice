import { ResidentSourceWorkerClient as ResidentEngineWorkerClient, takePreloadedResidentSourceWorker as takePreloadedResidentEngineWorker } from '@betteroffice/docx/peerless';
import { useWorkerHeaderFooterEditing } from '../hooks/useHeaderFooterEditing';
import { InlineHeaderFooterEditor } from '../../InlineHeaderFooterEditor';
import { WorkerContentControls } from './WorkerContentControls';
import { useWorkerRulerControls } from './workerPropertyControls';
import { useWorkerPropertyDialogs } from './useWorkerPropertyDialogs';
import { useWorkerQuery } from './useWorkerQuery';
import { useCommentSidebarItems } from '../../../hooks/useCommentSidebarItems';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  applyFrameDelta,
  decodeFrameDelta,
  findDisplayListHyperlinkAtPoint,
  yrsIdToNumericId,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { createRustMeasureSource, type ResidentFontRequirement } from '@betteroffice/docx/layout';
import { extractEmbeddedFontFaces, sanitizeHref, toArrayBuffer, type DocxInput } from '@betteroffice/docx/utils';
import { HyperlinkPopup, type HyperlinkPopupData } from '../../ui/HyperlinkPopup';
import { toast } from 'sonner';
import {
  residentCaretSnapshotForFrame,
  type DocxCommandFrame,
  type DocxCommandPageIdentity,
  type ResidentCommandConnection,
  type YrsResidentFontRegistration,
} from '@betteroffice/docx/yrs';
import { createResidentCommandConnection, discardCommandPresentation, SessionFailure, type ResidentCommandCheckpoint } from '@betteroffice/docx/peerless';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { createPeerlessShellController } from './shellController';
import { registerTableActionExecutor } from './browserHarness';
import { commandPageGeometry } from './commandPageGeometry';
import { createDocxPluginHost, type DocxPluginHost } from '../../../plugins/createDocxPluginHost';
import { usePluginReservations } from '../../../plugins/reservations';
import { DocxCommandProvider } from '../../../commands/DocxCommandProvider';
import { DocxEditorShell } from '../DocxEditorShell';
import { DocxEditorToolbar } from '../DocxEditorToolbar';
import { EditorChromeContext } from '../../EditorToolbarContext';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from '../../DocxEditorHelpers';
import { PluginOverlays } from '../../../plugins/PluginOverlays';
import { managedSidebarItems, mergeSidebarItems } from '../../../plugins/PluginSidebarItems';
import { UnifiedSidebar } from '../../UnifiedSidebar';
import { createWorkerPluginGeometry, workerDocumentVersion, workerReadRefusal } from './workerPluginGeometry';
import { workerProposalSidebarItems } from './workerProposalSidebar';
import { workerRevisionCallbacks, workerRevisionRequest } from './workerRevisions';
import { PluginDock } from '../../../plugins/PluginPanels';
import { useIsDark } from '../hooks/useIsDark';
import { copyWorkerSelection, createClipboardPrefetch, writeClipboardCapture, type ClipboardCapture, isPasteFragmentAllowed } from './clipboard';
import { admitInputIntents } from './inputAdmission';
import type { DocxClipboardFragment, DocxReservationsPayload } from '@betteroffice/docx/yrs';
import { useWorkerFileIO } from './useWorkerFileIO';
import { useKeyboardShortcuts } from '../hooks/useKeyboardShortcuts';
import { yrsSelectionState } from '../hooks/useSelectionTracker';
import type { PagedEditorRef } from '../PagedEditor';
import { openPrintWindow } from '@betteroffice/docx';
import { createWorkerPrintCapture, printBitmap } from './printPages';
import { decodeCommandFrameDeltas } from './frameDeltas';
import { useWorkerContextMenu } from './useWorkerContextMenu';
import { useWorkerComments, WorkerCommentButton } from './useWorkerComments';
import { useControllableBoolean } from '../hooks/useControllableBoolean';
import { useWorkerNavigation } from './useWorkerNavigation';
import { createWorkerSearch } from './workerSearch';
import { FindReplaceDialog } from '../../dialogs/FindReplaceDialog';
import { WorkerDocumentReader } from './WorkerDocumentReader';
import { DEFAULT_PAGE_WIDTH, pluginOverlaysStyles } from '../internals/styles';
import { OUTLINE_BUTTON_RESERVED_SPACE, OUTLINE_RESERVED_SPACE } from '../../DocumentOutline';
import { REVIEW_CARD_WIDTH, SIDEBAR_PAGE_GAP, SIDEBAR_WIDTH } from '../../sidebar/constants';
import { TrackedChangesReview } from '../../sidebar/TrackedChangesReview';
import { reviewRevisionIds, type ReviewSidebarItem } from '../../sidebar/trackedChanges';
import type { PeerlessDocumentSource, PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import { peerlessPreviewPaint, releaseWorkerPreview, WorkerDocumentPreview, workerPreviewPages, type WorkerPreviewJob } from './WorkerDocumentPreview';
import { fontEnvironmentForRegistrations, prepareWorkerFontEnvironment } from './workerFonts';
import type { Document } from '@betteroffice/docx/types/document';
import type { ResidentFontEnvironment } from '@betteroffice/docx/yrs';
import { PeerlessPagedEditor } from './PeerlessPagedEditor';
import { useWorkerCollaboration } from './useWorkerCollaboration';
import type { PeerlessFrameGeometry, PeerlessViewIdentity } from './viewStore';

const PREVIEW_CUT = { blocks: 200, paragraphBudget: 256 };
const SIDEBAR_EDGE = 16;
type ResidentResume = Parameters<NonNullable<UsePeerlessEditSessionOptions['recovery']>['reopen']>[0];

/** @internal */
export const PeerlessDocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function PeerlessDocxEditor(props, ref) {
  const errorHandler = useRef(props.onError);
  errorHandler.current = props.onError;
  const reportRecoverableError = useCallback((error: Error) => { errorHandler.current?.(error, { recoverable: true }); }, []);
  const recoverableProps = { ...props, onError: reportRecoverableError };
  const {
  initialZoom = 1,
  readOnly = false,
  className,
  style,
  placeholder,
  loadingIndicator,
  memoryBudget,
  measurementFontProvider,
  } = props;
  type ParsedDocument = NonNullable<DocxEditorProps['document']>;
  const [replacement, setReplacement] = useState<{ input: DocxInput | null; parsed: ParsedDocument | null; prop: DocxEditorProps['documentBuffer']; document: DocxEditorProps['document'] } | null>(null);
  const replaced = replacement && replacement.prop === props.documentBuffer && replacement.document === props.document;
  const documentBuffer = replaced ? replacement.input : props.documentBuffer;
  const parsedDocument = documentBuffer ? null : replaced ? replacement.parsed : props.document;
  const hasSource = !!(documentBuffer || parsedDocument);
  const workerCollaboration = useWorkerCollaboration(props.collaboration, documentBuffer ?? parsedDocument);
  const sourceKey = useMemo(() => crypto.randomUUID(), [documentBuffer, parsedDocument]);
  const openedSource = useRef(sourceKey);
  const loadRequest = useRef<{ input: DocxInput | ParsedDocument; resolve(): void; reject(error: Error): void } | null>(null);
  const load = useCallback((input: DocxInput | null, parsed: ParsedDocument | null) => {
    loadRequest.current?.reject(new Error('Document was replaced'));
    return new Promise<void>((resolve, reject) => {
      loadRequest.current = { input: (input ?? parsed)!, resolve, reject };
      setReplacement({ input, parsed, prop: props.documentBuffer, document: props.document });
    });
  }, [props.documentBuffer, props.document]);
  const loadBuffer = useCallback((input: DocxInput) => load(input, null), [load]);
  const loadDocument = useCallback((document: ParsedDocument) => load(null, document), [load]);
  const host = useRef<HTMLDivElement>(null);
  const container = useRef<HTMLDivElement>(null);
  const scroller = useRef<HTMLDivElement>(null);
  const [hostWidth, setHostWidth] = useState(0);
  const hostObserver = useRef<ResizeObserver | null>(null);
  const attachHost = useCallback((element: HTMLDivElement | null) => {
    host.current = element;
    hostObserver.current?.disconnect();
    hostObserver.current = null;
    if (!element) return;
    setHostWidth(element.clientWidth);
    hostObserver.current = new ResizeObserver(() => setHostWidth(element.clientWidth));
    hostObserver.current.observe(element);
  }, []);
  const content = useRef<HTMLDivElement>(null);
  const isDark = useIsDark(props.colorMode ?? 'light');
  const worker = useRef<ResidentEngineWorkerClient | null>(null);
  const releaseResident = useRef<(() => void) | null>(null);
  const openContext = useRef<{ signal: AbortSignal; getRasterScale?: () => number } | null>(null);
  const sourceComments = useRef<NonNullable<DocxEditorProps['comments']>>([]);
  const [zoom, setZoom] = useState(initialZoom);
  const [tableGridlines, setTableGridlines] = useState(props.showTableGridlines === true);
  const [gridlinesProp, setGridlinesProp] = useState(props.showTableGridlines);
  if (gridlinesProp !== props.showTableGridlines) {
    setGridlinesProp(props.showTableGridlines);
    setTableGridlines(props.showTableGridlines === true);
  }
  const [hyperlinkPopup, setHyperlinkPopup] = useState<HyperlinkPopupData | null>(null);
  const linkGesture = useRef(0);
  const focus = useCallback(() => host.current?.querySelector<HTMLTextAreaElement>('.paged-editor__peerless-input')?.focus({ preventScroll: true }), []);
  const read = useCallback<ResidentEngineWorkerClient['documentRead']>((request) => {
    if (!worker.current) return Promise.reject(new Error('Peer-less worker is not ready'));
    return worker.current.documentRead(request);
  }, []);
  const readAt = useCallback<ResidentEngineWorkerClient['documentReadAt']>((request, token) => {
    if (!worker.current) return Promise.reject(new Error('Peer-less worker is not ready'));
    return worker.current.documentReadAt(request, token);
  }, []);
  const save = useCallback(async () => {
    if (!worker.current) throw new Error('Peer-less worker is not ready');
    return (await worker.current.save(props.comments ? { comments: [...props.comments], commentsSource: 'host' } : { comments: [...sourceComments.current] })).bytes;
  }, [props.comments]);
  const printCaptures = useRef(new Set<ReturnType<typeof createWorkerPrintCapture>>());
  const print = useCallback(async (preview = false) => {
    const popup = openPrintWindow('Print', '');
    let capture: ReturnType<typeof createWorkerPrintCapture> | null = null;
    try {
      if (!popup) throw new Error('Print window was blocked');
      const current = active.current;
      const client = worker.current;
      if (!current || !client) throw new Error('Peer-less worker is not ready');
      capture = createWorkerPrintCapture(current.binding, printBitmap, () => active.current === current && !popup.closed);
      printCaptures.current.add(capture);
      const pages = await capture.collect(
        (indices, epoch) => client.buildPages(indices, epoch), (epoch) => client.completeLayout(epoch));
      const css = popup.document.createElement('style');
      css.textContent = 'body { margin: 0; } img { display: block; width: 100%; break-after: page; } img:last-child { break-after: auto; } @page { margin: 0; size: auto; } @media print { button { display: none; } }';
      popup.document.head.appendChild(css);
      const images = pages.map((src) => { const image = popup.document.createElement('img'); image.src = src;
        popup.document.body.appendChild(image); return image; });
      await Promise.all(images.map((image) => image.decode?.().catch(() => {}) ?? Promise.resolve()));
      capture.assertCurrent();
      const perform = () => { popup.focus(); popup.print(); props.onPrint?.(); popup.close(); };
      if (preview) {
        const button = popup.document.createElement('button'); button.textContent = 'Print'; button.onclick = perform;
        popup.document.body.prepend(button);
      } else perform();
    } catch (error) {
      popup?.close();
      reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
      throw error;
    } finally { if (capture) printCaptures.current.delete(capture); }
  }, [props.onPrint, reportRecoverableError]);
  const prepareFontFamily = useRef<(family: string) => Promise<ResidentFontEnvironment>>(async () => { throw new Error('Worker fonts are unavailable'); });
  const [commentsOpen, setCommentsOpen] = useControllableBoolean(props.commentsSidebarOpen, props.onCommentsSidebarOpenChange);
  const [expandedSidebarItem, setExpandedSidebarItem] = useState<string | null>(null);
  const controllerOptions = { ...recoverableProps, zoom, focus, read, readAt, save, loadBuffer, loadDocument, print, onZoom: setZoom, isDark,
    comments: () => props.comments ?? sourceComments.current, hostComments: props.comments,
    commentsSidebar: { open: commentsOpen, control: props.commentsSidebarOpen === undefined ? 'internal' as const : props.onCommentsSidebarOpenChange ? 'host' as const : 'fixed' as const,
      toggle: () => { setCommentsOpen(open => !open); setExpandedSidebarItem(null); } },
    tableGridlines, onTableGridlines: setTableGridlines,
    layoutRequest: () => {
      const opened = fontSource.current;
      const current = active.current;
      if (!opened || !current) return null;
      const layout = JSON.parse(opened.layoutInput);
      return JSON.stringify({ ...layout, renderEnv: { ...layout.renderEnv, ...JSON.parse(current.requestedRenderEnvironment || '{}') } });
    },
    prepareFontFamily: (family: string) => prepareFontFamily.current(family) };
  const [controller] = useState(() => createPeerlessShellController(controllerOptions));
  const [search] = useState(() => createWorkerSearch(controller, (rect) => {
    const element = scroller.current;
    if (element) element.scrollTop += (host.current?.getBoundingClientRect().top ?? 0) - element.getBoundingClientRect().top + rect.y * controller.ref.getZoom() - element.clientHeight / 2;
  }));
  const searchState = useSyncExternalStore(search.subscribe, search.getSnapshot, search.getSnapshot);
  const fileIO = useWorkerFileIO(recoverableProps, controller.ref, loadBuffer, () => active.current?.binding ?? null);
  const getBinding = useCallback(() => active.current?.binding ?? null, []);
  const navigation = useWorkerNavigation(controller, getBinding, scroller, host, zoom, recoverableProps);
  useLayoutEffect(() => controller.update({ ...controllerOptions, imagePropertiesAvailable: propertyDialogs.imagePropertiesAvailable, image: propertyDialogs.imageState, onImageWrap: propertyDialogs.applyImageWrap, onImageTransform: propertyDialogs.applyImageTransform, search: search.host, pickImage: fileIO.pickImage, openSplitCell: propertyDialogs.openSplitCell, openTableProperties: propertyDialogs.openTableProperties, perform: async (id) => {
    if (id === 'pageSetup') return propertyDialogs.openPageSetup();
    if (id === 'watermark') return propertyDialogs.openWatermark();
    if (id === 'insertLink') return propertyDialogs.openLink();
    if (id === 'imageProperties') return propertyDialogs.openImage();
    if (id === 'find' || id === 'replace') { search.open(id === 'replace'); return { ok: true, status: 'opened' }; }
    if (id === 'print') { await print(); return { ok: true, status: 'executed' }; }
    return fileIO.perform(id);
  }, readPositionAtPoint: async (x, y) => { await pluginGeometry.update(); return pluginGeometry.geometry()?.readPositionAtPoint(x, y) ?? null; },
    currentPage: navigation.readCurrentPage, scrollToPage: navigation.scrollToPage, scrollToParagraph: navigation.scrollToParagraph,
    scrollToReviewTarget: navigation.scrollToReviewTarget }));
  const shellState = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  useEffect(() => registerTableActionExecutor((action) =>
    controller.commands.store.execute('tableAction', action as never)), [controller]);
  const documentReady = shellState.status === 'ready' && openedSource.current === sourceKey;
  useEffect(() => {
    let current = true;
    if (shellState.status === 'ready' && (props.plugins?.length || props.showHostProposalsInSidebar) && shellState.view?.geometry && shellState.view.editState?.review) {
      void controller.review().catch(error => { if (current) reportRecoverableError(error instanceof Error ? error : new Error(String(error))); });
    }
    return () => { current = false; };
  }, [controller, shellState.status, shellState.view, props.plugins?.length, props.showHostProposalsInSidebar, reportRecoverableError]);
  const effectiveReadOnly = readOnly || (props.mode ?? shellState.mode) === 'viewing';
  const parts = useWorkerHeaderFooterEditing(controller, reportRecoverableError);
  const propertyDialogs = useWorkerPropertyDialogs(controller, recoverableProps);
  const rulers = useWorkerRulerControls(controller, reportRecoverableError, props.showRuler === true);
  const legacyRef = useRef<PagedEditorRef | null>(null);
  useKeyboardShortcuts({ commands: controller.commands, pagedEditorRef: legacyRef, containerRef: container,
    disableFindReplaceShortcuts: props.disableFindReplaceShortcuts ?? false });
  const [overlayLayer, setOverlayLayer] = useState<HTMLDivElement | null>(null);
  const overlayRef = useRef<HTMLDivElement | null>(null);
  const [geometryRevision, redrawGeometry] = useState(0);
  const sentReservationsRef = useRef<() => DocxReservationsPayload | null>(() => null);
  const [pluginGeometry] = useState(() => createWorkerPluginGeometry(controller, () => host.current, () => overlayRef.current, () => redrawGeometry(value => value + 1),
    (target, focus, expectVersion) => navigation.revealParagraph(target.paraId, target.story, focus, expectVersion), () => sentReservationsRef.current()));
  const setPluginLayer = useCallback((element: HTMLDivElement | null) => { overlayRef.current = element; setOverlayLayer(element); }, []);
  const [pluginHost] = useState<DocxPluginHost>(() => createDocxPluginHost({
    pagedEditorRef: { current: null }, writeMode: () => controller.environment().readOnly ? 'viewing' : controller.environment().mode,
    commands: () => controller.commands, layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: controller.subscribe, geometry: pluginGeometry.geometry, subscribeGeometry: pluginGeometry.subscribe, translate: (key) => controller.environment().translate(key),
    workerRead: {
      version: async () => { const result = await controller.ref.readParagraphs({ view: 'accepted', paraIds: [] });
        return 'version' in result ? { ok: true, version: result.version } : result; },
      readParagraphs: controller.ref.readParagraphs, findText: controller.ref.findText, validateEdits: controller.ref.validateEdits,
      listStories: () => controller.read({ kind: 'listStories' }),
      readStories: (request) => controller.read({ kind: 'readStories', request }),
    },
    workerEdits: { applyEdits: (request, authorize, commit) => controller.applyEdits(request, authorize,
      (view) => commit(() => pluginHost.workerChanged(workerDocumentVersion(view)))) },
    workerNavigation: pluginGeometry.navigation,
    workerRefusal: (error) => workerReadRefusal(controller, error),
  }));
  const activations = useSyncExternalStore(pluginHost.subscribe, pluginHost.activations, pluginHost.activations);
  const reservations = usePluginReservations(pluginHost, activations);
  const reservationsRef = useRef<{ payload: DocxReservationsPayload | null; sent: boolean; source: DocxReservationsPayload | string | null;
    keyed: DocxReservationsPayload | null; serial: number }>({ payload: reservations, sent: false, source: null, keyed: null, serial: 0 });
  reservationsRef.current.payload = reservations;
  /** The set the editor sends, keyed so a frame names the set it was laid out with; once sent, none is an empty set. */
  const sentReservations = useCallback((): DocxReservationsPayload | null => {
    const reserved = reservationsRef.current;
    const source = reserved.payload ?? (reserved.sent ? pluginHost.version() : null);
    if (source === null) return null;
    if (reserved.source !== source) {
      reserved.source = source;
      reserved.keyed = { ...(typeof source === 'string' ? { version: source, sets: [] } : source), key: String(++reserved.serial) };
    }
    return reserved.keyed;
  }, [pluginHost]);
  sentReservationsRef.current = sentReservations;
  useEffect(() => {
    pluginHost.setReporter(props.onPluginError);
    pluginHost.setGrants(props.pluginGrants);
    pluginHost.setPlugins(props.plugins);
    return pluginHost.subscribe(() => pluginHost.syncCommands());
  }, [pluginHost, props.plugins, props.pluginGrants, props.onPluginError]);
  useEffect(() => {
    if (documentReady) {
      pluginHost.modeChanged(shellState.mode, effectiveReadOnly);
      if (pluginHost.generation() === null) pluginHost.openWorker(workerDocumentVersion(shellState.view));
      else pluginHost.workerChanged(workerDocumentVersion(shellState.view));
      const review = controller.reviewSnapshot();
      if (review) pluginHost.workerProposalsChanged(review.proposals);
      const layout = pluginGeometry.layout();
      pluginHost.layoutChanged(layout);
      const toolbar = shellState.view?.editState?.toolbar;
      if (layout && pluginGeometry.geometry()) props.onRenderedDomContextReady?.(pluginGeometry.geometry()!.dom);
      pluginHost.selectionChanged({ formatting: toolbar ? yrsSelectionState({ context: toolbar, startParagraphIndex: -1, endParagraphIndex: -1 }) : null, displayRange: pluginGeometry.displaySelection() });
      if (layout) pluginHost.geometryPresented(layout);
      pluginHost.syncCommands();
    } else { pluginHost.close('document-replaced'); setExpandedSidebarItem(null); }
  }, [shellState, documentReady, pluginHost, props.mode, effectiveReadOnly, pluginGeometry, overlayLayer, geometryRevision]);
  const pluginSidebarItemCount = props.pluginSidebarItems?.length ?? 0;
  useLayoutEffect(() => { if (activations.length || pluginSidebarItemCount || props.onRenderedDomContextReady || props.showHostProposalsInSidebar) void pluginGeometry.update().catch(() => {}); },
    [shellState.view, zoom, overlayLayer, pluginGeometry, activations.length, pluginSidebarItemCount, props.onRenderedDomContextReady, props.showHostProposalsInSidebar]);
  useLayoutEffect(() => {
    const scrollerElement = scroller.current;
    const refresh = () => pluginHost.geometryChanged();
    scrollerElement?.addEventListener('scroll', refresh, { passive: true });
    window.addEventListener('resize', refresh);
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(refresh);
    for (const element of [scrollerElement, host.current, overlayRef.current]) if (element) observer?.observe(element);
    return () => { observer?.disconnect(); scrollerElement?.removeEventListener('scroll', refresh); window.removeEventListener('resize', refresh); };
  }, [pluginHost, overlayLayer]);
  useEffect(() => () => { loadRequest.current?.reject(new Error('Editor was closed')); loadRequest.current = null;
    pluginHost.close('unmounted'); pluginGeometry.dispose(); search.dispose(); controller.dispose(); }, [controller, pluginHost, search]);
  const identity = useRef<PeerlessViewIdentity | null>(null);
  const retained = useRef<RetainedFrame | null>(null);
  const pageIndexes = useRef(new Map<string, number>());
  const prepared = useRef(new WeakMap<DocxCommandFrame, PeerlessFrameGeometry>());
  const [displayWindow, setDisplayWindow] = useState<readonly [number, number]>([0, 3]);
  useLayoutEffect(() => { pluginGeometry.setDisplayWindow(displayWindow); }, [pluginGeometry, displayWindow]);
  const renderEnvironment = useRef({ theme: null as DocxEditorProps['theme'], requested: '' });
  const fontSource = useRef<{ document: Document; layoutInput: string; families: string[]; initial: ResidentFontEnvironment;
    signature: string | null; provider: DocxEditorProps['measurementFontProvider'] } | null>(null);
  const fontAbort = useRef<AbortController | null>(null);
  const fontJob = useRef<{ signature: string; provider: DocxEditorProps['measurementFontProvider'] } | null>(null);
  const settings = useRef({ readOnly: effectiveReadOnly, workerLimitBytes: memoryBudget?.workerLimitBytes, measurementFontProvider,
    theme: props.theme, showHiddenText: props.showHiddenText ?? false, showTableGridlines: tableGridlines, previewFirstPage: props.previewFirstPage, mediaTokens: props.mediaTokens, fonts: props.fonts, reportFonts: !!props.onFontsLoaded });
  useLayoutEffect(() => {
    settings.current = { readOnly: effectiveReadOnly, workerLimitBytes: memoryBudget?.workerLimitBytes, measurementFontProvider,
      theme: props.theme, showHiddenText: props.showHiddenText ?? false, showTableGridlines: tableGridlines, previewFirstPage: props.previewFirstPage, mediaTokens: props.mediaTokens, fonts: props.fonts, reportFonts: !!props.onFontsLoaded };
  }, [effectiveReadOnly, memoryBudget?.workerLimitBytes, measurementFontProvider, props.theme, props.showHiddenText, tableGridlines, props.previewFirstPage, props.mediaTokens, props.fonts, props.onFontsLoaded]);
  type Ticket = ReturnType<PeerlessEditSessionBinding['session']['submit']>;
  const active = useRef<{
    binding: PeerlessEditSessionBinding;
    requestedReadOnly: boolean;
    requestedRenderEnvironment: string;
    barrier?: Ticket;
    clipboard?: ReturnType<typeof createClipboardPrefetch>;
  } | null>(null);
  const sourceGeneration = useMemo(() => ({}), [documentBuffer, parsedDocument, replaced ? replacement : null]);
  const [loaded, setLoaded] = useState<{ generation: object; source: Uint8Array } | null>(null);
  const reportError = useCallback((error: Error) => { workerCollaboration.retire(); controller.fail(error); errorHandler.current?.(error); }, [controller, workerCollaboration.retire]);
  const [previewJob, setPreviewJob] = useState<WorkerPreviewJob | null>(null);
  const onPreviewPresented = useCallback((frame: WorkerPreviewJob['frame']) => {
    controller.previewPresented(frame);
    if (applyFrameDelta(null, decodeFrameDelta(frame.frame)).pages.some(({ page }) => !page.unbuilt)) controller.painted();
  }, [controller]);
  const source = useMemo(() => {
    if (parsedDocument) return { ...parsedDocument };
    if (documentBuffer instanceof Uint8Array) return new Uint8Array(documentBuffer);
    if (documentBuffer instanceof ArrayBuffer) return new Uint8Array(documentBuffer.slice(0));
    return loaded?.generation === sourceGeneration ? new Uint8Array(loaded.source) : null;
  }, [documentBuffer, parsedDocument, sourceGeneration, loaded, workerCollaboration.reopenKey]);

  useEffect(() => {
    if (!documentBuffer || documentBuffer instanceof Uint8Array || documentBuffer instanceof ArrayBuffer) return;
    let current = true;
    void toArrayBuffer(documentBuffer).then((buffer) => {
      if (current) setLoaded({ generation: sourceGeneration, source: new Uint8Array(buffer) });
    }, (error: unknown) => {
      if (current) reportError(error instanceof Error ? error : new Error(String(error)));
    });
    return () => { current = false; };
  }, [documentBuffer, sourceGeneration, reportError]);

  const openResident = useCallback(async (input: PeerlessDocumentSource, signal: AbortSignal, getRasterScale?: () => number,
    resume?: ResidentResume): Promise<ResidentCommandConnection> => {
    signal.throwIfAborted();
    const checkpoint = resume?.checkpoint as ResidentCommandCheckpoint | undefined;
    if (checkpoint && !checkpoint.recoveryState) {
      throw new SessionFailure('message', 'Cannot restore DOCX checkpoint command-local state and proposals');
    }
    releaseResident.current?.();
    const client = (resume ? null : takePreloadedResidentEngineWorker()) ?? new ResidentEngineWorkerClient();
    const collaboration = workerCollaboration.begin(client);
    worker.current = client;
    const stopFonts = client.onFontEnvironment(event => {
      if (worker.current === client && !signal.aborted) controller.fontEnvironment(event);
    });
    const stopMemory = client.onMemory(samples => {
      if (worker.current === client && !signal.aborted) controller.memory(samples);
    });
    const fonts: YrsResidentFontRegistration[] = [];
    const fontIds: number[] = [];
    let nextFontId = 0;
    let measurement: ReturnType<typeof createRustMeasureSource> | undefined;
    let initialFrame: DocxCommandFrame | undefined;
    let disposed = false;
    const disposeFonts = () => {
      if (disposed) return;
      disposed = true;
      fonts.length = 0;
      fontIds.length = 0;
      measurement?.setEmbeddedFaces([]);
      measurement?.clear();
    };
    const appendFont = (font: YrsResidentFontRegistration) => {
      signal.throwIfAborted();
      if (disposed) throw new Error('DOCX command fonts are disposed');
      return fonts.push(font instanceof Uint8Array ? font.slice() : font) - 1;
    };
    let released = false;
    const abort = () => {
      if (released) return;
      released = true;
      signal.removeEventListener('abort', abort);
      if (releaseResident.current === abort) releaseResident.current = null;
      stopFonts(); stopMemory(); disposeFonts(); client.destroy();
      if (worker.current === client) { worker.current = null; fontAbort.current?.abort(); fontJob.current = null; fontSource.current = null; }
    };
    signal.addEventListener('abort', abort, { once: true });
    try {
      await client.prewarm();
      signal.throwIfAborted();
      const measure = async (metadata: Document) => {
        const theme = settings.current.theme ?? metadata.package.theme;
        const themeColors = Object.fromEntries(Object.entries(theme?.colorScheme ?? {}).filter((entry): entry is [string, string] => typeof entry[1] === 'string'));
        const environment = { themeColors, showHiddenText: settings.current.showHiddenText, showTableGridlines: settings.current.showTableGridlines };
        const layout = buildResidentRegionLayoutRequest(metadata, 24, { ...environment,
          defaultTabStopTwips: metadata.package.settings?.defaultTabStop ?? null, mediaTokens: settings.current.mediaTokens === true });
        const requirements = JSON.parse(await client.fontRequirements(JSON.stringify(layout))) as ResidentFontRequirement[];
        signal.throwIfAborted();
        measurement ??= createRustMeasureSource({
          engine: {
            registerFont: (data) => {
              const slot = appendFont(data);
              fontIds[slot] = nextFontId++;
              return slot;
            },
            registerSubstituteFont: (substituteOf, family) => {
              const slot = appendFont({ substituteOf, family });
              fontIds[slot] = fontIds[substituteOf];
              return slot;
            },
            clearFonts: disposeFonts,
          },
          bundled: settings.current.measurementFontProvider,
        });
        measurement.setCompat(metadata.package.settings?.compatibilityFlags);
        const embeddedFaces = await extractEmbeddedFontFaces(metadata);
        signal.throwIfAborted();
        measurement.setEmbeddedFaces(embeddedFaces);
        await measurement.prepareFontRequirements(requirements);
        signal.throwIfAborted();
        const config = measurement.measurementConfigForRequirements(requirements);
        if (!config) throw new Error('DOCX command fonts are not ready');
        layout.measurement = {
          ...config,
          fontChains: Object.fromEntries(Object.entries(config.fontChains).map(([key, ids]) =>
            [key, [...new Set(ids.map((id) => fontIds[id]))]])),
        };
        return { environment, layout, requirements, config };
      };
      let measuring: Promise<unknown> = Promise.resolve();
      const measured = (metadata: Document) => {
        const turn = measuring.then(() => measure(metadata));
        measuring = turn.catch(() => {});
        return turn;
      };
      let painted = false;
      let dropPreview = () => {};
      const showPreview = async ({ layout, config }: Awaited<ReturnType<typeof measure>>, live: () => boolean) => {
        if (!live()) return false;
        const preview = await client.previewSource({ fonts: fonts.slice(), fontRegistrationSlots: { fontChains: config.fontChains },
          layoutInput: JSON.stringify({ ...layout, cachedPageTotals: true }), extras: JSON.stringify({ fontChains: config.fontChains }) });
        signal.throwIfAborted();
        const pages = workerPreviewPages(preview);
        if (!live() || pages.presented.length === 0) return false;
        controller.preview(preview);
        return await new Promise<boolean>((resolve, reject) => {
          const settle = (operation: () => void) => { signal.removeEventListener('abort', cancel); operation(); };
          const cancel = () => settle(() => reject(signal.reason));
          signal.addEventListener('abort', cancel, { once: true });
          const job: WorkerPreviewJob = { client, frame: preview, pages, live: () => live() && !signal.aborted,
            resolve: () => settle(() => { painted = true; resolve(true); }), reject: (error) => settle(() => reject(error)) };
          dropPreview = () => settle(() => {
            setPreviewJob(current => current === job ? null : current);
            releaseWorkerPreview(client);
            resolve(false);
          });
          setPreviewJob(job);
        });
      };
      const withinPreviewBudget = async (attempt: (live: () => boolean) => Promise<boolean>): Promise<boolean | null> => {
        let live = true;
        let timer: ReturnType<typeof setTimeout> | undefined;
        try {
          return await Promise.race([attempt(() => live),
            new Promise<null>((resolve) => { timer = setTimeout(() => resolve(null), peerlessPreviewPaint.timeoutMs); })]);
        } finally {
          live = false;
          clearTimeout(timer);
          if (!painted) dropPreview();
        }
      };
      const previewFailed = (error: unknown) => {
        signal.throwIfAborted();
        reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
      };
      let previewed = false;
      if (!resume && settings.current.previewFirstPage === true && input instanceof Uint8Array && settings.current.mediaTokens !== true &&
        !workerCollaboration.openOptions?.initialUpdate) {
        try {
          previewed = await withinPreviewBudget(async (live) => {
            const cut = await client.openSourceCut({ bytes: input, cut: PREVIEW_CUT, heapLimitBytes: settings.current.workerLimitBytes });
            signal.throwIfAborted();
            return cut !== null && live() && await showPreview(await measured(cut), live);
          }) !== false;
        } catch (error) { previewed = true; previewFailed(error); }
      }
      let opened: Awaited<ReturnType<typeof client.openSource>>;
      try {
        opened = await client.openSource({ ...(input instanceof Uint8Array
          ? { bytes: input, ...(resume?.checkpoint ? { state: resume.checkpoint.state } : {}) } : { document: input }),
          heapLimitBytes: settings.current.workerLimitBytes, mediaTokens: settings.current.mediaTokens === true,
          collaboration });
      } catch (error) {
        if (!painted) throw error;
        signal.throwIfAborted();
        const failure = error instanceof Error ? error : new Error(String(error));
        controller.failOpening(failure);
        errorHandler.current?.(failure);
        return await new Promise<never>((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
      }
      signal.throwIfAborted();
      if (collaboration) workerCollaboration.opened(client, opened);
      const metadata = opened.document;
      sourceComments.current = metadata.package.document.comments ?? [];
      const opening = await measured(metadata);
      const { environment, layout, requirements, config } = opening;
      renderEnvironment.current = { theme: metadata.package.theme, requested: JSON.stringify(environment) };
      if (!resume && settings.current.previewFirstPage === true && !previewed) {
        try { await withinPreviewBudget((live) => showPreview(opening, live)); } catch (error) { previewFailed(error); }
      }
      const initial = await client.commandInit({
        session: resume?.session ?? crypto.randomUUID(), incarnation: resume?.incarnation ?? 1, epoch: resume?.epoch ?? 1,
        ...(resume ? { version: resume.checkpoint?.version ?? 0, throughSeq: resume.checkpoint?.throughSeq ?? 0,
          refusedSeqs: resume.refusedSeqs, replayAppliedSeqs: resume.replayAppliedSeqs,
          recoveryState: checkpoint?.recoveryState } : {}),
        readOnly: settings.current.readOnly,
        fonts: fonts.slice(),
        rasterScale: getRasterScale?.() ?? 1,
        fontRegistrationSlots: { fontChains: config.fontChains },
        layoutInput: JSON.stringify(layout),
        layoutWithRegions: true, layoutRevision: 1, extras: JSON.stringify({ fontChains: config.fontChains }),
        displayWindow: [0, 3], provisionalPages: 3,
      });
      initialFrame = initial.frame;
      const connection = await createResidentCommandConnection(client, initial);
      signal.throwIfAborted();
      const families = [...new Set([...requirements.map(item => item.family),
        ...(metadata.package.fontTable?.fonts ?? []).map(font => font.name)])];
      fontSource.current = { document: metadata, layoutInput: JSON.stringify(layout), families,
        initial: fontEnvironmentForRegistrations(fonts, config.fontChains, requirements, families),
        signature: settings.current.fonts === undefined && !settings.current.reportFonts ? '[]' : null,
        provider: settings.current.measurementFontProvider };
      controller.fontChoices(families);
      identity.current = {
        session: connection.initial.session,
        incarnation: connection.initial.incarnation,
        epoch: connection.initial.epoch,
      };
      retained.current = null;
      pageIndexes.current = new Map();
      prepared.current = new WeakMap();
      releaseResident.current = abort;
      return connection;
    } catch (error) {
      discardCommandPresentation(initialFrame);
      abort();
      throw error;
    } finally {
      disposeFonts();
    }
  }, [controller, workerCollaboration.openOptions, workerCollaboration.begin, workerCollaboration.opened]);
  const openConnection = useCallback((input: PeerlessDocumentSource, signal: AbortSignal, getRasterScale?: () => number) => {
    signal.throwIfAborted();
    active.current = null;
    openContext.current = { signal, getRasterScale };
    return openResident(input, signal, getRasterScale);
  }, [openResident]);

  const updateReadOnly = useCallback((): void => {
    const current = active.current;
    if (!current || current.requestedReadOnly === settings.current.readOnly) return;
    const value = settings.current.readOnly;
    const ticket = current.binding.admission.setReadOnly(value);
    current.requestedReadOnly = value;
    current.barrier = ticket;
    void ticket.ack.then((outcome) => {
      if (active.current !== current || current.barrier !== ticket) return;
      current.barrier = undefined;
      if (outcome.status === 'refused' || outcome.status === 'cancelled') {
        current.requestedReadOnly = current.binding.session.state.editState?.readOnly ?? !value;
        reportRecoverableError(new Error('DOCX read-only transition was rejected'));
      }
    }, (error: unknown) => {
      if (active.current !== current || current.barrier !== ticket) return;
      current.barrier = undefined;
      current.requestedReadOnly = current.binding.session.state.editState?.readOnly ?? !value;
      reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
    });
  }, [reportRecoverableError]);
  const updateRenderEnvironment = useCallback(() => {
    const current = active.current;
    if (!current) return;
    const theme = settings.current.theme ?? renderEnvironment.current.theme;
    const themeColors = Object.fromEntries(Object.entries(theme?.colorScheme ?? {}).filter((entry): entry is [string, string] => typeof entry[1] === 'string'));
    const reserved = reservationsRef.current;
    const reservations = sentReservations();
    const payload = {
      themeColors, showHiddenText: settings.current.showHiddenText, showTableGridlines: settings.current.showTableGridlines, ...(reservations ? { reservations } : {}),
    };
    const requested = JSON.stringify(payload);
    if (current.requestedRenderEnvironment === requested) return;
    const result = current.binding.commands.submit({ kind: 'setRenderEnvironment', payload }, { fence: true, onSubmitted: (ticket) => {
      void ticket.ack.then((outcome) => {
        if (active.current !== current || current.requestedRenderEnvironment !== requested) return;
        if (outcome.status === 'refused' || outcome.status === 'cancelled') {
          current.requestedRenderEnvironment = '';
          reportRecoverableError(new Error(`DOCX render environment was rejected: ${outcome.reason}`));
        }
      }, (error: unknown) => {
        if (active.current !== current) return;
        current.requestedRenderEnvironment = '';
        reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
      });
    } });
    if (result.status !== 'ignored') {
      current.requestedRenderEnvironment = requested;
      reserved.sent = reserved.payload !== null;
    }
  }, [sentReservations, reportRecoverableError]);
  const updateFontEnvironment = useCallback(() => {
    const current = active.current;
    const client = worker.current;
    const opened = fontSource.current;
    if (!current || !client || !opened) return;
    const signature = JSON.stringify(settings.current.fonts ?? []);
    const provider = settings.current.measurementFontProvider;
    if (fontJob.current?.signature === signature && fontJob.current.provider === provider) return;
    if (!fontJob.current && opened.signature === signature && opened.provider === provider) return;
    fontAbort.current?.abort();
    const abort = new AbortController();
    fontAbort.current = abort;
    const job = { signature, provider };
    fontJob.current = job;
    const fonts = settings.current.fonts;
    const isCurrent = () => !abort.signal.aborted && active.current === current && fontSource.current === opened && worker.current === client;
    void (async () => {
      await current.binding.admission.whenCompositionSettled?.();
      await current.binding.session.flush();
      if (!isCurrent()) return;
      let payload: ResidentFontEnvironment;
      if (opened.signature === null && !fonts?.length && provider === opened.provider) payload = structuredClone(opened.initial);
      else {
        const requirements = JSON.parse(await client.fontRequirements(opened.layoutInput)) as ResidentFontRequirement[];
        const embedded = await extractEmbeddedFontFaces(opened.document);
        payload = await prepareWorkerFontEnvironment({ fonts, provider, embedded, requirements, families: opened.families,
          compat: opened.document.package.settings?.compatibilityFlags, signal: abort.signal });
      }
      if (!isCurrent()) return;
      const result = current.binding.commands.submit({ kind: 'setFontEnvironment', payload }, { fence: true, onSubmitted: ticket => {
        void ticket.ack.then(outcome => {
          payload.faces.length = 0;
          if (!isCurrent()) return;
          if (outcome.status === 'refused' || outcome.status === 'cancelled') {
            reportRecoverableError(new Error(`DOCX font environment was rejected: ${outcome.reason}`));
          } else { opened.signature = signature; opened.provider = provider; opened.initial.faces.length = 0; }
          if (fontJob.current === job) fontJob.current = null;
        }, (error: unknown) => {
          payload.faces.length = 0;
          if (isCurrent()) reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
          if (fontJob.current === job) fontJob.current = null;
        });
      } });
      if (result.status === 'ignored') throw new Error('DOCX font environment was not admitted');
    })().catch((error: unknown) => {
      if (isCurrent()) reportRecoverableError(error instanceof Error ? error : new Error(String(error)));
      if (fontJob.current === job) fontJob.current = null;
    });
  }, [reportRecoverableError]);
  prepareFontFamily.current = async family => {
    const current = active.current;
    const client = worker.current;
    const opened = fontSource.current;
    if (!current || !client || !opened) throw new Error('Worker fonts are unavailable');
    fontAbort.current?.abort();
    fontJob.current = null;
    const abort = new AbortController();
    fontAbort.current = abort;
    await current.binding.admission.whenCompositionSettled?.();
    await current.binding.session.flush();
    abort.signal.throwIfAborted();
    const requirements = JSON.parse(await client.fontRequirements(opened.layoutInput)) as ResidentFontRequirement[];
    for (const bold of [false, true]) for (const italic of [false, true]) {
      const key = `${family.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
      if (!requirements.some(entry => entry.key === key)) requirements.push({ key, family, bold, italic });
    }
    const embedded = await extractEmbeddedFontFaces(opened.document);
    const payload = await prepareWorkerFontEnvironment({ fonts: settings.current.fonts,
      provider: settings.current.measurementFontProvider, embedded, requirements, families: [...opened.families, family],
      compat: opened.document.package.settings?.compatibilityFlags, signal: abort.signal });
    if (active.current !== current || fontSource.current !== opened || worker.current !== client) throw new Error('Document was replaced');
    abort.signal.throwIfAborted();
    return payload;
  };
  const onReady = useCallback((binding: PeerlessEditSessionBinding) => {
    controller.attach(binding);
    workerCollaboration.ready(worker.current, binding);
    if (active.current?.binding !== binding) {
      active.current?.clipboard?.dispose();
      active.current = {
        binding, requestedReadOnly: binding.session.state.editState?.readOnly ?? settings.current.readOnly,
        requestedRenderEnvironment: renderEnvironment.current.requested,
      };
      active.current.clipboard = createClipboardPrefetch({ binding, isCurrent: () => active.current?.binding === binding,
        readSelection: (selection) => controller.read({ kind: 'rawSelectionInfo', selection }),
        readClipboardSlice: (selection) => controller.read({ kind: 'clipboardSlice', selection }) });
    }
    updateReadOnly();
    updateRenderEnvironment();
    updateFontEnvironment();
  }, [updateReadOnly, updateRenderEnvironment, updateFontEnvironment, controller, workerCollaboration.ready]);
  useLayoutEffect(updateReadOnly, [effectiveReadOnly, updateReadOnly]);
  useLayoutEffect(updateRenderEnvironment, [props.theme, props.showHiddenText, tableGridlines, reservations, updateRenderEnvironment]);
  useLayoutEffect(updateFontEnvironment, [props.fonts, measurementFontProvider, props.onFontsLoaded, updateFontEnvironment]);
  useLayoutEffect(() => () => { active.current?.clipboard?.dispose(); active.current = null; fontAbort.current?.abort(); fontJob.current = null; fontSource.current = null; }, []);

  const canvasForPage = useCallback((page: DocxCommandPageIdentity) => {
    if (!/^[1-9]\d*$/.test(page.pageId) || !Number.isSafeInteger(page.pageIndex) || page.pageIndex < 0) return null;
    const selector = `[data-page-id="${page.pageId}"] canvas, canvas[data-page-id="${page.pageId}"], ` +
      `[data-page-key="${page.pageId}"] canvas, canvas[data-page-key="${page.pageId}"]`;
    const canvas = host.current?.querySelector<HTMLCanvasElement>(selector);
    if (canvas) return canvas;
    const index = pageIndexes.current.get(page.pageId);
    if (index === undefined) return null;
    return host.current?.querySelector<HTMLCanvasElement>(
      `[data-page-index="${index}"]:not([data-page-id]):not([data-page-key]) canvas:not([data-preview-page-id])`
    ) ?? null;
  }, []);

  const prepareGeometry = useCallback((frame: DocxCommandFrame): PeerlessFrameGeometry | undefined => {
    for (const capture of printCaptures.current) capture.accept(frame);
    if (!identity.current) return undefined;
    const cached = prepared.current.get(frame);
    if (cached) return cached;
    try {
      let next = retained.current;
      for (const delta of decodeCommandFrameDeltas(frame)) next = applyFrameDelta(next, delta);
      if (!next) return undefined;
      const caret = residentCaretSnapshotForFrame(frame.payload.caret, next);
      if (!caret) return undefined;
      const pages = commandPageGeometry(next, frame.payload);
      const geometry: PeerlessFrameGeometry = {
        identity: identity.current,
        coversThroughSeq: frame.coversThroughSeq,
        frameEpoch: next.frameEpoch,
        layoutRevision: frame.payload.layoutRevision,
        documentVersion: frame.payload.documentVersion,
        objects: frame.payload.geometry,
        documentWidth: frame.payload.documentWidth,
        layoutProvisional: frame.payload.layoutProvisional,
        ...(frame.payload.reservations ? { reservations: frame.payload.reservations } : {}),
        pages,
        caret: caret.caretRect ? { ...caret.caretRect,
          x: caret.caretRect.x + (pages[caret.caretRect.pageIndex]?.bounds.x ?? 0), y: caret.caretRect.y + (pages[caret.caretRect.pageIndex]?.bounds.y ?? 0), width: 1 } : null,
        selectionFocus: frame.payload.selectionFocus ? { ...frame.payload.selectionFocus,
          x: frame.payload.selectionFocus.x + (pages[frame.payload.selectionFocus.pageIndex]?.bounds.x ?? 0),
          y: frame.payload.selectionFocus.y + (pages[frame.payload.selectionFocus.pageIndex]?.bounds.y ?? 0), width: 1 } : null,
        parts: frame.payload.parts,
        selectionRects: (frame.payload.selectionRects ?? []).map(({ pageIndex, x, y, width, height }) => ({
          x: x + (pages[pageIndex]?.bounds.x ?? 0), y: y + (pages[pageIndex]?.bounds.y ?? 0), width, height,
        })),
      };
      retained.current = next;
      pageIndexes.current = new Map(next.pages.map(({ pageId, pageIndex }) => [pageId.toString(), pageIndex]));
      prepared.current.set(frame, geometry);
      return geometry;
    } catch {
      return undefined;
    }
  }, []);

  const recovery = useMemo<UsePeerlessEditSessionOptions['recovery']>(() => ({
    checkpoint: source instanceof Uint8Array ? (resident) => resident.client.call.checkpoint() : undefined,
    reopen: async (resume) => {
      const context = openContext.current;
      if (!context) throw new Error('Editor was closed');
      return await openResident(resume.source, context.signal, context.getRasterScale, resume);
    },
    onRecovery: (event) => {
      if (event.phase === 'recovered') {
        updateReadOnly(); updateRenderEnvironment(); updateFontEnvironment();
      }
    },
  }), [source, openResident, updateReadOnly, updateRenderEnvironment, updateFontEnvironment]);
  const sessionOptions = useMemo<UsePeerlessEditSessionOptions | null>(() => source ? {
    source, openConnection, recovery: workerCollaboration.openOptions ? undefined : recovery, canvasForPage, prepareGeometry, onError: reportError,
    onRecoverableError: error => { controller.stalled(error); reportRecoverableError(error); },
  } : null, [source, openConnection, recovery, workerCollaboration.openOptions, canvasForPage, prepareGeometry, reportError, reportRecoverableError, controller]);
  useEffect(() => {
    if (shellState.status !== 'ready' || !previewJob) return;
    setPreviewJob(null);
    releaseWorkerPreview(previewJob.client);
  }, [shellState.status, previewJob]);
  const clipboardReplay = useRef<{ capture: ClipboardCapture; written: boolean } | null>(null);
  const onClipboard = useCallback(async (cut: boolean, data?: DataTransfer | null) => {
    const current = active.current;
    if (!current) return;
    if (data && clipboardReplay.current) {
      try { clipboardReplay.current.written = writeClipboardCapture(data, clipboardReplay.current.capture); } catch {}
      return;
    }
    try {
      const copied = await copyWorkerSelection({ binding: current.binding, cut, data, prefetched: current.clipboard?.take(),
        isCurrent: () => active.current === current,
        readCapture: () => current.clipboard!.read(),
        writeNative: (capture) => {
          const input = host.current?.querySelector('.paged-editor__peerless-input');
          if (!data || active.current !== current || document.activeElement !== input || !navigator.userActivation?.isActive ||
            current.clipboard?.take() !== capture) return false;
          const replay = { capture, written: false };
          clipboardReplay.current = replay;
          let completed = false;
          try { completed = document.execCommand('copy'); } catch {} finally { clipboardReplay.current = null; }
          return completed && replay.written;
        },
        readSelection: (selection) => controller.read({ kind: 'rawSelectionInfo', selection }),
        readClipboardSlice: (selection) => controller.read({ kind: 'clipboardSlice', selection }),
        write: (text) => navigator.clipboard.writeText(text),
        writeSlice: async (text, html) => {
          try {
            if (typeof ClipboardItem === 'undefined' || !navigator.clipboard.write) throw new Error('ClipboardItem unavailable');
            await navigator.clipboard.write([new ClipboardItem({
              'text/plain': new Blob([text], { type: 'text/plain' }),
              'text/html': new Blob([html], { type: 'text/html' }),
            })]);
          } catch {
            await navigator.clipboard.writeText(text);
          }
        } });
      if (copied) { if (cut) props.onCut?.(); else props.onCopy?.(); }
    } catch (error) { reportRecoverableError(error instanceof Error ? error : new Error(String(error))); }
  }, [controller, props.onCopy, props.onCut, reportRecoverableError]);
  const onPasteFragment = useCallback((fragment: DocxClipboardFragment, inputTime: number): boolean => {
    const current = active.current;
    if (!current) return false;
    if (!isPasteFragmentAllowed(controller.commands.store)) return false;
    admitInputIntents(current.binding.admission, [{ kind: 'pasteFragment', payload: { fragment } }],
      { inputTime, source: 'paste', fence: true });
    props.onPaste?.();
    return true;
  }, [controller, props.onPaste]);
  const selectionAnchor = useCallback(() => {
    const frame = controller.getSnapshot().view?.geometry;
    const rect = frame?.selectionRects[0];
    const page = rect && frame.pages.find(({ bounds }) => rect.y >= bounds.y && rect.y < bounds.y + bounds.height);
    if (!rect || !page) return null;
    const bounds = pluginGeometry.geometry()?.dom.getPageBounds(page.page.pageIndex);
    return { y: (bounds?.y ?? page.bounds.y + 48 / zoom) + rect.y - page.bounds.y, right: bounds ? bounds.x + bounds.width : page.bounds.x + page.bounds.width };
  }, [controller, pluginGeometry, zoom]);
  const workerComments = useWorkerComments({ controller, documentKey: shellState.view?.identity.session, readOnly: effectiveReadOnly, setOpen: setCommentsOpen,
    setExpanded: setExpandedSidebarItem, selectionAnchor, focus, onError: reportRecoverableError });
  const contextMenu = useWorkerContextMenu({ binding: () => active.current?.binding ?? null,
    environment: controller.environment, commands: controller.commands.store, clipboard: onClipboard,
    submit: controller.submit, onWrapImage: (handle, wrap) => propertyDialogs.applyImageWrap(wrap, handle),
    onPaste: props.onPaste, onError: reportRecoverableError, beginAddComment: workerComments.beginAddComment });
  useLayoutEffect(() => {
    active.current?.clipboard?.dispose();
    active.current = null;
    setPreviewJob(null);
    setHyperlinkPopup(null);
    linkGesture.current += 1;
    openedSource.current = sourceKey;
    pluginGeometry.clear();
    if (hasSource) controller.opening(sourceGeneration); else controller.empty();
  }, [controller, sourceGeneration, source, sourceKey, hasSource, pluginGeometry]);
  useEffect(() => {
    const request = loadRequest.current;
    if (!request) return;
    if (request.input !== (documentBuffer ?? parsedDocument)) { request.reject(new Error('Document was replaced')); loadRequest.current = null; }
    else if (shellState.error && controller.getSnapshot().error === shellState.error) { request.reject(shellState.error); loadRequest.current = null; }
    else if (shellState.status === 'ready' && controller.getSnapshot().status === 'ready') { request.resolve(); loadRequest.current = null; }
  }, [shellState, documentBuffer, parsedDocument, controller]);
  useImperativeHandle(ref, () => controller.ref, [controller]);

  const onPartPoint: NonNullable<React.ComponentProps<typeof PeerlessPagedEditor>['onPartPoint']> = ({ page, x, y, clickCount, intent }) => {
    const gesture = ++linkGesture.current;
    setHyperlinkPopup(null);
    const geometry = controller.getSnapshot().view?.geometry;
    const section = 'final';
    const targets = geometry?.parts?.filter(part => part.pageId === page.page.pageId &&
      x >= part.x && x <= part.x + part.width && y >= part.y && y <= part.y + part.height) ?? [];
    if (targets.some(part => part.kind === 'footnote' || part.kind === 'endnote')) {
      void parts.openAtPoint(intent, section);
      return true;
    }
    const body = page.displayPage?.contentBounds;
    const band = targets.find(part => part.kind === 'header' || part.kind === 'footer')?.kind ?? (body &&
      (y < body.y ? 'header' : y > body.y + body.height ? 'footer' : null));
    if (band === 'header' || band === 'footer') {
      const activePart = controller.getSnapshot().view?.editState?.activePart;
      if (activePart?.kind !== band || activePart.page !== page.page.pageIndex) {
        if (clickCount === 2) {
          const target = { kind: band, section, page: page.page.pageIndex };
          void (effectiveReadOnly ? parts.open(target, page.displayPage?.sectionIndex) : parts.ensure(target, page.displayPage?.sectionIndex));
        }
        return true;
      }
    }
    const region = band === 'header' || band === 'footer' ? { kind: band } : { kind: 'body' as const };
    const hit = clickCount === 1 && page.displayPage ? findDisplayListHyperlinkAtPoint({ pages: [page.displayPage] }, 0, x, y, region) : null;
    const href = sanitizeHref(hit?.href ?? '');
    const bounds = host.current?.querySelector('[data-peerless-pages-host]')?.getBoundingClientRect();
    const hostBounds = host.current?.getBoundingClientRect();
    if (!hit || !href || !bounds || !hostBounds) return false;
    const data = { href, displayText: hit.displayText, tooltip: hit.tooltip, position: {
      left: bounds.left - hostBounds.left + (page.bounds.x + hit.rect.x) * zoom,
      top: bounds.top - hostBounds.top + (page.bounds.y + hit.rect.y + hit.rect.h) * zoom + 4,
    } };
    void controller.submit({ kind: 'selectPoint', payload: intent }).then(async result => {
      if (!result.ok || gesture !== linkGesture.current) return;
      const selection = controller.getSnapshot().view?.workerSelection;
      if (!selection || selection.anchor.story !== selection.focus.story || selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset) return;
      focus();
      if (href.startsWith('#')) {
        const story = region.kind === 'body' ? 'body' : selection.focus.story;
        await navigation.scrollToBookmark(href.slice(1), story);
      } else setHyperlinkPopup(data);
    }).catch(error => reportRecoverableError(error instanceof Error ? error : new Error(String(error))));
    return true;
  };
  const activeBox = shellState.view?.geometry && parts.activePart && shellState.view.geometry.coversThroughSeq === shellState.view.throughSeq ?
    shellState.view.geometry.parts?.find(part => part.story === parts.activePart?.story && part.pageIndex === parts.activePart.page) : undefined;
  const activePage = activeBox && shellState.view?.geometry?.pages.find(page => page.page.pageId === activeBox.pageId);
  const partChrome = activeBox && activePage && (activeBox.kind === 'header' || activeBox.kind === 'footer') ?
    <InlineHeaderFooterEditor position={activeBox.kind} targetRect={{ top: activePage.bounds.y + activeBox.y,
      left: activePage.bounds.x + activeBox.x, width: activeBox.width, height: activeBox.height }}
      onClose={() => { void parts.close(); }} onRemove={effectiveReadOnly ? undefined : () => { void parts.remove(); }} /> :
    activeBox && activePage ? <div data-docx-escape-layer style={{ position: 'absolute', left: activePage.bounds.x + activeBox.x,
      top: activePage.bounds.y + activeBox.y - 24, zIndex: 10 }}><button type="button" onClick={() => { void parts.close(); }}>Close note</button>{!effectiveReadOnly && <button type="button" onClick={() => { void propertyDialogs.openNotes(); }}>Note properties</button>}</div> : null;
  const ignore = () => {};
  const horizontalRulerProps = {
    sectionProps: rulers.sectionProps, zoom, unit: props.rulerUnit ?? 'inch', editable: rulers.editable,
    paragraphEditable: rulers.paragraphEditable,
    onLeftMarginChange: rulers.handleLeftMarginChange, onRightMarginChange: rulers.handleRightMarginChange,
    indentLeft: rulers.indentLeft, indentRight: rulers.indentRight,
    onIndentLeftChange: rulers.handleIndentLeftChange, onIndentRightChange: rulers.handleIndentRightChange,
    showFirstLineIndent: true, firstLineIndent: rulers.firstLineIndent, hangingIndent: rulers.hangingIndent,
    onFirstLineIndentChange: rulers.handleFirstLineIndentChange, tabStops: rulers.tabStops,
    onTabStopRemove: rulers.handleTabStopRemove, onTabStopsChange: rulers.handleTabStopsChange,
    onDragStart: rulers.startDrag, onDragEnd: () => { void rulers.endDrag(); },
  };
  const chrome = props.showToolbar === false ? null : props.toolbar !== undefined ? props.toolbar : readOnly ? null :
    <DocxEditorToolbar renderLogo={props.renderLogo} documentName={props.documentName}
      onDocumentNameChange={props.onDocumentNameChange} documentNameEditable={props.documentNameEditable}
      renderTitleBarRight={props.renderTitleBarRight} />;
  const geometry = documentReady && pluginGeometry.layout() ? pluginGeometry.geometry() : null;
  const managedItems = managedSidebarItems(pluginHost, activations, anchor => {
    const at = { paraId: anchor.paraId, offset: 0 };
    const target = { kind: 'range' as const, version: anchor.version, range: { story: anchor.story, start: at, end: at, view: 'accepted' as const } };
    if (!geometry || anchor.version !== pluginHost.version()) return null;
    pluginGeometry.requestTarget(target);
    const placed = geometry?.getAnchorGeometry(target);
    if (!placed?.ok) return null;
    const origin = overlayRef.current?.getBoundingClientRect();
    const parent = host.current?.getBoundingClientRect();
    return { position: 0, y: placed.anchor.y / zoom + ((origin?.top ?? 0) - (parent?.top ?? 0)) / zoom };
  });
  const commentReview = useWorkerQuery<'reviewSnapshot'>(controller, { kind: 'reviewSnapshot' });
  const currentReview = controller.reviewSnapshot();
  const review = currentReview ?? commentReview?.value;
  const toolbarMetadata = useWorkerQuery<'toolbarMetadata'>(controller, { kind: 'toolbarMetadata' });
  useLayoutEffect(() => controller.toolbarMetadata(toolbarMetadata), [controller, toolbarMetadata]);
  const revisionRead = useWorkerQuery<'sidebar'>(controller, workerRevisionRequest(commentReview));
  useLayoutEffect(() => controller.revisions(commentReview, revisionRead), [controller, commentReview, revisionRead]);
  const trackedChanges = controller.trackedChanges();
  const revisionCallbacks = useMemo(() => workerRevisionCallbacks(controller), [controller]);
  const comments = props.comments ?? review?.uiComments ?? sourceComments.current;
  const documentKey = shellState.view?.identity.session;
  const dotsMode = props.trackedChanges?.display === 'dots';
  const hasComments = comments.some(comment => comment.parentId == null);
  const hasCards = (dotsMode ? hasComments : comments.length > 0) || (!dotsMode && !!trackedChanges?.entries.length);
  const autoOpened = useRef<string | undefined>(undefined);
  useLayoutEffect(() => {
    if (!hasCards || documentKey === undefined || autoOpened.current === documentKey) return;
    autoOpened.current = documentKey;
    setCommentsOpen(true);
  }, [hasCards, documentKey, setCommentsOpen]);
  const commentCallbacks = useMemo(() => ({ ...revisionCallbacks, ...workerComments.callbacks }), [revisionCallbacks, workerComments.callbacks]);
  const commentItems = useCommentSidebarItems({ comments: commentsOpen ? comments : dotsMode ? comments.filter(comment => comment.parentId != null) : [], trackedChanges: commentsOpen || dotsMode ? trackedChanges?.entries ?? [] : [],
    callbacks: commentCallbacks, showResolved: commentsOpen, isAddingComment: commentsOpen && workerComments.adding, addCommentYPosition: workerComments.addCommentY });
  const positionedComments = commentItems.map(item => {
    if (item.isTemporary) return workerComments.composerHidden ? { ...item, hidden: true } : item;
    const nativeId = Object.entries(review?.commentIds ?? {}).find(([, id]) => item.id === `comment-${id}`)?.[0];
    const firstPage = shellState.view?.geometry?.pages[0]?.page.pageIndex;
    const rect = !documentReady ? undefined : (nativeId ? review?.commentGeometry[nativeId]?.[0] : item.anchorKey ? trackedChanges?.anchors.get(item.anchorKey) : undefined) ??
      (review && item.anchorKey?.startsWith('comment-') && firstPage !== undefined ? { pageIndex: firstPage, y: 0 } : undefined);
    const page = rect && shellState.view?.geometry?.pages.find(page => page.page.pageIndex === rect.pageIndex);
    const bounds = rect && (geometry?.dom.getPageBounds(rect.pageIndex) ?? (page ? { y: page.bounds.y + 48 / zoom } : null));
    return bounds && rect ? { ...item, fixedY: bounds.y + rect.y } : { ...item, hidden: true };
  });
  const selected = shellState.view?.workerSelection;
  const floating = documentReady && !effectiveReadOnly && !workerComments.adding && !!selected &&
    (selected.focus.story === 'body' || selected.focus.story.startsWith('body:')) && selected.anchor.story === selected.focus.story &&
    (selected.anchor.paraId !== selected.focus.paraId || selected.anchor.offset !== selected.focus.offset) &&
    (shellState.view?.geometry?.coversThroughSeq ?? -1) >= (shellState.view?.throughSeq ?? 0) ? selectionAnchor() : null;
  const hostProposalItems = workerProposalSidebarItems(controller, currentReview, geometry,
    documentReady && props.allowHostProposals === true && props.showHostProposalsInSidebar === true && (commentsOpen || dotsMode), reportRecoverableError);
  const reviewItems = [...positionedComments, ...hostProposalItems].filter(item => item.review).sort((first, second) => first.review!.change.from - second.review!.change.from);
  const sidebarItems = mergeSidebarItems([...positionedComments, ...hostProposalItems].filter(item => !dotsMode || !item.review), props.pluginSidebarItems ?? [], managedItems);
  const sidebarOpen = (commentsOpen && (!dotsMode || hasComments || workerComments.adding)) || sidebarItems.some((item) => !item.hidden);
  const anchorPositions = new Map([...positionedComments, ...hostProposalItems].flatMap(item =>
    !item.hidden && item.anchorKey && item.fixedY != null ? [[item.anchorKey, item.fixedY] as const] : []));
  const columnWidth = sidebarItems.every(item => item.hidden || reviewItems.some(review => review.id === item.id)) ? REVIEW_CARD_WIDTH : SIDEBAR_WIDTH;
  const sidebarWidth = sidebarOpen ? SIDEBAR_PAGE_GAP + columnWidth : 0;
  const pages = shellState.view?.geometry?.pages;
  const pageWidth = shellState.view?.geometry?.documentWidth ??
    (pages?.length ? pages.reduce((maximum, { bounds }) => Math.max(maximum, bounds.x + bounds.width), 0) : DEFAULT_PAGE_WIDTH);
  const documentShift = sidebarOpen ? Math.min(sidebarWidth / 2,
    Math.max(0, Math.ceil((pageWidth * zoom) / 2 + sidebarWidth + SIDEBAR_EDGE - hostWidth / 2))) : 0;
  const reviewRects = new Map(reviewItems.map(item => {
    const entry = item.review!;
    const ids = new Set(reviewRevisionIds(entry.change));
    const rects = entry.proposalId ? review?.proposalHighlights[entry.proposalId] ?? []
      : Object.entries(review?.revisionGeometry ?? {}).filter(([id]) => ids.has(yrsIdToNumericId(id))).flatMap(([, rects]) => rects);
    return [item.id, [...rects].sort((first, second) => first.pageIndex - second.pageIndex || first.y - second.y || first.x - second.x)] as const;
  }));
  const selectReviewItem = (item: ReviewSidebarItem) => {
    const revisionId = review?.revisions.find(revision => yrsIdToNumericId(revision.revisionId) === item.review!.change.revisionId)?.revisionId;
    const first = reviewRects.get(item.id)?.[0];
    const element = scroller.current;
    const page = first && pages?.find(page => page.page.pageIndex === first.pageIndex);
    if (element && first && page) element.scrollTop += (host.current?.getBoundingClientRect().top ?? 0) - element.getBoundingClientRect().top + (page.bounds.y + first.y) * zoom - 80;
    if (revisionId) void controller.revealReviewTarget({ kind: 'revision', revisionId }).catch(reportRecoverableError);
  };
  const selectionKey = JSON.stringify(selected);
  const lastReviewSelection = useRef<string | undefined>(undefined);
  useEffect(() => {
    if (!review || (review.revisions.length > 0 && reviewItems.length === 0)) return;
    const key = `${sourceKey}:${documentKey}:${selectionKey}`;
    if (lastReviewSelection.current === key) return;
    if (!selected) { lastReviewSelection.current = key; return; }
    const point = selected.focus;
    const revisions = review.revisions.filter(revision => {
      const range = revision.range;
      return range.story === point.story && ((range.start.paraId === point.paraId && point.offset >= range.start.offset &&
        (range.end.paraId !== point.paraId || point.offset <= range.end.offset)) ||
        (range.end.paraId === point.paraId && range.start.paraId !== point.paraId && point.offset <= range.end.offset));
    });
    const matches = reviewItems.filter(item => revisions.some(revision => reviewRevisionIds(item.review!.change).includes(yrsIdToNumericId(revision.revisionId))));
    if (revisions.length > 0 && matches.length === 0) return;
    lastReviewSelection.current = key;
    const item = matches.find(item => item.id === expandedSidebarItem) ?? matches[0];
    if (item) setExpandedSidebarItem(item.id);
    else setExpandedSidebarItem(current => current && reviewItems.some(item => item.id === current) ? null : current);
  }, [sourceKey, documentKey, selectionKey, selected, review, reviewItems, expandedSidebarItem]);
  const acceptAllReviewItems = () => {
    const shown = [...reviewItems];
    void (async () => {
      const snapshot = controller.reviewSnapshot() ?? (await controller.query<'reviewSnapshot'>({ kind: 'reviewSnapshot' })).value;
      const nativeIds = new Map(snapshot?.revisions.map(revision => [yrsIdToNumericId(revision.revisionId), revision.revisionId]));
      const revisions = shown.filter(item => !item.review!.proposalId);
      const unresolved = revisions.filter(item => !reviewRevisionIds(item.review!.change).some(id => nativeIds.has(id)));
      for (const id of new Set(revisions.flatMap(item => reviewRevisionIds(item.review!.change)))) {
        const revisionId = nativeIds.get(id);
        if (revisionId) {
          const result = await controller.commands.store.execute('reviewAccept', { revisionId });
          if (!result.ok) throw new Error(result.failure.message);
        }
      }
      const proposals = [...new Set(shown.flatMap(item => item.review!.proposalId ? [item.review!.proposalId] : []))];
      if (proposals.length) {
        const current = await controller.ref.getProposals();
        const result = await controller.ref.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
          changes: proposals.map(id => ({ id, state: 'accepted' as const })) });
        if (!result.ok) throw new Error(result.failure.message ?? result.failure.code);
      }
      if (unresolved.length) throw new Error(`Accept all could not resolve revisions ${unresolved.map(item => item.review!.change.revisionId).join(', ')}`);
    })().catch(reportRecoverableError);
  };
  const outlineWidth = navigation.showOutline ? OUTLINE_RESERVED_SPACE : props.showOutlineButton !== false ? OUTLINE_BUTTON_RESERVED_SPACE : 20;
  const minLayoutWidth = pageWidth * zoom + 2 * outlineWidth + sidebarWidth;
  useLayoutEffect(() => { pluginHost.geometryChanged(); }, [pluginHost, minLayoutWidth, documentShift, outlineWidth, zoom]);
  return <DocxCommandProvider commands={controller.commands.store}>
    <DocxEditorShell i18n={props.i18n} isDark={isDark} onEditorError={reportError}
      containerRef={container} scrollContainerRef={scroller} editorContentRef={content}
      className={className} containerStyle={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%', minHeight: 0, minWidth: 0, backgroundColor: 'var(--doc-bg)', ...style }}
      mainContentStyle={{ display: 'flex', flex: 1, minHeight: 0, minWidth: 0 }}
      editorContainerStyle={{ flex: 1, minHeight: 0, minWidth: 0, overflowX: 'auto', overflowY: 'auto', overscrollBehavior: 'contain', position: 'relative', overflowAnchor: 'none' }}
      showRuler={props.showRuler ?? false} readOnlyProp={readOnly} showOutline={navigation.showOutline} showOutlineButton={props.showOutlineButton ?? true}
      sidebarOpen={sidebarOpen} documentShift={documentShift} minLayoutWidth={minLayoutWidth} editorScrollLeft={navigation.pageInfo.scrollLeft} expandedSidebarItem={expandedSidebarItem}
      trackedChanges={trackedChanges?.entries ?? []} onScrollContainerMouseDown={ignore} onEditorBgMouseDown={(event) => {
        if (event.target !== event.currentTarget) return;
        event.preventDefault(); focus();
      }} onEditorContextMenu={contextMenu.open}
      horizontalRulerProps={horizontalRulerProps}
      verticalRulerProps={{ sectionProps: rulers.sectionProps, zoom, unit: props.rulerUnit ?? 'inch', editable: rulers.editable,
        onTopMarginChange: rulers.handleTopMarginChange, onBottomMarginChange: rulers.handleBottomMarginChange,
        onDragStart: rulers.startDrag, onDragEnd: () => { void rulers.endDrag(); } }}
      outlineProps={{ headings: navigation.headings.map((heading) => ({ text: heading.text, level: heading.level, pmPos: heading.position })),
        onHeadingClick: (position) => { const heading = navigation.headings.find((heading) => heading.position === position);
          if (heading) void navigation.scrollToParagraph(heading.paraId).catch((error) => reportRecoverableError(error)); },
        onClose: () => navigation.setShowOutline(false), scrollLeft: navigation.pageInfo.scrollLeft }}
      onToggleOutline={() => navigation.setShowOutline(!navigation.showOutline)} scrollPageInfo={navigation.pageInfo}
      toolbar={<EditorChromeContext.Provider value={{ showZoomControl: props.showZoomControl ?? true, toolbarExtra: props.toolbarExtra }}>{chrome}</EditorChromeContext.Provider>}
      renderDock={(placement, available) => <PluginDock host={pluginHost} placement={placement}
        activations={activations.filter((activation) => activation.plugin.panel?.placement === placement)} available={available} />}
      overlays={<>{contextMenu.element}<PluginOverlays host={pluginHost} activations={activations} target={host.current}
        layerRef={setPluginLayer} heldGeometry={documentReady ? pluginGeometry.geometry() : null} /></>} dialogs={<>{propertyDialogs.element}<FindReplaceDialog isOpen={searchState.isOpen} replaceMode={searchState.replaceMode}
        onClose={(text, options) => { search.close(text, options); focus(); }} currentResult={searchState.result}
        onFind={(text, options) => { void search.find(text, options).catch((error) => reportRecoverableError(error)); return null; }}
        onFindNext={() => { search.host.searchNext(); return search.getSnapshot().result?.matches[search.getSnapshot().result!.currentIndex] ?? null; }}
        onFindPrevious={() => { search.host.searchPrevious(); return search.getSnapshot().result?.matches[search.getSnapshot().result!.currentIndex] ?? null; }}
        onReplace={(text) => search.replace(text).then(count => count > 0).catch((error) => { reportRecoverableError(error); return false; })}
        onReplaceAll={async (query, text, options) => { try { await search.find(query, options); return await search.replace(text, true); }
          catch (error) { reportRecoverableError(error instanceof Error ? error : new Error(String(error))); return 0; } }} /></>} fileInputs={<><input ref={fileIO.imageInput} type="file" accept="image/*" hidden data-testid="docx-image-input" onChange={fileIO.imageChange} /><input ref={fileIO.input} type="file" accept=".docx" hidden data-testid="docx-file-input" onChange={fileIO.change} /></>}
      loadingIndicator={hasSource && !shellState.error && shellState.status !== 'ready' && !shellState.previewPainted ? loadingIndicator ?? <DefaultLoadingIndicator /> : null}
      scrollContainerProps={{ tabIndex: 0, onKeyDown: (event) => {
        if ((event.key !== 'PageUp' && event.key !== 'PageDown') || event.defaultPrevented ||
          event.ctrlKey || event.metaKey || event.altKey || event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229 ||
          !(event.target instanceof HTMLTextAreaElement) || !event.target.classList.contains('paged-editor__peerless-input')) return;
        event.preventDefault();
        event.currentTarget.scrollTop += event.currentTarget.clientHeight * (event.key === 'PageDown' ? 0.9 : -0.9);
      } }}
      pagedArea={<div ref={attachHost} style={{ position: 'relative', paddingTop: 48, paddingBottom: 48, paddingRight: documentShift * 2 }} onMouseDown={(event) => {
        if (event.target === event.currentTarget) { event.preventDefault(); focus(); }
      }}>
      <TrackedChangesReview key={sourceKey} items={reviewItems} options={props.trackedChanges} collaboration={props.collaboration && { ...props.collaboration, clientId: workerCollaboration.connection?.replica.clientId ?? props.collaboration.clientId }}
        pages={pages ?? []} rects={reviewRects} zoom={zoom} pageWidth={pageWidth} sidebarOpen={sidebarOpen}
        hostRef={host} scrollerRef={scroller} activeId={expandedSidebarItem} setActive={setExpandedSidebarItem} select={selectReviewItem}
        readOnly={effectiveReadOnly} acceptAll={acceptAllReviewItems}>
      {floating && <WorkerCommentButton top={floating.y * zoom} left={floating.right * zoom + 24} onAdd={() => { void workerComments.beginAddComment(); }} />}
      {documentReady && sidebarItems.length > 0 && <UnifiedSidebar key={`${sourceKey}:${JSON.stringify(shellState.view?.identity)}`} items={sidebarItems} anchorPositions={anchorPositions} renderedDomContext={geometry?.dom ?? null}
        pageWidth={pageWidth} zoom={zoom} editorContainerRef={scroller} documentShift={documentShift} width={columnWidth} activeItemId={expandedSidebarItem} onExpandedItemChange={setExpandedSidebarItem} />}
      </TrackedChangesReview>
      {documentReady && props.pluginOverlays && <div className="paged-editor__plugin-overlays" style={pluginOverlaysStyles}>{props.pluginOverlays}</div>}
      {shellState.status === 'ready' && <WorkerDocumentReader controller={controller} version={shellState.view?.version ?? 0} />}
      {previewJob && shellState.preview === previewJob.frame &&
        <WorkerDocumentPreview job={previewJob} zoom={zoom} onPresented={onPreviewPresented} />}
      <HyperlinkPopup data={hyperlinkPopup} readOnly={effectiveReadOnly} onClose={() => setHyperlinkPopup(null)}
        onNavigate={href => { window.open(href, '_blank', 'noopener,noreferrer'); }}
        onCopy={async href => {
          try { await navigator.clipboard.writeText(href); } catch {
            const textarea = document.createElement('textarea');
            textarea.value = href; textarea.style.position = 'fixed'; textarea.style.opacity = '0';
            document.body.appendChild(textarea); textarea.select(); document.execCommand('copy'); textarea.remove();
          }
        }}
        onEdit={(displayText, href) => { void propertyDialogs.applyPopupLink(displayText, href, hyperlinkPopup?.tooltip).then(applied => { if (applied) setHyperlinkPopup(null); }); }}
        onRemove={() => { void propertyDialogs.removeLinkAtCaret().then(applied => { if (applied) { setHyperlinkPopup(null); toast('Link removed'); } }); }} />
      {shellState.status === 'failed' && shellState.error ? <ParseError message={shellState.error.message} /> : sessionOptions ? (
        <PeerlessPagedEditor sessionOptions={sessionOptions} zoom={zoom} onReady={onReady} autoFocus={!effectiveReadOnly}
          overlay={workerCollaboration.overlay(displayWindow)}
          displayWindow={displayWindow} onDisplayWindowChange={setDisplayWindow} onClipboard={onClipboard} onPasteFragment={onPasteFragment} highlights={searchState.highlights} onNavigateCell={controller.navigateCell} onPartPoint={onPartPoint}
          pageMirror={props.pageMirror} pageOverlays={<>{partChrome}<WorkerContentControls controller={controller} onError={reportRecoverableError} /></>} />
      ) : hasSource ? null : placeholder ?? <DefaultPlaceholder />}
    </div>} />
  </DocxCommandProvider>;
});
