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
import { createResidentCommandConnection, discardCommandPresentation } from '@betteroffice/docx/peerless';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { createPeerlessShellController } from './shellController';
import { commandPageGeometry } from './commandPageGeometry';
import { createDocxPluginHost } from '../../../plugins/createDocxPluginHost';
import { usePluginReservations } from '../../../plugins/reservations';
import { DocxCommandProvider } from '../../../commands/DocxCommandProvider';
import { DocxEditorShell } from '../DocxEditorShell';
import { DocxEditorToolbar } from '../DocxEditorToolbar';
import { EditorChromeContext } from '../../EditorToolbarContext';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from '../../DocxEditorHelpers';
import { PluginOverlays } from '../../../plugins/PluginOverlays';
import { managedSidebarItems, mergeSidebarItems } from '../../../plugins/PluginSidebarItems';
import { UnifiedSidebar } from '../../UnifiedSidebar';
import { createWorkerPluginGeometry } from './workerPluginGeometry';
import { workerProposalSidebarItems } from './workerProposalSidebar';
import { PluginDock } from '../../../plugins/PluginPanels';
import { useIsDark } from '../hooks/useIsDark';
import { copyWorkerSelection, isPasteFragmentAllowed } from './clipboard';
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
import { useWorkerNavigation } from './useWorkerNavigation';
import { createWorkerSearch } from './workerSearch';
import { FindReplaceDialog } from '../../dialogs/FindReplaceDialog';
import { WorkerDocumentReader } from './WorkerDocumentReader';
import { DEFAULT_PAGE_WIDTH, pluginOverlaysStyles } from '../internals/styles';
import { OUTLINE_BUTTON_RESERVED_SPACE, OUTLINE_RESERVED_SPACE } from '../../DocumentOutline';
import { SIDEBAR_DOCUMENT_SHIFT } from '../../sidebar/constants';
import type { PeerlessDocumentSource, PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import { peerlessPreviewPaint, releaseWorkerPreview, WorkerDocumentPreview, workerPreviewPages, type WorkerPreviewJob } from './WorkerDocumentPreview';
import { fontEnvironmentForRegistrations, prepareWorkerFontEnvironment } from './workerFonts';
import type { Document } from '@betteroffice/docx/types/document';
import type { ResidentFontEnvironment } from '@betteroffice/docx/yrs';
import { PeerlessPagedEditor } from './PeerlessPagedEditor';
import type { PeerlessFrameGeometry, PeerlessViewIdentity } from './viewStore';

const PREVIEW_CUT = { blocks: 200, paragraphBudget: 256 };

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
  const [replacement, setReplacement] = useState<{ input: DocxInput; prop: DocxEditorProps['documentBuffer']; document: DocxEditorProps['document'] } | null>(null);
  const replaced = replacement && replacement.prop === props.documentBuffer && replacement.document === props.document;
  const documentBuffer = replaced ? replacement.input : props.documentBuffer;
  const parsedDocument = documentBuffer ? null : props.document;
  const hasSource = !!(documentBuffer || parsedDocument);
  const sourceKey = useMemo(() => crypto.randomUUID(), [documentBuffer, parsedDocument]);
  const openedSource = useRef(sourceKey);
  const loadRequest = useRef<{ input: DocxInput; resolve(): void; reject(error: Error): void } | null>(null);
  const loadBuffer = useCallback((input: DocxInput) => {
    loadRequest.current?.reject(new Error('Document was replaced'));
    return new Promise<void>((resolve, reject) => {
      loadRequest.current = { input, resolve, reject };
      setReplacement({ input, prop: props.documentBuffer, document: props.document });
    });
  }, [props.documentBuffer, props.document]);
  const host = useRef<HTMLDivElement>(null);
  const container = useRef<HTMLDivElement>(null);
  const scroller = useRef<HTMLDivElement>(null);
  const content = useRef<HTMLDivElement>(null);
  const isDark = useIsDark(props.colorMode ?? 'light');
  const worker = useRef<ResidentEngineWorkerClient | null>(null);
  const sourceComments = useRef<NonNullable<DocxEditorProps['comments']>>([]);
  const [zoom, setZoom] = useState(initialZoom);
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
    return (await worker.current.save({ comments: [...(props.comments ?? sourceComments.current)] })).bytes;
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
  const controllerOptions = { ...recoverableProps, zoom, focus, read, readAt, save, loadBuffer, print, onZoom: setZoom, isDark,
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
  useLayoutEffect(() => controller.update({ ...controllerOptions, imagePropertiesAvailable: propertyDialogs.imagePropertiesAvailable, search: search.host, pickImage: fileIO.pickImage, openSplitCell: propertyDialogs.openSplitCell, perform: async (id) => {
    if (id === 'pageSetup') return propertyDialogs.openPageSetup();
    if (id === 'watermark') return propertyDialogs.openWatermark();
    if (id === 'insertLink') return propertyDialogs.openLink();
    if (id === 'imageProperties') return propertyDialogs.openImage();
    if (id === 'find' || id === 'replace') { search.open(id === 'replace'); return { ok: true, status: 'opened' }; }
    if (id === 'print') { await print(); return { ok: true, status: 'executed' }; }
    return fileIO.perform(id);
  }, readPositionAtPoint: async (x, y) => { await pluginGeometry.update(); return pluginGeometry.geometry()?.readPositionAtPoint(x, y) ?? null; },
    currentPage: navigation.readCurrentPage, scrollToPage: navigation.scrollToPage, scrollToParagraph: navigation.scrollToParagraph }));
  const shellState = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
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
  useEffect(() => {
    const request = loadRequest.current;
    if (!request) return;
    if (request.input !== documentBuffer) { request.reject(new Error('Document was replaced')); loadRequest.current = null; }
    else if (shellState.error && controller.getSnapshot().error === shellState.error) { request.reject(shellState.error); loadRequest.current = null; }
    else if (shellState.status === 'ready' && controller.getSnapshot().status === 'ready') { request.resolve(); loadRequest.current = null; }
  }, [shellState, documentBuffer, controller]);
  const [overlayLayer, setOverlayLayer] = useState<HTMLDivElement | null>(null);
  const overlayRef = useRef<HTMLDivElement | null>(null);
  const [geometryRevision, redrawGeometry] = useState(0);
  const sentReservationsRef = useRef<() => DocxReservationsPayload | null>(() => null);
  const [pluginGeometry] = useState(() => createWorkerPluginGeometry(controller, () => host.current, () => overlayRef.current, () => redrawGeometry(value => value + 1),
    (target, focus) => navigation.revealParagraph(target.paraId, target.story, focus), () => sentReservationsRef.current()));
  const setPluginLayer = useCallback((element: HTMLDivElement | null) => { overlayRef.current = element; setOverlayLayer(element); }, []);
  const [expandedSidebarItem, setExpandedSidebarItem] = useState<string | null>(null);
  const [pluginHost] = useState(() => createDocxPluginHost({
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
    workerEdits: { applyEdits: (request, authorize) => controller.applyEdits(request, authorize) },
    workerNavigation: pluginGeometry.navigation,
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
      if (pluginHost.generation() === null) pluginHost.openWorker(shellState.view?.editState?.documentVersion ?? shellState.view?.geometry?.documentVersion ?? String(shellState.view?.version ?? 0));
      else pluginHost.workerChanged(shellState.view?.editState?.documentVersion ?? shellState.view?.geometry?.documentVersion ?? String(shellState.view?.version ?? 0));
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
    theme: props.theme, showHiddenText: props.showHiddenText ?? false, previewFirstPage: props.previewFirstPage, mediaTokens: props.mediaTokens, fonts: props.fonts, reportFonts: !!props.onFontsLoaded });
  useLayoutEffect(() => {
    settings.current = { readOnly: effectiveReadOnly, workerLimitBytes: memoryBudget?.workerLimitBytes, measurementFontProvider,
      theme: props.theme, showHiddenText: props.showHiddenText ?? false, previewFirstPage: props.previewFirstPage, mediaTokens: props.mediaTokens, fonts: props.fonts, reportFonts: !!props.onFontsLoaded };
  }, [effectiveReadOnly, memoryBudget?.workerLimitBytes, measurementFontProvider, props.theme, props.showHiddenText, props.previewFirstPage, props.mediaTokens, props.fonts, props.onFontsLoaded]);
  type Ticket = ReturnType<PeerlessEditSessionBinding['session']['submit']>;
  const active = useRef<{
    binding: PeerlessEditSessionBinding;
    requestedReadOnly: boolean;
    requestedRenderEnvironment: string;
    barrier?: Ticket;
  } | null>(null);
  const [loaded, setLoaded] = useState<{ input: DocxInput; source: Uint8Array } | null>(null);
  const reportError = useCallback((error: Error) => { controller.fail(error); errorHandler.current?.(error); }, [controller]);
  const [previewJob, setPreviewJob] = useState<WorkerPreviewJob | null>(null);
  const onPreviewPresented = useCallback((frame: WorkerPreviewJob['frame']) => {
    controller.previewPresented(frame);
    if (applyFrameDelta(null, decodeFrameDelta(frame.frame)).pages.some(({ page }) => !page.unbuilt)) controller.painted();
  }, [controller]);
  const source = useMemo(() => {
    if (parsedDocument) return parsedDocument;
    if (documentBuffer instanceof Uint8Array) return new Uint8Array(documentBuffer);
    if (documentBuffer instanceof ArrayBuffer) return new Uint8Array(documentBuffer.slice(0));
    return loaded && loaded.input === documentBuffer ? loaded.source : null;
  }, [documentBuffer, parsedDocument, loaded]);

  useEffect(() => {
    if (!documentBuffer || documentBuffer instanceof Uint8Array || documentBuffer instanceof ArrayBuffer) return;
    let current = true;
    void toArrayBuffer(documentBuffer).then((buffer) => {
      if (current) setLoaded({ input: documentBuffer, source: new Uint8Array(buffer) });
    }, (error: unknown) => {
      if (current) reportError(error instanceof Error ? error : new Error(String(error)));
    });
    return () => { current = false; };
  }, [documentBuffer, reportError]);

  const openConnection = useCallback(async (input: PeerlessDocumentSource, signal: AbortSignal, getRasterScale?: () => number): Promise<ResidentCommandConnection> => {
    signal.throwIfAborted();
    active.current = null;
    const client = takePreloadedResidentEngineWorker() ?? new ResidentEngineWorkerClient();
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
    const abort = () => {
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
        const environment = { themeColors, showHiddenText: settings.current.showHiddenText };
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
      if (settings.current.previewFirstPage === true && input instanceof Uint8Array && settings.current.mediaTokens !== true) {
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
        opened = await client.openSource({ ...(input instanceof Uint8Array ? { bytes: input } : { document: input }),
          heapLimitBytes: settings.current.workerLimitBytes, mediaTokens: settings.current.mediaTokens === true });
      } catch (error) {
        if (!painted) throw error;
        signal.throwIfAborted();
        const failure = error instanceof Error ? error : new Error(String(error));
        controller.failOpening(failure);
        errorHandler.current?.(failure);
        return await new Promise<never>((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
      }
      signal.throwIfAborted();
      const metadata = opened.document;
      sourceComments.current = metadata.package.document.comments ?? [];
      const opening = await measured(metadata);
      const { environment, layout, requirements, config } = opening;
      renderEnvironment.current = { theme: metadata.package.theme, requested: JSON.stringify(environment) };
      if (settings.current.previewFirstPage === true && !previewed) {
        try { await withinPreviewBudget((live) => showPreview(opening, live)); } catch (error) { previewFailed(error); }
      }
      const initial = await client.commandInit({
        session: crypto.randomUUID(), incarnation: 1, epoch: 1, readOnly: settings.current.readOnly,
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
      return connection;
    } catch (error) {
      signal.removeEventListener('abort', abort);
      discardCommandPresentation(initialFrame);
      abort();
      throw error;
    } finally {
      disposeFonts();
    }
  }, [controller]);

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
      themeColors, showHiddenText: settings.current.showHiddenText, ...(reservations ? { reservations } : {}),
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
    if (active.current?.binding !== binding) {
      active.current = {
        binding, requestedReadOnly: binding.session.state.editState?.readOnly ?? settings.current.readOnly,
        requestedRenderEnvironment: renderEnvironment.current.requested,
      };
    }
    updateReadOnly();
    updateRenderEnvironment();
    updateFontEnvironment();
  }, [updateReadOnly, updateRenderEnvironment, updateFontEnvironment, controller]);
  useLayoutEffect(updateReadOnly, [effectiveReadOnly, updateReadOnly]);
  useLayoutEffect(updateRenderEnvironment, [props.theme, props.showHiddenText, reservations, updateRenderEnvironment]);
  useLayoutEffect(updateFontEnvironment, [props.fonts, measurementFontProvider, props.onFontsLoaded, updateFontEnvironment]);
  useLayoutEffect(() => () => { active.current = null; fontAbort.current?.abort(); fontJob.current = null; fontSource.current = null; }, []);

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

  const sessionOptions = useMemo<UsePeerlessEditSessionOptions | null>(() => source ? {
    source, openConnection, canvasForPage, prepareGeometry, onError: reportError,
    onRecoverableError: error => { controller.stalled(error); reportRecoverableError(error); },
  } : null, [source, openConnection, canvasForPage, prepareGeometry, reportError, reportRecoverableError, controller]);
  useEffect(() => {
    if (shellState.status !== 'ready' || !previewJob) return;
    setPreviewJob(null);
    releaseWorkerPreview(previewJob.client);
  }, [shellState.status, previewJob]);
  const onClipboard = useCallback(async (cut: boolean) => {
    const current = active.current;
    if (!current) return;
    try {
      const copied = await copyWorkerSelection({ binding: current.binding, cut,
        isCurrent: () => active.current === current,
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
  const contextMenu = useWorkerContextMenu({ binding: () => active.current?.binding ?? null,
    environment: controller.environment, commands: controller.commands.store, clipboard: onClipboard,
    submit: controller.submit, onPaste: props.onPaste, onError: reportRecoverableError });
  useLayoutEffect(() => {
    active.current = null;
    setPreviewJob(null);
    setHyperlinkPopup(null);
    linkGesture.current += 1;
    openedSource.current = sourceKey;
    pluginGeometry.clear();
    if (hasSource) controller.opening(documentBuffer ?? parsedDocument); else controller.empty();
  }, [controller, source, sourceKey, documentBuffer, parsedDocument, hasSource, pluginGeometry]);
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
  const comments = review?.uiComments ?? props.comments ?? sourceComments.current;
  const commentItems = useCommentSidebarItems({ comments: props.commentsSidebarOpen === false ? [] : comments, trackedChanges: [], callbacks: {} });
  const positionedComments = commentItems.map(item => {
    const nativeId = Object.entries(review?.commentIds ?? {}).find(([, id]) => item.id === `comment-${id}`)?.[0];
    const rect = documentReady && nativeId ? review?.commentGeometry[nativeId]?.[0] : undefined;
    const page = rect && shellState.view?.geometry?.pages.find(page => page.page.pageIndex === rect.pageIndex);
    const bounds = rect && (geometry?.dom.getPageBounds(rect.pageIndex) ?? (page ? { y: page.bounds.y + 48 / zoom } : null));
    return bounds && rect ? { ...item, fixedY: bounds.y + rect.y } : { ...item, hidden: true };
  });
  const commentsOpen = props.commentsSidebarOpen === true || (props.commentsSidebarOpen !== false && !!comments.length);
  const hostProposalItems = workerProposalSidebarItems(controller, currentReview, geometry,
    documentReady && props.allowHostProposals === true && props.showHostProposalsInSidebar === true && commentsOpen, reportRecoverableError);
  const sidebarItems = mergeSidebarItems([...positionedComments, ...hostProposalItems], props.pluginSidebarItems ?? [], managedItems);
  const sidebarOpen = commentsOpen || sidebarItems.some((item) => !item.hidden);
  const anchorPositions = new Map([...positionedComments, ...hostProposalItems].flatMap(item =>
    !item.hidden && item.anchorKey && item.fixedY != null ? [[item.anchorKey, item.fixedY] as const] : []));
  const sidebarWidth = sidebarOpen ? SIDEBAR_DOCUMENT_SHIFT * 2 : 0;
  const pages = shellState.view?.geometry?.pages;
  const pageWidth = shellState.view?.geometry?.documentWidth ??
    (pages?.length ? pages.reduce((maximum, { bounds }) => Math.max(maximum, bounds.x + bounds.width), 0) : DEFAULT_PAGE_WIDTH);
  const outlineWidth = navigation.showOutline ? OUTLINE_RESERVED_SPACE : props.showOutlineButton !== false ? OUTLINE_BUTTON_RESERVED_SPACE : 20;
  const minLayoutWidth = pageWidth * zoom + 2 * outlineWidth + sidebarWidth;
  useLayoutEffect(() => { pluginHost.geometryChanged(); }, [pluginHost, minLayoutWidth, sidebarWidth, outlineWidth, zoom]);
  return <DocxCommandProvider commands={controller.commands.store}>
    <DocxEditorShell i18n={props.i18n} isDark={isDark} onEditorError={reportError}
      containerRef={container} scrollContainerRef={scroller} editorContentRef={content}
      className={className} containerStyle={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%', minHeight: 0, minWidth: 0, backgroundColor: 'var(--doc-bg)', ...style }}
      mainContentStyle={{ display: 'flex', flex: 1, minHeight: 0, minWidth: 0 }}
      editorContainerStyle={{ flex: 1, minHeight: 0, minWidth: 0, overflowX: 'auto', overflowY: 'auto', position: 'relative', overflowAnchor: 'none' }}
      showRuler={props.showRuler ?? false} readOnlyProp={readOnly} showOutline={navigation.showOutline} showOutlineButton={props.showOutlineButton ?? true}
      sidebarOpen={sidebarOpen} minLayoutWidth={minLayoutWidth} editorScrollLeft={navigation.pageInfo.scrollLeft} expandedSidebarItem={expandedSidebarItem}
      trackedChanges={[]} onScrollContainerMouseDown={ignore} onEditorBgMouseDown={(event) => {
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
        onClose={() => search.close()} currentResult={searchState.result}
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
      pagedArea={<div ref={host} style={{ position: 'relative', paddingTop: 48, paddingBottom: 48, paddingRight: sidebarWidth }} onMouseDown={(event) => {
        if (event.target === event.currentTarget) { event.preventDefault(); focus(); }
      }}>
      {documentReady && sidebarItems.length > 0 && <UnifiedSidebar key={`${sourceKey}:${JSON.stringify(shellState.view?.identity)}`} items={sidebarItems} anchorPositions={anchorPositions} renderedDomContext={geometry?.dom ?? null}
        pageWidth={pageWidth} zoom={zoom} editorContainerRef={scroller} activeItemId={expandedSidebarItem} onExpandedItemChange={setExpandedSidebarItem} />}
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
          displayWindow={displayWindow} onDisplayWindowChange={setDisplayWindow} onClipboard={onClipboard} onPasteFragment={onPasteFragment} highlights={searchState.highlights} onNavigateCell={controller.navigateCell} onPartPoint={onPartPoint}
          pageMirror={props.pageMirror} pageOverlays={<>{partChrome}<WorkerContentControls controller={controller} onError={reportRecoverableError} /></>} />
      ) : hasSource ? null : placeholder ?? <DefaultPlaceholder />}
    </div>} />
  </DocxCommandProvider>;
});
