import { ResidentSourceWorkerClient as ResidentEngineWorkerClient, takePreloadedResidentSourceWorker as takePreloadedResidentEngineWorker } from '@betteroffice/docx/peerless';
import { useWorkerPropertyDialogs } from './useWorkerPropertyDialogs';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  applyFrameDelta,
  CANVAS_PAGE_GAP_PX,
  decodeFrameDelta,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { createRustMeasureSource, type ResidentFontRequirement } from '@betteroffice/docx/layout';
import { extractEmbeddedFontFaces, toArrayBuffer, type DocxInput } from '@betteroffice/docx/utils';
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
import { createDocxPluginHost } from '../../../plugins/createDocxPluginHost';
import { DocxCommandProvider } from '../../../commands/DocxCommandProvider';
import { DocxEditorShell } from '../DocxEditorShell';
import { DocxEditorToolbar } from '../DocxEditorToolbar';
import { EditorChromeContext } from '../../EditorToolbarContext';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from '../../DocxEditorHelpers';
import { PluginDock } from '../../../plugins/PluginPanels';
import { useIsDark } from '../hooks/useIsDark';
import { copyWorkerSelection } from './clipboard';
import { useWorkerFileIO } from './useWorkerFileIO';
import { useKeyboardShortcuts } from '../hooks/useKeyboardShortcuts';
import type { PagedEditorRef } from '../PagedEditor';
import { openPrintWindow } from '@betteroffice/docx';
import { createWorkerPrintCapture, printBitmap } from './printPages';
import { useWorkerContextMenu } from './useWorkerContextMenu';
import { useWorkerNavigation } from './useWorkerNavigation';
import { createWorkerSearch } from './workerSearch';
import { FindReplaceDialog } from '../../dialogs/FindReplaceDialog';
import { WorkerDocumentReader } from './WorkerDocumentReader';
import { DEFAULT_PAGE_WIDTH } from '../internals/styles';
import { OUTLINE_BUTTON_RESERVED_SPACE, OUTLINE_RESERVED_SPACE } from '../../DocumentOutline';
import { SIDEBAR_DOCUMENT_SHIFT } from '../../sidebar/constants';
import type { PeerlessDocumentSource, PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import { WorkerDocumentPreview, type WorkerPreviewJob } from './WorkerDocumentPreview';
import { fontEnvironmentForRegistrations, prepareWorkerFontEnvironment } from './workerFonts';
import type { Document } from '@betteroffice/docx/types/document';
import type { ResidentFontEnvironment } from '@betteroffice/docx/yrs';
import { getPrimaryFontFamily } from '../../ui/fontPickerValue';
import { PeerlessPagedEditor } from './PeerlessPagedEditor';
import type { PeerlessFrameGeometry, PeerlessViewIdentity } from './viewStore';

/** @internal */
export const PeerlessDocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function PeerlessDocxEditor(props, ref) {
  const {
  initialZoom = 1,
  readOnly = false,
  className,
  style,
  placeholder,
  loadingIndicator,
  memoryBudget,
  measurementFontProvider,
  onError,
  } = props;
  const [replacement, setReplacement] = useState<{ input: DocxInput; prop: DocxEditorProps['documentBuffer']; document: DocxEditorProps['document'] } | null>(null);
  const replaced = replacement && replacement.prop === props.documentBuffer && replacement.document === props.document;
  const documentBuffer = replaced ? replacement.input : props.documentBuffer;
  const parsedDocument = documentBuffer ? null : props.document;
  const hasSource = !!(documentBuffer || parsedDocument);
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
  const printCapture = useRef<ReturnType<typeof createWorkerPrintCapture> | null>(null);
  const print = useCallback(async (preview = false) => {
    const popup = openPrintWindow('Print', '');
    let capture: ReturnType<typeof createWorkerPrintCapture> | null = null;
    try {
      if (!popup) throw new Error('Print window was blocked');
      const current = active.current;
      const client = worker.current;
      if (!current || !client) throw new Error('Peer-less worker is not ready');
      capture = createWorkerPrintCapture(current.binding, printBitmap, () => active.current === current && !popup.closed);
      printCapture.current = capture;
      const pages = await capture.collect((indices, epoch) => client.buildPages(indices, epoch));
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
      props.onError?.(error instanceof Error ? error : new Error(String(error)));
      throw error;
    } finally { if (printCapture.current === capture) printCapture.current = null; }
  }, [props.onPrint, props.onError]);
  const prepareFontFamily = useRef<(family: string) => Promise<ResidentFontEnvironment>>(async () => { throw new Error('Worker fonts are unavailable'); });
  const controllerOptions = { ...props, zoom, focus, read, readAt, save, loadBuffer, print, onZoom: setZoom, isDark,
    prepareFontFamily: (family: string) => prepareFontFamily.current(family) };
  const [controller] = useState(() => createPeerlessShellController(controllerOptions));
  const [search] = useState(() => createWorkerSearch(controller, (rect) => {
    const element = scroller.current;
    if (element) element.scrollTop += (host.current?.getBoundingClientRect().top ?? 0) - element.getBoundingClientRect().top + rect.y * controller.ref.getZoom() - element.clientHeight / 2;
  }));
  const searchState = useSyncExternalStore(search.subscribe, search.getSnapshot, search.getSnapshot);
  const fileIO = useWorkerFileIO(props, controller.ref, loadBuffer, () => active.current?.binding ?? null);
  const getBinding = useCallback(() => active.current?.binding ?? null, []);
  const navigation = useWorkerNavigation(controller, getBinding, scroller, host, zoom, props);
  useLayoutEffect(() => controller.update({ ...controllerOptions, search: search.host, perform: async (id) => {
    if (id === 'pageSetup') return propertyDialogs.openPageSetup();
    if (id === 'watermark') return propertyDialogs.openWatermark();
    if (id === 'insertLink') return propertyDialogs.openLink();
    if (id === 'find' || id === 'replace') { search.open(id === 'replace'); return { ok: true, status: 'opened' }; }
    if (id === 'print') { await print(); return { ok: true, status: 'executed' }; }
    return fileIO.perform(id);
  }, currentPage: navigation.readCurrentPage, scrollToPage: navigation.scrollToPage, scrollToParagraph: navigation.scrollToParagraph }));
  useImperativeHandle(ref, () => controller.ref, [controller]);
  const shellState = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const effectiveReadOnly = readOnly || (props.mode ?? shellState.mode) === 'viewing';
  const propertyDialogs = useWorkerPropertyDialogs(controller, props);
  const legacyRef = useRef<PagedEditorRef | null>(null);
  useKeyboardShortcuts({ commands: controller.commands, pagedEditorRef: legacyRef, containerRef: container,
    disableFindReplaceShortcuts: props.disableFindReplaceShortcuts ?? false });
  useEffect(() => {
    const request = loadRequest.current;
    if (!request) return;
    if (request.input !== documentBuffer) { request.reject(new Error('Document was replaced')); loadRequest.current = null; }
    else if (shellState.error) { request.reject(shellState.error); loadRequest.current = null; }
    else if (shellState.status === 'ready' && controller.getSnapshot().status === 'ready') { request.resolve(); loadRequest.current = null; }
  }, [shellState, documentBuffer, controller]);
  const [pluginHost] = useState(() => createDocxPluginHost({
    pagedEditorRef: { current: null }, writeMode: () => controller.environment().readOnly ? 'viewing' : controller.environment().mode,
    commands: () => controller.commands, layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: controller.subscribe, geometry: () => null, translate: (key) => controller.environment().translate(key),
    workerRead: {
      version: async () => { const result = await controller.ref.readParagraphs({ view: 'accepted', paraIds: [] });
        return 'version' in result ? { ok: true, version: result.version } : result; },
      readParagraphs: controller.ref.readParagraphs, findText: controller.ref.findText,
      listStories: () => controller.read({ kind: 'listStories' }),
      readStories: (request) => controller.read({ kind: 'readStories', request }),
    },
  }));
  const activations = useSyncExternalStore(pluginHost.subscribe, pluginHost.activations, pluginHost.activations);
  useEffect(() => {
    pluginHost.setReporter(props.onPluginError);
    pluginHost.setGrants(props.pluginGrants);
    pluginHost.setPlugins(props.plugins);
    return pluginHost.subscribe(() => pluginHost.syncCommands());
  }, [pluginHost, props.plugins, props.pluginGrants, props.onPluginError]);
  useEffect(() => {
    if (shellState.status === 'ready') {
      if (pluginHost.generation() === null) pluginHost.openWorker(String(shellState.view?.version ?? 0));
      else pluginHost.workerChanged(String(shellState.view?.version ?? 0));
      pluginHost.modeChanged(shellState.mode, readOnly);
      pluginHost.syncCommands();
    } else pluginHost.close('document-replaced');
  }, [shellState, pluginHost, props.mode, readOnly]);
  useEffect(() => () => { loadRequest.current?.reject(new Error('Editor was closed')); loadRequest.current = null;
    pluginHost.close('unmounted'); search.dispose(); controller.dispose(); }, [controller, pluginHost, search]);
  const identity = useRef<PeerlessViewIdentity | null>(null);
  const retained = useRef<RetainedFrame | null>(null);
  const pageIndexes = useRef(new Map<string, number>());
  const prepared = useRef(new WeakMap<DocxCommandFrame, PeerlessFrameGeometry>());
  const [displayWindow, setDisplayWindow] = useState<readonly [number, number]>([0, 3]);
  const renderEnvironment = useRef({ theme: null as DocxEditorProps['theme'], requested: '' });
  const fontSource = useRef<{ document: Document; layoutInput: string; families: string[]; initial: ResidentFontEnvironment;
    signature: string | null; provider: DocxEditorProps['measurementFontProvider'] } | null>(null);
  const fontAbort = useRef<AbortController | null>(null);
  const fontJob = useRef<{ signature: string; provider: DocxEditorProps['measurementFontProvider'] } | null>(null);
  const workerOnlyFonts = useRef(new Set<string>());
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
  const errorHandler = useRef(onError);
  errorHandler.current = onError;
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
      const opened = await client.openSource({ ...(input instanceof Uint8Array ? { bytes: input } : { document: input }),
        heapLimitBytes: settings.current.workerLimitBytes, preview: settings.current.previewFirstPage === true,
        mediaTokens: settings.current.mediaTokens === true });
      signal.throwIfAborted();
      if (opened.preview) {
        controller.preview(opened.preview);
        await new Promise<void>((resolve, reject) => {
          const cancel = () => reject(signal.reason);
          signal.addEventListener('abort', cancel, { once: true });
          const settle = (operation: () => void) => { signal.removeEventListener('abort', cancel); operation(); };
          setPreviewJob({ client, frame: opened.preview!, signal,
            resolve: () => settle(resolve), reject: (error) => settle(() => reject(error)) });
        });
        signal.throwIfAborted();
      }
      const metadata = opened.document;
      sourceComments.current = metadata.package.document.comments ?? [];
      const theme = settings.current.theme ?? metadata.package.theme;
      const themeColors = Object.fromEntries(Object.entries(theme?.colorScheme ?? {}).filter((entry): entry is [string, string] => typeof entry[1] === 'string'));
      const environment = { themeColors, showHiddenText: settings.current.showHiddenText };
      renderEnvironment.current = { theme: metadata.package.theme, requested: JSON.stringify(environment) };
      const layout = buildResidentRegionLayoutRequest(metadata, 24, { ...environment,
        defaultTabStopTwips: metadata.package.settings?.defaultTabStop ?? null, mediaTokens: settings.current.mediaTokens === true });
      const requirements = JSON.parse(await client.fontRequirements(JSON.stringify(layout))) as ResidentFontRequirement[];
      signal.throwIfAborted();
      measurement = createRustMeasureSource({
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
      const initial = await client.commandInit({
        session: crypto.randomUUID(), incarnation: 1, epoch: 1, readOnly: settings.current.readOnly,
        fonts: fonts.slice(),
        rasterScale: getRasterScale?.() ?? 1,
        fontRegistrationSlots: { fontChains: config.fontChains },
        layoutInput: JSON.stringify(layout),
        layoutWithRegions: true, layoutRevision: 1, extras: JSON.stringify({ fontChains: config.fontChains }), displayWindow: [0, 3],
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
        errorHandler.current?.(new Error('DOCX read-only transition was rejected'));
      }
    }, (error: unknown) => {
      if (active.current !== current || current.barrier !== ticket) return;
      current.barrier = undefined;
      current.requestedReadOnly = current.binding.session.state.editState?.readOnly ?? !value;
      errorHandler.current?.(error instanceof Error ? error : new Error(String(error)));
    });
  }, []);
  const updateRenderEnvironment = useCallback(() => {
    const current = active.current;
    if (!current) return;
    const theme = settings.current.theme ?? renderEnvironment.current.theme;
    const themeColors = Object.fromEntries(Object.entries(theme?.colorScheme ?? {}).filter((entry): entry is [string, string] => typeof entry[1] === 'string'));
    const payload = { themeColors, showHiddenText: settings.current.showHiddenText };
    const requested = JSON.stringify(payload);
    if (current.requestedRenderEnvironment === requested) return;
    const result = current.binding.commands.submit({ kind: 'setRenderEnvironment', payload }, { fence: true, onSubmitted: (ticket) => {
      void ticket.ack.then((outcome) => {
        if (active.current !== current || current.requestedRenderEnvironment !== requested) return;
        if (outcome.status === 'refused' || outcome.status === 'cancelled') {
          current.requestedRenderEnvironment = '';
          errorHandler.current?.(new Error(`DOCX render environment was rejected: ${outcome.reason}`));
        }
      }, (error: unknown) => {
        if (active.current !== current) return;
        current.requestedRenderEnvironment = '';
        errorHandler.current?.(error instanceof Error ? error : new Error(String(error)));
      });
    } });
    if (result.status !== 'ignored') current.requestedRenderEnvironment = requested;
  }, []);
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
      for (const font of fonts ?? []) workerOnlyFonts.current.add(font.family.toLowerCase());
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
            errorHandler.current?.(new Error(`DOCX font environment was rejected: ${outcome.reason}`));
          } else { opened.signature = signature; opened.provider = provider; opened.initial.faces.length = 0; }
          if (fontJob.current === job) fontJob.current = null;
        }, (error: unknown) => {
          payload.faces.length = 0;
          if (isCurrent()) errorHandler.current?.(error instanceof Error ? error : new Error(String(error)));
          if (fontJob.current === job) fontJob.current = null;
        });
      } });
      if (result.status === 'ignored') throw new Error('DOCX font environment was not admitted');
    })().catch((error: unknown) => {
      if (isCurrent()) errorHandler.current?.(error instanceof Error ? error : new Error(String(error)));
      if (fontJob.current === job) fontJob.current = null;
    });
  }, []);
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
  useLayoutEffect(updateRenderEnvironment, [props.theme, props.showHiddenText, updateRenderEnvironment]);
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
      `[data-page-index="${index}"]:not([data-page-id]):not([data-page-key]) canvas`
    ) ?? null;
  }, []);

  const prepareGeometry = useCallback((frame: DocxCommandFrame): PeerlessFrameGeometry | undefined => {
    printCapture.current?.accept(frame);
    if (!identity.current) return undefined;
    const cached = prepared.current.get(frame);
    if (cached) return cached;
    try {
      let next = retained.current;
      for (const bytes of frame.payload.pageFrames ?? []) next = applyFrameDelta(next, decodeFrameDelta(bytes));
      next = applyFrameDelta(next, decodeFrameDelta(frame.payload.frame));
      const caret = residentCaretSnapshotForFrame(frame.payload.caret, next);
      if (!caret) return undefined;
      let y = 0;
      const width = next.pages.reduce((maximum, { page }) => Math.max(maximum, page.width), 0);
      const pages = next.pages.map(({ pageId, pageIndex, page }): PeerlessFrameGeometry['pages'][number] => {
        const top = y;
        const left = (width - page.width) / 2;
        y += page.height + CANVAS_PAGE_GAP_PX;
        return {
          page: { pageId: pageId.toString(), pageIndex },
          bounds: { x: left, y: top, width: page.width, height: page.height },
          transform: [1, 0, 0, 1, left, top],
          displayPage: page,
        };
      });
      const geometry: PeerlessFrameGeometry = {
        identity: identity.current,
        coversThroughSeq: frame.coversThroughSeq,
        frameEpoch: next.frameEpoch,
        layoutRevision: frame.payload.layoutRevision,
        pages,
        caret: caret.caretRect ? { ...caret.caretRect,
          x: caret.caretRect.x + (pages[caret.caretRect.pageIndex]?.bounds.x ?? 0), y: caret.caretRect.y + (pages[caret.caretRect.pageIndex]?.bounds.y ?? 0), width: 1 } : null,
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
    canEchoFont: family => !workerOnlyFonts.current.has(getPrimaryFontFamily(family).toLowerCase()),
  } : null, [source, openConnection, canvasForPage, prepareGeometry, reportError]);
  const onClipboard = useCallback(async (cut: boolean) => {
    const current = active.current;
    if (!current) return;
    try {
      const copied = await copyWorkerSelection({ binding: current.binding, cut,
        isCurrent: () => active.current === current,
        readSelection: (selection) => controller.read({ kind: 'rawSelectionInfo', selection }),
        write: (text) => navigator.clipboard.writeText(text) });
      if (copied) { if (cut) props.onCut?.(); else props.onCopy?.(); }
    } catch (error) { errorHandler.current?.(error instanceof Error ? error : new Error(String(error))); }
  }, [controller, props.onCopy, props.onCut]);
  const contextMenu = useWorkerContextMenu({ binding: () => active.current?.binding ?? null,
    environment: controller.environment, commands: controller.commands.store, clipboard: onClipboard,
    onPaste: props.onPaste, onError: props.onError });
  useLayoutEffect(() => {
    active.current = null;
    setPreviewJob(null);
    if (hasSource) controller.opening(); else controller.empty();
  }, [controller, source, documentBuffer, hasSource]);

  const ignore = () => {};
  const horizontalRulerProps = {
    sectionProps: undefined, zoom, unit: 'inch' as const, editable: false,
    onLeftMarginChange: ignore, onRightMarginChange: ignore, indentLeft: 0, indentRight: 0,
    onIndentLeftChange: ignore, onIndentRightChange: ignore, firstLineIndent: 0,
    hangingIndent: false, onFirstLineIndentChange: ignore, tabStops: null, onTabStopRemove: ignore,
  };
  const chrome = props.showToolbar === false ? null : props.toolbar !== undefined ? props.toolbar : readOnly ? null :
    <DocxEditorToolbar renderLogo={props.renderLogo} documentName={props.documentName}
      onDocumentNameChange={props.onDocumentNameChange} documentNameEditable={props.documentNameEditable}
      renderTitleBarRight={props.renderTitleBarRight} />;
  const sidebarOpen = props.commentsSidebarOpen === true || (props.commentsSidebarOpen !== false && !!props.comments?.length) ||
    (props.pluginSidebarItems ?? []).some((item) => !item.hidden);
  const sidebarWidth = sidebarOpen ? SIDEBAR_DOCUMENT_SHIFT * 2 : 0;
  const pages = shellState.view?.geometry?.pages;
  const pageWidth = pages?.length ? pages.reduce((maximum, { bounds }) => Math.max(maximum, bounds.x + bounds.width), 0) : DEFAULT_PAGE_WIDTH;
  const outlineWidth = navigation.showOutline ? OUTLINE_RESERVED_SPACE : props.showOutlineButton !== false ? OUTLINE_BUTTON_RESERVED_SPACE : 20;
  const minLayoutWidth = pageWidth * zoom + 2 * outlineWidth + sidebarWidth;
  return <DocxCommandProvider commands={controller.commands.store}>
    <DocxEditorShell i18n={props.i18n} isDark={isDark} onEditorError={reportError}
      containerRef={container} scrollContainerRef={scroller} editorContentRef={content}
      className={className} containerStyle={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%', minHeight: 0, minWidth: 0, backgroundColor: 'var(--doc-bg)', ...style }}
      mainContentStyle={{ display: 'flex', flex: 1, minHeight: 0, minWidth: 0 }}
      editorContainerStyle={{ flex: 1, minHeight: 0, minWidth: 0, overflowX: 'auto', overflowY: 'auto', position: 'relative', overflowAnchor: 'none' }}
      showRuler={false} readOnlyProp={readOnly} showOutline={navigation.showOutline} showOutlineButton={props.showOutlineButton ?? true}
      sidebarOpen={sidebarOpen} minLayoutWidth={minLayoutWidth} editorScrollLeft={navigation.pageInfo.scrollLeft} expandedSidebarItem={null}
      trackedChanges={[]} onScrollContainerMouseDown={ignore} onEditorBgMouseDown={(event) => {
        if (event.target !== event.currentTarget) return;
        event.preventDefault(); focus();
      }} onEditorContextMenu={contextMenu.open}
      horizontalRulerProps={horizontalRulerProps}
      verticalRulerProps={{ sectionProps: undefined, zoom, unit: 'inch', editable: false, onTopMarginChange: ignore, onBottomMarginChange: ignore }}
      outlineProps={{ headings: navigation.headings.map((heading) => ({ text: heading.text, level: heading.level, pmPos: heading.position })),
        onHeadingClick: (position) => { const heading = navigation.headings.find((heading) => heading.position === position);
          if (heading) void navigation.scrollToParagraph(heading.paraId).catch((error) => props.onError?.(error)); },
        onClose: () => navigation.setShowOutline(false), scrollLeft: navigation.pageInfo.scrollLeft }}
      onToggleOutline={() => navigation.setShowOutline(!navigation.showOutline)} scrollPageInfo={navigation.pageInfo}
      toolbar={<EditorChromeContext.Provider value={{ showZoomControl: props.showZoomControl ?? true, toolbarExtra: props.toolbarExtra }}>{chrome}</EditorChromeContext.Provider>}
      renderDock={(placement, available) => <PluginDock host={pluginHost} placement={placement}
        activations={activations.filter((activation) => activation.plugin.panel?.placement === placement)} available={available} />}
      overlays={contextMenu.element} dialogs={<>{propertyDialogs.element}<FindReplaceDialog isOpen={searchState.isOpen} replaceMode={searchState.replaceMode}
        onClose={() => search.close()} currentResult={searchState.result}
        onFind={(text, options) => { void search.find(text, options).catch((error) => props.onError?.(error)); return null; }}
        onFindNext={() => { search.host.searchNext(); return search.getSnapshot().result?.matches[search.getSnapshot().result!.currentIndex] ?? null; }}
        onFindPrevious={() => { search.host.searchPrevious(); return search.getSnapshot().result?.matches[search.getSnapshot().result!.currentIndex] ?? null; }}
        onReplace={(text) => search.replace(text).then(count => count > 0).catch((error) => { props.onError?.(error); return false; })}
        onReplaceAll={async (query, text, options) => { try { await search.find(query, options); return await search.replace(text, true); }
          catch (error) { props.onError?.(error instanceof Error ? error : new Error(String(error))); return 0; } }} /></>} fileInputs={<input ref={fileIO.input} type="file" accept=".docx" hidden data-testid="docx-file-input" onChange={fileIO.change} />}
      loadingIndicator={hasSource && !shellState.error && shellState.status !== 'ready' && !shellState.previewPainted ? loadingIndicator ?? <DefaultLoadingIndicator /> : null}
      scrollContainerProps={{ tabIndex: 0, onKeyDown: (event) => {
        if ((event.key !== 'PageUp' && event.key !== 'PageDown') || event.defaultPrevented ||
          event.ctrlKey || event.metaKey || event.altKey || event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229 ||
          !(event.target instanceof HTMLTextAreaElement) || !event.target.classList.contains('paged-editor__peerless-input')) return;
        event.preventDefault();
        event.currentTarget.scrollTop += event.currentTarget.clientHeight * (event.key === 'PageDown' ? 0.9 : -0.9);
      } }}
      pagedArea={<div ref={host} style={{ paddingTop: 48, paddingBottom: 48, paddingRight: sidebarWidth }} onMouseDown={(event) => {
        if (event.target === event.currentTarget) { event.preventDefault(); focus(); }
      }}>
      {shellState.status === 'ready' && <WorkerDocumentReader controller={controller} version={shellState.view?.version ?? 0} />}
      {previewJob && shellState.preview === previewJob.frame && !shellState.error &&
        <WorkerDocumentPreview job={previewJob} zoom={zoom} onPresented={onPreviewPresented} />}
      {shellState.error ? <ParseError message={shellState.error.message} /> : sessionOptions ? (
        <PeerlessPagedEditor sessionOptions={sessionOptions} zoom={zoom} onReady={onReady} autoFocus={!effectiveReadOnly}
          displayWindow={displayWindow} onDisplayWindowChange={setDisplayWindow} onClipboard={onClipboard} highlights={searchState.highlights} onNavigateCell={controller.navigateCell} />
      ) : hasSource ? null : placeholder ?? <DefaultPlaceholder />}
    </div>} />
  </DocxCommandProvider>;
});
