import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { CSSProperties, RefObject } from 'react';
import { usePeerlessCommandBinding } from '../commands/usePeerlessCommandBinding';
import { PptxCommandContext } from '../commands/PptxCommandProvider';
import type { PptxCommandStore } from '../commands/types';
import { usePeerlessInsertion, peerlessInsertionOrigin, type PeerlessInsertionOrigin } from './usePeerlessInsertion';
import { PeerlessImagePicker } from './PeerlessImagePicker';
import { fitSlideScale, type SlideFit } from '../fitSlideScale';
import type { PptxAsyncEditorProps } from '../PptxEditor';
import { EditorToolbar } from '../components/EditorToolbar';
import { EditorChromeContext } from '../components/EditorToolbarContext';
import { useTranslation } from '../i18n';
import { PeerlessSlideStrip } from './PeerlessSlideStrip';
import { PeerlessNotes } from './PeerlessNotes';
import { PeerlessRecovery } from './PeerlessRecovery';
import { peerlessClipboard } from './peerlessClipboard';
import { PeerlessReview, type PeerlessReviewState } from './PeerlessReview';
import { peerlessExportPng } from './peerlessFrameSource';
import { PeerlessPresence } from './PeerlessPresence';
import { peerlessObjectCommand } from './peerlessObjectCommands';
import { createPeerlessReplica } from './peerlessCollaboration';
import { navigatePeerlessSlide } from './peerlessNavigation';
import { peerlessHistoryKey } from './peerlessHistory';
import { PeerlessCaret } from './PeerlessCaret';
import { PeerlessTextInput } from './PeerlessTextInput';
import {
  type PeerlessConfirmedFrame,
} from './peerlessViewStore';
import { usePeerlessPresentation, type PeerlessPresentation } from './usePeerlessPresentation';
import { createPeerlessEditorApi, type PptxAsyncEditorApi } from './createPeerlessEditorApi';
import { usePeerlessTextSelection } from './usePeerlessTextSelection';
import { PeerlessShapeOverlay } from './PeerlessShapeOverlay';
import { usePeerlessShapeGestures, type PeerlessShapePreviewStore } from './usePeerlessShapeGestures';

function useStageFit(stage: RefObject<HTMLDivElement | null>, run: PeerlessPresentation | null,
  chrome: PeerlessPresentation | null,
  toolbar: unknown, showToolbar: boolean | undefined, zoom: number | 'fit') {
  const owner = useRef(run);
  owner.current = run;
  const generation = useRef(0);
  const measured = useRef<{ run: PeerlessPresentation; token: SlideFit; width: number; height: number } | null>(null);
  const [fit, setFit] = useState<SlideFit | null>(null);
  const measure = useCallback((width?: number, height?: number) => {
    const element = stage.current;
    const current = owner.current;
    if (!element || !current?.current) return;
    const slide = current.store.getSnapshot().confirmed?.displayList;
    if (!slide) return;
    width ??= element.clientWidth;
    height ??= element.clientHeight;
    const scale = zoom === 'fit' ? fitSlideScale(width, height, slide) : zoom;
    const dpr = element.ownerDocument.defaultView?.devicePixelRatio ?? 1;
    const previous = measured.current;
    if (previous?.run === current && previous.width === width && previous.height === height &&
      previous.token.scale === scale && previous.token.dpr === dpr) return;
    const token = { scale, dpr, generation: ++generation.current };
    measured.current = { run: current, token, width, height };
    current.fit(token);
    setFit(token);
  }, [stage, zoom]);
  useLayoutEffect(() => {
    const element = stage.current;
    if (!element) return;
    const observer = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver((entries) => {
      const entry = entries.find((value) => value.target === element);
      if (entry) measure(entry.contentRect.width, entry.contentRect.height);
    });
    observer?.observe(element);
    const resize = () => measure();
    window.addEventListener('resize', resize);
    return () => { observer?.disconnect(); window.removeEventListener('resize', resize); };
  }, [stage, measure]);
  useLayoutEffect(() => { measure(); }, [run, chrome, toolbar, showToolbar, measure]);
  const previewFit = useCallback((frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null => {
    const token = measured.current?.run === run ? measured.current.token : run.getVisualSnapshot()?.fit;
    return token ? { ...token, scale: fitSlideScale(128, Number.MAX_VALUE, frame.displayList, 0) } : null;
  }, [fit]);
  const current = useCallback((token: SlideFit, run: PeerlessPresentation) => {
    const value = measured.current;
    return value !== null && owner.current === run && value.run === run && run.current && Boolean(value.token.scale) &&
      value.token.generation === token.generation;
  }, []);
  return { previewFit, current };
}

function Overlays({ run, commands, preview, onError }: { run: PeerlessPresentation; commands: PptxCommandStore; preview: PeerlessShapePreviewStore; onError(error: Error): void }) {
  const snapshot = useSyncExternalStore(run.subscribeVisual, run.getVisualSnapshot, run.getVisualSnapshot);
  if (!snapshot?.view.confirmed) return null;
  const { view, fit } = snapshot;
  const confirmed = view.confirmed!;
  return <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none',
    touchAction: 'none', overflow: 'hidden',
    width: confirmed.displayList.width * fit.scale, height: confirmed.displayList.height * fit.scale }}>
    <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none',
      width: confirmed.displayList.width, height: confirmed.displayList.height,
      transform: `scale(${fit.scale})`, transformOrigin: '0 0' }}>
      <PeerlessCaret store={run.visual} />
    </div>
    {view.selection?.kind === 'object' && <PeerlessShapeOverlay selection={view.selection} preview={preview}
      scale={fit.scale} readOnly={view.editState.readOnly} />}
    <PeerlessTextInput store={run.store} commands={commands} visual={run.visual} viewport={{ left: 0, top: 0, scale: fit.scale }}
      onShortcut={(shortcut) => {
        if (shortcut === 'copy' || shortcut === 'cut') void peerlessClipboard(run, shortcut === 'cut').catch((error: unknown) => { if (!run.failure) onError(error instanceof Error ? error : new Error(String(error))); });
        if (shortcut === 'save' || shortcut === 'undo' || shortcut === 'redo') void commands.execute(shortcut, null);
      }} />
  </div>;
}

export function PeerlessPptxSurface(props: PptxAsyncEditorProps) {
  const { t } = useTranslation();
  const canvas = useRef<HTMLCanvasElement>(null);
  const stage = useRef<HTMLDivElement>(null);
  const { presentation, pending, failure } = usePeerlessPresentation(props, canvas);
  const chrome = pending ?? presentation;
  const subscribeStructure = useCallback((listener: () => void) => chrome?.store.subscribeStructure(listener) ?? (() => {}), [chrome]);
  const structureSnapshot = useCallback(() => chrome?.client.state.slides ?? null, [chrome]);
  const slides = useSyncExternalStore(subscribeStructure, structureSnapshot, structureSnapshot);
  const empty = Boolean(presentation && presentation.store.getSnapshot().confirmed === null);
  useLayoutEffect(() => {
    if (!empty || !canvas.current) return;
    const surface = canvas.current;
    const ctx = surface.getContext('2d');
    ctx?.setTransform(1, 0, 0, 1, 0, 0);
    ctx?.clearRect(0, 0, surface.width, surface.height);
  }, [empty]);
  const { toolbar, showToolbar } = props;
  const [zoom, setZoom] = useState<number | 'fit'>('fit');
  const toolbarVisible = showToolbar !== false && toolbar !== null && (toolbar !== undefined || !props.readOnly);
  const { previewFit, current } = useStageFit(stage, pending ?? presentation, chrome, toolbar, toolbarVisible, zoom);
  const callbacks = useRef(props);
  callbacks.current = props;
  const api = useRef<PptxAsyncEditorApi | null>(null);
  const focus = useCallback(() => {
    stage.current?.querySelector('textarea')?.focus();
  }, []);
  const navigate = useCallback(() => { stage.current?.focus(); }, []);
  const selection = usePeerlessTextSelection(presentation?.store ?? null, { canvas, presentation, focus });
  const insertion = usePeerlessInsertion(presentation, canvas, focus);
  const [review, setReview] = useState<PeerlessReviewState>({ open: false, selectedId: null, diff: false, presenting: false });
  const subscribeReview = useCallback((listener: () => void) => presentation?.store.subscribeEditState(listener) ?? (() => {}), [presentation]);
  const reviewSnapshot = useCallback(() => presentation?.store.getSnapshot().editState ?? null, [presentation]);
  const proposals = useSyncExternalStore(subscribeReview, reviewSnapshot, reviewSnapshot)?.proposals ?? [];
  useLayoutEffect(() => {
    if (review.selectedId && !proposals.some((proposal) => proposal.id === review.selectedId)) {
      setReview((state) => ({ ...state, selectedId: null, diff: false }));
    }
  }, [proposals, review.selectedId]);
  useEffect(() => { if (!review.diff) focus(); }, [review.diff, focus]);
  const exportPng = useCallback(async () => {
    if (!presentation) return;
    const blob = await peerlessExportPng(presentation);
    presentation.require();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a'); link.href = url; link.download = 'slide.png'; link.click(); URL.revokeObjectURL(url);
  }, [presentation]);
  const picker = useRef<HTMLInputElement>(null);
  const pictureOrigin = useRef<PeerlessInsertionOrigin | null>(null);
  const openPicture = useCallback(() => {
    pictureOrigin.current = peerlessInsertionOrigin(presentation);
    if (pictureOrigin.current) picker.current?.click();
  }, [presentation]);
  const shapes = usePeerlessShapeGestures(presentation?.store ?? null, { canvas, presentation });
  const save = useCallback(async () => {
    try {
      const value = api.current;
      if (!value || !presentation?.current) return;
      const request = callbacks.current.onSaveRequest;
      if (request && await request() !== true) return;
      const bytes = await value.saveAsync();
      if (api.current !== value) return;
      if (callbacks.current.onSave) { callbacks.current.onSave(bytes); return; }
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }));
      const link = document.createElement('a');
      link.href = url;
      link.download = callbacks.current.fileName ?? 'presentation.pptx';
      link.click();
      URL.revokeObjectURL(url);
    } catch (error: unknown) {
      if (!presentation?.session.failure) callbacks.current.onError?.(error instanceof Error ? error : new Error(String(error)));
      throw error;
    }
  }, [presentation]);
  const commands = usePeerlessCommandBinding(presentation, { tool: insertion.tool, setTool: insertion.setTool,
    openPicture, save, focus, i18n: props.i18n, zoom, setZoom, review, proposals, setReview, exportPng });
  useEffect(() => {
    api.current = null;
    if (!presentation?.current || !canvas.current) return;
    const value = createPeerlessEditorApi(presentation, canvas.current, focus, commands);
    api.current = value;
    let current = true;
    let cleanup: void | (() => void);
    void (async () => {
      while (current && presentation.current) {
        await presentation.ready;
        if (current && presentation.current && presentation.hydrated) {
          cleanup = callbacks.current.onReady?.(value);
          return;
        }
      }
    })().catch(() => {});
    return () => { current = false; api.current = null; if (typeof cleanup === 'function') cleanup(); };
  }, [presentation, focus, commands]);
  useEffect(() => {
    const collaboration = props.collaboration;
    if (!presentation?.current || failure || !collaboration?.onReplica) return;
    collaboration.onReplica(createPeerlessReplica(presentation, collaboration.clientId));
    return () => { collaboration.onReplica?.(null); };
  }, [presentation, failure, props.collaboration?.onReplica]);
  const insertionScale = presentation?.getVisualSnapshot()?.fit.scale ?? 1;
  return <PptxCommandContext.Provider value={commands}><div className={props.className} style={styles.editor} data-testid="pptx-peerless-editor">
    {toolbarVisible && <EditorChromeContext.Provider value>
      {toolbar ?? (chrome ? <EditorToolbar mode="commands" /> : null)}
    </EditorChromeContext.Provider>}
    {(!props.file || !presentation || empty || failure) && <div role="status" data-testid="pptx-worker-status" style={{ position: 'absolute' }}>
      {failure?.message ?? t(!props.file ? 'editor.openPrompt' : empty ? 'editor.noSlides' : 'worker.preparing')}
    </div>}
    {presentation && failure && <PeerlessRecovery run={presentation} fileName={props.fileName} />}
    <div style={styles.workspace}>
      <PeerlessImagePicker input={picker} origin={pictureOrigin} onError={(error) => callbacks.current.onError?.(error)} />
      {chrome && slides && slides.length > 0 && <PeerlessSlideStrip run={chrome} previewFit={previewFit} current={current}
        onNavigate={navigate} />}
      <div ref={stage} role="application" aria-label={t('editor.appLabel')} tabIndex={0}
        aria-readonly={Boolean(props.readOnly || failure)} style={styles.stage}
        onKeyDownCapture={(event) => { if (event.key === 'Escape') insertion.cancel(); shapes.onKeyDown(event); }}
        onKeyDown={(event) => {
          if (!presentation?.current || review.diff) return;
          const selected = presentation.store.getSnapshot().selection;
          if ((event.key === 'Delete' || event.key === 'Backspace') && selected?.kind === 'object') {
            event.preventDefault();
            if (presentation.inputEnabled) void peerlessObjectCommand(presentation, { kind: 'remove' }).catch(() => {});
            return;
          }
          peerlessHistoryKey(event, presentation.store);
          if (!event.defaultPrevented) navigatePeerlessSlide(event, presentation);
        }}
        onPointerDown={(event) => { if (!review.diff && !insertion.onPointerDown(event) && !shapes.onPointerDown(event)) selection.onPointerDown(event); }}
        onPointerMove={(event) => { if (!review.diff && !insertion.onPointerMove(event) && !shapes.onPointerMove(event)) selection.onPointerMove(event); }}
        onPointerUp={(event) => { if (!review.diff && !insertion.onPointerUp(event) && !shapes.onPointerUp(event)) selection.onPointerUp(event); }}
        onPointerCancel={(event) => { insertion.onPointerCancel(event); if (!shapes.onPointerCancel(event)) selection.onPointerCancel(event); }}
        onLostPointerCapture={(event) => { insertion.onPointerCancel(event); if (!shapes.onLostPointerCapture(event)) selection.onLostPointerCapture(event); }}>
        <canvas ref={canvas} style={styles.canvas} data-testid="pptx-slide-canvas" />
        {insertion.preview && <div data-testid="pptx-insertion-preview" style={{ position: 'absolute', pointerEvents: 'none',
          left: insertion.preview.x * insertionScale, top: insertion.preview.y * insertionScale,
          width: insertion.preview.width * insertionScale, height: insertion.preview.height * insertionScale, border: '1px dashed #2563eb' }} />}
        {presentation && !failure && props.collaboration?.presence && <PeerlessPresence run={presentation} presence={props.collaboration.presence} />}
        {presentation && !review.diff && <Overlays run={presentation} commands={commands} preview={shapes.preview} onError={(error) => callbacks.current.onError?.(error)} />}
        {presentation && <PeerlessReview run={presentation} state={review} proposals={proposals} commands={commands}
          onError={(error) => callbacks.current.onError?.(error)} exit={() => { setReview({ ...review, presenting: false }); focus(); }} />}
      </div>
    </div>
    {presentation && <PeerlessNotes run={presentation} onError={(error) => callbacks.current.onError?.(error)} />}
  </div></PptxCommandContext.Provider>;
}

const styles: Record<string, CSSProperties> = {
  editor: { display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 },
  workspace: { display: 'flex', flex: 1, minHeight: 0 },
  stage: { position: 'relative', flex: 1, minWidth: 0, minHeight: 0, overflow: 'auto', background: '#e2e8f0', touchAction: 'none' },
  canvas: { display: 'block', touchAction: 'none' },
};
