import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type { PresentationFrame } from '@betteroffice/pptx';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent } from 'react';
import type { PptxWorkerViewerProps } from '../PptxEditor';
import { createPptxCommandController } from '../commands/createPptxCommandStore';
import type { PptxCommandBinding } from '../commands/createPptxCommandStore';
import { DEFAULT_FONT_FAMILIES, DEFAULT_FONT_SIZES } from '../commands/evaluate';
import type { PptxCommandEnvironment } from '../commands/evaluate';
import { PptxCommandContext } from '../commands/PptxCommandProvider';
import type { PptxCommandArgs, PptxCommandResult } from '../commands/types';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { EditorChromeContext } from '../components/EditorToolbarContext';
import type { PptxZoom } from '../components/toolbarTypes';
import { useTranslation } from '../i18n';
import { currentContext } from './sessionPaint';
import { useSessionPresentation, type ViewerSession } from './useSessionPresentation';

export function PptxSessionViewer(props: PptxWorkerViewerProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const railRef = useRef<HTMLElement>(null);
  const hostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [commands] = useState(createPptxCommandController);
  const [zoom, setZoom] = useState<PptxZoom>('fit');
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [dpr, setDpr] = useState(devicePixelRatio);
  const focus = useCallback(() => stageRef.current?.focus(), []);
  const { run, loading, error, notes, reportError } = useSessionPresentation(props, commands.store, focus);
  const frame = run?.current ? run.frame(run.active) : undefined;
  const navigation = run?.navigation;
  const latest = useRef({ run, props, zoom, loading, t, reportError });
  latest.current = { run, props, zoom, loading, t, reportError };
  const pendingSave = useRef<{ run: ViewerSession; promise: Promise<PptxCommandResult> } | null>(null);

  const binding = useMemo<PptxCommandBinding>(() => ({
    environment: (): PptxCommandEnvironment => {
      const { run, zoom, loading, t } = latest.current;
      const slide = run?.session.state.slides[run.active];
      return {
        status: run?.current ? 'ready' : loading ? 'loading' : 'empty',
        readOnly: true, reviewing: false, pendingInput: false, canUndo: false, canRedo: false,
        slide: slide ?? null, layouts: [], text: null, shape: null, tool: 'select', zoom,
        fontFamilies: DEFAULT_FONT_FAMILIES, fontSizes: DEFAULT_FONT_SIZES, proposals: null,
        hostDisabled: (id) => id === 'slideshow' || id === 'exportPng' || id === 'tool', translate: t,
      };
    },
    ordered: () => false,
    admit: async (operation) => operation(),
    perform: (id, args) => {
      const { run, props, zoom } = latest.current;
      if (!run?.current) return { ok: false, failure: { code: 'document-replaced', message: 'Presentation is no longer open' } };
      if (id === 'zoom') {
        const next = (args as PptxCommandArgs['zoom']).scale;
        setZoom(next);
        return { ok: true, status: next === zoom ? 'noop' : 'executed' };
      }
      if (id !== 'save') return { ok: false, failure: { code: 'read-only', message: 'The editor is read-only' } };
      if (pendingSave.current?.run === run) return pendingSave.current.promise;
      const promise = Promise.resolve().then(async (): Promise<PptxCommandResult> => {
        if (props.onSaveRequest && await props.onSaveRequest() !== true) return { ok: true, status: 'requested' };
        if (!run.current) throw new Error('Presentation is no longer open');
        const bytes = await run.session.save();
        if (!run.current) throw new Error('Presentation is no longer open');
        if (props.onSave) props.onSave(bytes);
        else download(bytes, props.fileName ?? 'presentation.pptx');
        return { ok: true, status: 'executed' };
      }).catch((error): PptxCommandResult => {
        if (run.current) latest.current.reportError(error);
        return { ok: false, failure: {
          code: run.current ? 'command-failed' : 'document-replaced', message: String(error),
        } };
      }).finally(() => { if (pendingSave.current?.promise === promise) pendingSave.current = null; });
      pendingSave.current = { run, promise };
      return promise;
    },
    capture: () => latest.current.run?.current ? { generation: 0, viewer: latest.current.run } : null,
    resume: (origin) => 'viewer' in origin && origin.viewer === latest.current.run &&
      latest.current.run?.current ? null : 'document-replaced',
    chrome: () => ({ i18n: latest.current.props.i18n }), focusEditor: focus,
  }), [focus]);
  useLayoutEffect(() => {
    commands.attach(binding);
    return () => commands.detach(binding);
  }, [commands, binding]);
  useLayoutEffect(() => commands.refresh());
  useCommandShortcuts({ commands, containerRef: rootRef, suspended: () => false });

  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const update = () => {
      const ratio = devicePixelRatio();
      setDpr(ratio);
      media?.removeEventListener('change', update);
      media = window.matchMedia(`(resolution: ${ratio}dppx)`);
      media.addEventListener('change', update);
    };
    update();
    return () => media?.removeEventListener('change', update);
  }, []);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const update = () => setViewport({ width: host.clientWidth, height: host.clientHeight });
    update();
    const observer = new ResizeObserver(update);
    observer.observe(host);
    return () => observer.disconnect();
  }, []);
  const scale = zoom === 'fit' ? !frame || viewport.width <= 0 || viewport.height <= 0 ? 1 :
    Math.min(Math.max(1, viewport.width - 40) / Math.max(1, frame.displayList.width),
      Math.max(1, viewport.height - 40) / Math.max(1, frame.displayList.height), 1) : zoom;

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !frame || !run?.current) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    let cancelled = false;
    const current = () => !cancelled && run.current && run.frame(run.active) === frame;
    sizeCanvasForSlide(canvas, frame.displayList, dpr, scale);
    if (canvas.width < 1) canvas.width = 1;
    if (canvas.height < 1) canvas.height = 1;
    void Promise.resolve().then(async () => {
      if (!current()) return;
      const resolveImage = run.images.resolve(frame);
      try {
        await paintSlide(currentContext(ctx, current), frame.displayList, dpr, scale, { resolveImage });
      } finally { resolveImage.release(); }
    }).then(() => { if (current()) run.didPaint(frame); })
      .catch((error) => { if (current()) run.fail(error); });
    return () => { cancelled = true; };
  }, [run, frame, scale, dpr, navigation]);

  useEffect(() => {
    const rail = railRef.current;
    if (!rail || !run?.current || typeof IntersectionObserver === 'undefined') return;
    const observer = new IntersectionObserver((entries) => {
      for (const entry of entries) run.visibility(Number((entry.target as HTMLElement).dataset.slideIndex), entry.isIntersecting);
    }, { root: rail });
    for (const row of rail.querySelectorAll('[data-slide-index]')) observer.observe(row);
    return () => observer.disconnect();
  }, [run]);

  const navigate = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey ||
      (event.target as HTMLElement).closest('input, textarea, select, [contenteditable]')) return;
    const previous = ['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key);
    const next = ['ArrowRight', 'ArrowDown', 'PageDown'].includes(event.key);
    if (!previous && !next) return;
    event.preventDefault();
    if (run?.show(run.active + 1 + (previous ? -1 : 1))) focus();
  };
  return (
    <PptxCommandContext.Provider value={commands.store}>
      <div ref={rootRef} className={props.className} style={styles.root}>
        {props.showToolbar !== false && props.toolbar != null ? (
          <div style={styles.toolbarShell}>
            <EditorChromeContext.Provider value>
              <div style={{ flex: '1 1 auto', minWidth: 0 }}>{props.toolbar}</div>
            </EditorChromeContext.Provider>
          </div>
        ) : null}
        <div style={styles.workspace}>
          <aside ref={railRef} style={styles.slideStrip} aria-label={t('slides.panelLabel')}>
            {run?.session.state.slides.map((slide, index) => {
              const cached = run.frame(index);
              const visible = run.visible.has(index);
              const title = slide.name || t('slides.fallbackTitle', { number: index + 1 });
              return (
                <button type="button" key={slide.id} data-slide-index={index}
                  aria-current={index === run.active ? 'page' : undefined}
                  style={{ ...styles.slideButton, border: index === run.active ? '2px solid #325ee6' : '2px solid transparent' }}
                  onClick={() => { if (run.show(index + 1)) focus(); }}>
                  <span style={styles.slideNumber}>{index + 1}</span>
                  <span style={styles.slidePreview}>
                    {visible && cached && frame && run.isPainted(frame) ? (
                      <Thumbnail run={run} frame={cached} dpr={dpr} />
                    ) : <span style={styles.slideTitle}>{title}</span>}
                  </span>
                </button>
              );
            })}
          </aside>
          <div ref={stageRef} tabIndex={0} style={styles.stage} onKeyDown={navigate}>
            <div ref={hostRef} style={styles.canvasHost}>
              {frame ? <canvas ref={canvasRef} style={styles.canvas} /> : (
                <div style={styles.empty}>{loading ? t('editor.opening') : props.file ? t('editor.noSlides') : t('editor.openPrompt')}</div>
              )}
            </div>
            {error ? <div style={styles.error}>{error}</div> : null}
          </div>
        </div>
        {run?.session.state.slides[run.active] ? (
          <div style={styles.notesPanel}>
            <span style={styles.notesLabel}>{t('notes.panelLabel')}</span>
            <textarea data-testid="pptx-notes-textarea" style={styles.notesTextarea} value={notes}
              disabled aria-label={t('notes.panelLabel')} placeholder={t('notes.placeholder')} />
          </div>
        ) : null}
      </div>
    </PptxCommandContext.Provider>
  );
}

function Thumbnail({ run, frame, dpr }: { run: ViewerSession; frame: PresentationFrame; dpr: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = ref.current;
    const ctx = canvas?.getContext('2d');
    if (!canvas || !ctx || !run.current) return;
    let cancelled = false;
    const current = () => !cancelled && run.current && run.visible.has(frame.slideIndex);
    const scale = 128 / frame.displayList.width;
    sizeCanvasForSlide(canvas, frame.displayList, dpr, scale);
    if (canvas.width < 1) canvas.width = 1;
    if (canvas.height < 1) canvas.height = 1;
    void Promise.resolve().then(async () => {
      if (!current()) return;
      const resolveImage = run.images.resolve(frame);
      try {
        await paintSlide(currentContext(ctx, current), frame.displayList, dpr, scale, { resolveImage });
      } finally { resolveImage.release(); }
    }).catch(() => {});
    return () => { cancelled = true; };
  }, [run, frame, dpr]);
  return <canvas ref={ref} style={{ display: 'block', maxWidth: '100%', height: 'auto' }} aria-hidden="true" />;
}

function devicePixelRatio(): number {
  const ratio = typeof window === 'undefined' ? 1 : window.devicePixelRatio;
  return Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
}

function download(bytes: Uint8Array, name: string): void {
  const blob = new Blob([new Uint8Array(bytes)], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', flexDirection: 'column', width: '100%', height: '100%', minHeight: 480, overflow: 'hidden',
    color: '#172033', background: '#f3f5f8', fontFamily: 'ui-sans-serif, system-ui, sans-serif' },
  toolbarShell: { display: 'flex', alignItems: 'center', flex: '0 0 auto', padding: '4px 0 5px', background: '#ffffff', borderBottom: '1px solid #e2e8f0' },
  workspace: { display: 'flex', flex: 1, minHeight: 0 },
  slideStrip: { width: 184, padding: '14px 10px', overflowY: 'auto', background: '#eef1f5', borderRight: '1px solid #d8dee9', boxSizing: 'border-box' },
  slideButton: { display: 'flex', alignItems: 'flex-start', gap: 7, width: '100%', marginBottom: 12, padding: 4, borderRadius: 5, background: 'transparent', cursor: 'pointer' },
  slideNumber: { width: 18, flex: '0 0 auto', paddingTop: 3, fontSize: 11, color: '#647087', textAlign: 'right' },
  slidePreview: { position: 'relative', display: 'flex', alignItems: 'center', justifyContent: 'center', width: 132, aspectRatio: '16 / 9', padding: 8,
    overflow: 'hidden', background: '#ffffff', boxSizing: 'border-box', boxShadow: '0 1px 4px rgba(20, 31, 50, 0.16)' },
  slideTitle: { fontSize: 9, lineHeight: 1.25, color: '#39445a', textAlign: 'center' },
  stage: { position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0, outline: 'none', overflow: 'hidden' },
  canvasHost: { display: 'flex', flex: 1, minHeight: 0, alignItems: 'center', justifyContent: 'center', width: '100%', overflow: 'auto' },
  canvas: { display: 'block', flex: '0 0 auto', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' },
  notesPanel: { flex: '0 0 auto', display: 'flex', flexDirection: 'column', gap: 4, padding: '8px 14px 10px', background: '#ffffff', borderTop: '1px solid #e2e8f0' },
  notesLabel: { fontSize: 11, fontWeight: 650, color: '#647087', textTransform: 'uppercase', letterSpacing: '0.02em' },
  notesTextarea: { width: '100%', height: 64, resize: 'vertical', border: '1px solid #d8dee9', borderRadius: 6, padding: '6px 8px',
    font: '13px ui-sans-serif, system-ui, sans-serif', color: '#172033', boxSizing: 'border-box', outline: 'none' },
  empty: { margin: 'auto', color: '#6b7587', fontSize: 14 },
  error: { position: 'absolute', left: 16, right: 16, bottom: 14, padding: '9px 12px', color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0', borderRadius: 6, fontSize: 12 },
};
