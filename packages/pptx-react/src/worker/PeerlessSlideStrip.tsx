import { memo, useCallback, useEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import type { CSSProperties } from 'react';
import type { SlideFit } from '../fitSlideScale';
import { useTranslation } from '../i18n';
import { slideActionStyles } from '../components/editorChromeStyles';
import type { PeerlessConfirmedFrame, PeerlessViewSnapshot } from './peerlessViewStore';
import { peerlessDeleteSlide, peerlessInsertSlide, peerlessMoveSlide } from './peerlessSlideCommands';
import { activatePeerlessSlide } from './peerlessNavigation';
import { peerlessPainterFactories, type PeerlessPresentation } from './usePeerlessPresentation';

const Preview = memo(function Preview({ frame, run, fit, previewFit, current }: {
  frame: PeerlessConfirmedFrame; run: PeerlessPresentation; fit: SlideFit;
  previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
}) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const latest = useRef({ previewFit, current });
  latest.current = { previewFit, current };
  const resume = useRef<((current: () => boolean) => Promise<void>) | null>(null);
  useEffect(() => {
    const valid = () => {
      const token = latest.current.previewFit(frame, run);
      return Boolean(token && token.scale === fit.scale && token.dpr === fit.dpr && latest.current.current(token, run));
    };
    if (!canvas.current) return;
    const painter = peerlessPainterFactories.frame(canvas.current, { ...fit, images: run.images, current: valid });
    const view: PeerlessViewSnapshot = { confirmed: frame, selection: null, focused: false, caret: null,
      selectionRects: [], composition: null, echo: null, activeLine: null, activeContextBytes: 0,
      pending: [], waitingInputCount: 0, appliedThroughSeq: frame.frame.coversThroughSeq,
      presentedThroughSeq: frame.frame.coversThroughSeq,
      editState: { readOnly: true, canUndo: false, canRedo: false,
        toolbar: { target: null, storedMarks: {}, formatting: null, alignment: null,
          hyperlink: null, table: null, canEditText: false } },
      thumbnails: new Map(), dirtyThumbnailIds: new Set<string>() };
    painter.paintFrame(frame, view);
    void painter.whenPainted().catch(() => {});
    resume.current = async (current) => {
      if (!await painter.whenPainted() && current() && valid()) {
        painter.paintFrame(frame, view);
        await painter.whenPainted();
      }
    };
    return () => { resume.current = null; painter.dispose(); };
  }, [frame, run, fit.scale, fit.dpr]);
  useEffect(() => {
    let alive = true;
    void resume.current?.(() => alive).catch(() => {});
    return () => { alive = false; };
  }, [frame, run, fit.scale, fit.dpr, fit.generation]);
  return <canvas ref={canvas} style={{ width: 128, display: 'block' }} aria-hidden="true" />;
});

const Slide = memo(function Slide({ id, name, index, active, disabled, frame, run, previewFit, current, onNavigate }: {
  id: string; name: string | null; index: number; active: boolean; disabled: boolean;
  frame: PeerlessConfirmedFrame | null | undefined; run: PeerlessPresentation;
  previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
  onNavigate(): void;
}) {
  const { t } = useTranslation();
  const fit = useMemo(() => frame ? previewFit(frame, run) : null, [frame, run, previewFit]);
  return <button type="button" disabled={disabled}
    aria-current={active ? 'page' : undefined} style={active ? styles.activeSlide : styles.slide}
    draggable={!run.store.getSnapshot().editState.readOnly}
    onDragStart={(event) => event.dataTransfer.setData('text/plain', id)}
    onDragOver={(event) => event.preventDefault()}
    onDrop={(event) => {
      event.preventDefault();
      const source = event.dataTransfer.getData('text/plain');
      if (run.client.state.slides.some((slide) => slide.id === source)) void peerlessMoveSlide(run, source, index).catch(() => {});
    }}
    onKeyDown={(event) => {
      if (!event.altKey || !['ArrowUp', 'ArrowDown'].includes(event.key)) return;
      event.preventDefault();
      const next = index + (event.key === 'ArrowUp' ? -1 : 1);
      if (next >= 0 && next < run.client.state.slides.length) void peerlessMoveSlide(run, id, next).catch(() => {});
    }}
    onClick={() => {
      void activatePeerlessSlide(run, id).catch(() => {});
      onNavigate();
    }}>
    <span style={styles.number}>{index + 1}</span>
    <span style={styles.preview}>
      {frame && fit ? <Preview frame={frame} run={run} fit={fit} previewFit={previewFit} current={current} /> : <span style={styles.title}>{name ?? t('slides.fallbackTitle', { number: index + 1 })}</span>}
    </span>
  </button>;
});

function SlideActions({ run, active, disabled }: { run: PeerlessPresentation; active?: string; disabled: boolean }) {
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => run.store.subscribeEditState(listener), [run]);
  const snapshot = useCallback(() => run.store.getSnapshot().editState.readOnly, [run]);
  const readOnly = useSyncExternalStore(subscribe, snapshot, snapshot);
  return readOnly ? null : <div style={slideActionStyles.row}>
    <button type="button" disabled={disabled} style={actionButton(disabled)} onClick={() => { void peerlessInsertSlide(run).catch(() => {}); }}>{t('toolbar.addSlide')}</button>
    <button type="button" disabled={disabled || !active} style={actionButton(disabled || !active)} onClick={() => { if (active) void peerlessDeleteSlide(run, active).catch(() => {}); }}>{t('toolbar.deleteSlide')}</button>
  </div>;
}

export const PeerlessSlideStrip = memo(function PeerlessSlideStrip({ run, previewFit, current, onNavigate }: {
  run: PeerlessPresentation; previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
  onNavigate(): void;
}) {
  const { store, client } = run;
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => {
    const offStructure = store.subscribeStructure(listener);
    const offPresentation = store.subscribePresentation(listener);
    return () => { offStructure(); offPresentation(); };
  }, [store]);
  const snapshot = useMemo(() => {
    let previous: { active: string | undefined; disabled: boolean; slides: { id: string; name: string | null;
      frame: PeerlessConfirmedFrame | null | undefined }[] } | undefined;
    return () => {
      const view = store.getSnapshot();
      const active = view.confirmed?.frame.slideId;
      const disabled = !run.current || Boolean(run.failure);
      const slides = client.state.slides.map((slide) => ({ id: slide.id, name: slide.name,
        frame: slide.id === active ? view.confirmed : view.thumbnails.get(slide.id) }));
      if (previous && previous.active === active && previous.disabled === disabled && previous.slides.length === slides.length &&
        slides.every((slide, index) => slide.id === previous!.slides[index]!.id &&
          slide.name === previous!.slides[index]!.name && slide.frame === previous!.slides[index]!.frame)) return previous;
      previous = { active, disabled, slides };
      return previous;
    };
  }, [store, client, run]);
  const view = useSyncExternalStore(subscribe, snapshot, snapshot);
  return <div style={styles.strip}><aside style={{ overflowY: 'auto', flex: 1 }} aria-label={t('slides.panelLabel')}>
    {view.slides.map((slide, index) => <Slide key={slide.id} {...slide} index={index} active={slide.id === view.active}
      disabled={view.disabled} run={run} previewFit={previewFit} current={current} onNavigate={onNavigate} />)}
  </aside><SlideActions run={run} active={view.active} disabled={view.disabled} /></div>;
});

function actionButton(disabled: boolean): CSSProperties {
  return disabled ? { ...slideActionStyles.button, opacity: 0.5, cursor: 'default' } : slideActionStyles.button;
}

const styles: Record<string, CSSProperties> = {
  strip: { width: 184, display: 'flex', flexDirection: 'column', padding: '14px 10px', background: '#eef1f5',
    borderRight: '1px solid #d8dee9', boxSizing: 'border-box' },
  slide: { display: 'flex', alignItems: 'flex-start', gap: 6, width: '100%', marginBottom: 12, padding: 4,
    border: '2px solid transparent', borderRadius: 5, background: 'transparent', cursor: 'pointer' },
  activeSlide: { display: 'flex', alignItems: 'flex-start', gap: 6, width: '100%', marginBottom: 12, padding: 4,
    border: '2px solid #325ee6', borderRadius: 5, background: 'transparent', cursor: 'pointer' },
  number: { width: 18, flex: '0 0 auto', paddingTop: 3, fontSize: 11, color: '#647087', textAlign: 'right' },
  preview: { display: 'flex', alignItems: 'center', justifyContent: 'center', minWidth: 0, background: '#ffffff',
    boxShadow: '0 1px 4px rgba(20, 31, 50, 0.16)' },
  title: { padding: 8, fontSize: 9, lineHeight: 1.25, color: '#39445a', textAlign: 'center' },
};
