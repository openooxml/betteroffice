import { lazy, Suspense, useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { RefObject } from 'react';
import { createPortal } from 'react-dom';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type { Proposal, ProposalDiffSlide } from '@betteroffice/pptx';
import type { PptxCommandStore } from '../commands/types';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { commandReason } from '../commands/evaluate';
import { useTranslation } from '../i18n';
import { proposalsPanelStyles as styles } from '../components/proposalsPanelStyles';
import { PeerlessPresentationOverlay } from './PeerlessPresentationOverlay';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { peerlessFrameSource, peerlessProposalFrameSource } from './peerlessFrameSource';
import { commitComposition, PptxCommandRefusedError } from './peerlessViewStore';
import { ProposalShapeOutlines } from '../components/ProposalCanvas';
import { proposalImages, proposalPaintChanges } from '../components/proposalImages';

const location = { ...styles.link, cursor: 'default' };

const ProposalChangeDetails = lazy(() => import('../components/ProposalDetails').then((module) => ({ default: module.ProposalChangeDetails })));

function changesKey(proposal: Proposal): string {
  return JSON.stringify(proposal.changes, (_key, value: unknown) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b))) : value);
}

export interface PeerlessReviewState { open: boolean; selectedId: string | null; diff: boolean; presenting: boolean; slideId?: string | null; slideVersion?: string }

function Diff({ run, id, slideId, onPreview, onError, onTarget }: { run: PeerlessPresentation; id: string; slideId?: string | null; onPreview(proposal: Proposal | null): void; onError(error: Error): void; onTarget(): void }) {
  const { t } = useTranslation();
  const decodeImageError = t('errors.decodeSlideImage');
  const canvas = useRef<HTMLCanvasElement>(null);
  const [painted, setPainted] = useState<{ diff: ProposalDiffSlide; scale: number } | null>(null);
  const reportError = useRef(onError);
  reportError.current = onError;
  const visual = useSyncExternalStore(run.subscribeVisual, run.getVisualSnapshot, run.getVisualSnapshot);
  useEffect(() => {
    let cancelled = false;
    onPreview(null);
    setPainted(null);
    void (async () => {
      const source = await peerlessProposalFrameSource(run, id, slideId);
      try {
        const diff = source;
        run.require();
        const scale = visual?.fit.scale ?? 1;
        const buffer = document.createElement('canvas');
        const context = buffer.getContext('2d');
        if (!context) throw new Error('Canvas is unavailable');
        sizeCanvasForSlide(buffer, diff.frame, 1, scale);
        const images = proposalImages(diff, source.resolveImage, decodeImageError);
        try {
          await paintSlide(context, diff.frame, 1, scale, { resolveImage: images.resolveImage, textChanges: diff.textChanges, shapeChanges: proposalPaintChanges(diff) });
        } finally { images.release(); }
        if (!cancelled && canvas.current && run.current) {
          sizeCanvasForSlide(canvas.current, diff.frame, 1, scale);
          canvas.current.getContext('2d')?.drawImage(buffer, 0, 0);
          setPainted({ diff, scale });
          onPreview(diff.proposal);
        }
      } finally { source.release(); }
    })().catch((error: unknown) => { if (!cancelled && run.current) reportError.current(error instanceof Error ? error : new Error(String(error))); });
    return () => { cancelled = true; };
  }, [run, id, slideId, visual, onPreview, decodeImageError]);
  return <>
    <canvas ref={canvas} data-testid="pptx-proposal-diff" data-slide-id={slideId ?? run.client.state.activeSlideId} style={{ position: 'absolute', left: 0, top: 0 }} />
    {painted && <ProposalShapeOutlines diff={painted.diff} current={painted.diff.snapshot} frame={{ ...painted.diff.frame, primitives: [] }} scale={painted.scale} onTarget={onTarget} />}
  </>;
}

export function PeerlessReview({ run, diffHost, state, proposals, commands, onError, exit }: {
  run: PeerlessPresentation; diffHost: RefObject<HTMLElement | null>; state: PeerlessReviewState; proposals: Proposal[]; commands: PptxCommandStore;
  onError(error: Error): void; exit(): void;
}) {
  const { t, lang } = useTranslation();
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [reviewed, setReviewed] = useState<Proposal | null>(null);
  const accepting = useRef(false);
  const [previewGeneration, setPreviewGeneration] = useState(0);
  const selected = useRef(state.selectedId);
  selected.current = state.diff ? state.selectedId : null;
  const onPreview = useCallback((proposal: Proposal | null) => setReviewed(proposal), [run]);
  const frameSource = useCallback((index: number) => peerlessFrameSource(run, index), [run]);
  const subscribeRevision = useCallback((listener: () => void) => {
    const offStructure = run.store.subscribeStructure(listener);
    const offPresentation = run.store.subscribePresentation(listener);
    return () => { offStructure(); offPresentation(); };
  }, [run]);
  const revision = useCallback((index: number) => {
    const id = run.client.state.slides[index]?.id;
    if (!id) return null;
    const view = run.store.getSnapshot();
    const frame = id === view.confirmed?.frame.slideId ? view.confirmed : view.thumbnails.get(id);
    return `${id}:${frame?.frame.slideRevision ?? 'stale'}`;
  }, [run]);
  const acceptPreview = async () => {
    if (!reviewed || selected.current !== reviewed.id || accepting.current) return;
    const target = reviewed;
    const accept = pptxCommandController(commands)?.defer('proposalAccept', { proposalId: target.id, force: true });
    if (!accept) return;
    accepting.current = true;
    try {
      await accept.complete(async (perform) => {
        const latest = await run.client.access!('previewProposal', target.id);
        run.require();
        if (selected.current !== target.id) return { ok: false, failure: commandReason('target-changed', { translate: t }) };
        if (changesKey(latest.proposal) !== changesKey(target)) {
          setReviewed(latest.proposal);
          setPreviewError(t('proposals.changedAgain'));
          return { ok: false, failure: commandReason('target-changed', { translate: t }) };
        }
        return perform();
      });
    } catch (error: unknown) {
      if (run.current) {
        setReviewed(null);
        const failure = error instanceof Error ? error : new Error(String(error));
        setPreviewError(failure.message);
        onError(failure);
      }
    } finally { accepting.current = false; }
  };
  const target = (change: Proposal['changes'][number]) => {
    const slide = run.client.state.slides.findIndex((slide) => slide.id === change.slideId);
    return slide < 0 && change.kind === 'inserted' ? t('proposals.target', { slide: (change.afterIndex ?? 0) + 1, name: t('proposals.insertedSlide') })
      : slide < 0 ? t('proposals.targetMissing')
      : t('proposals.target', { slide: slide + 1, name: change.after?.name || change.before?.name || t('notes.panelLabel') });
  };
  const preview = async (proposal: Proposal) => {
    setPreviewError(null);
    setReviewed(null);
    setPreviewGeneration((generation) => generation + 1);
    try {
      const composition = run.store.getSnapshot().composition;
      if (composition) {
        const ack = await commitComposition(run.store, composition.text, performance.now());
        const refused = ack?.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
        if (refused) throw new PptxCommandRefusedError('compositionCommit', refused);
        await run.store.fence();
      }
      run.require();
      const slideId = proposal.changes[0]?.slideId;
      if (slideId && run.client.state.slides.some(slide => slide.id === slideId) && slideId !== run.client.state.activeSlideId) {
        const ack = await run.store.dispatch({ kind: 'activateSlide', payload: { slideId } });
        const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
        if (refused) throw new PptxCommandRefusedError('activateSlide', refused);
        await run.ready;
      }
      const selected = await commands.execute('proposalSelect', { proposalId: proposal.id });
      if (!selected.ok) throw new Error(selected.failure.message);
      const diff = await commands.execute('proposalDiff', { enabled: true });
      if (!diff.ok) throw new Error(diff.failure.message);
    } catch (error: unknown) {
      if (run.current) {
        const failure = error instanceof Error ? error : new Error(String(error));
        setPreviewError(failure.message);
        onError(failure);
      }
    }
  };
  return <>
    {state.open && <aside data-testid="pptx-proposals-panel" aria-label={t('proposals.title')} style={styles.panel}>
      <header style={styles.header}>
        <strong>{t('proposals.title')} <span style={styles.count}>{proposals.length}</span></strong>
        <button type="button" aria-label={t('proposals.close')} style={styles.button} onClick={() => { void commands.execute('proposalsPanel', { open: false }); }}>×</button>
      </header>
      {previewError && <div role="alert" style={styles.warning}>{previewError}</div>}
      {proposals.length === 0 && <p style={styles.muted}>{t('proposals.empty')}</p>}
      {proposals.map((proposal) => <section key={proposal.id} data-testid="pptx-proposal" data-proposal-id={proposal.id} style={styles.card}>
        <strong style={styles.author}>{proposal.agentId}</strong>
        {proposal.note && <p style={styles.note}>{proposal.note}</p>}
        <Suspense fallback={null}>
          {(state.diff && state.selectedId === proposal.id && reviewed?.id === proposal.id ? reviewed : proposal).changes.map((change, index) => <div key={index} style={styles.change}>
            <span style={location}>{target(change)}</span>
            <ProposalChangeDetails change={change} t={t} lang={lang} testId={change.shapeId === null ? 'pptx-proposal-notes-diff' : undefined} />
          </div>)}
          {proposal.staleTargets.length > 0 && state.diff && state.selectedId === proposal.id && reviewed?.id === proposal.id && <>
            <p style={styles.warning}>{t('proposals.stalePreview')}</p>
            <button type="button" data-testid="pptx-proposal-force" style={styles.primary} onClick={() => { void acceptPreview(); }}>{t('proposals.force')}</button>
          </>}
        </Suspense>
        {proposal.staleTargets.length > 0 && <p role="alert" data-testid="pptx-proposal-stale" style={styles.warning}>{t('proposals.stale')}</p>}
        <div style={styles.actions}>
          <button type="button" data-testid="pptx-proposal-preview" style={styles.button} onClick={() => { void preview(proposal); }}>{t('proposals.preview')}</button>
        </div>
        <div style={styles.actions}>
          <button type="button" data-testid="pptx-proposal-accept" disabled={proposal.staleTargets.length > 0}
            style={proposal.staleTargets.length > 0 ? { ...styles.primary, opacity: 0.5, cursor: 'default' } : styles.primary}
            onClick={() => { void commands.execute('proposalAccept', { proposalId: proposal.id }); }}>{t('proposals.accept')}</button>
          <button type="button" data-testid="pptx-proposal-reject" style={styles.button} onClick={() => { void commands.execute('proposalReject', { proposalId: proposal.id }); }}>{t('proposals.reject')}</button>
        </div>
      </section>)}
    </aside>}
    {state.diff && state.selectedId && diffHost.current && createPortal(<Diff key={previewGeneration} run={run} id={state.selectedId} slideId={state.slideVersion && state.slideVersion !== run.client.state.version ? null : state.slideId} onPreview={onPreview} onError={onError} onTarget={() => { void commands.execute('proposalsPanel', { open: true }); }} />, diffHost.current)}
    {state.presenting && <PeerlessPresentationOverlay frameSource={frameSource} subscribeRevision={subscribeRevision} revision={revision}
      slideCount={run.client.state.slides.length} startIndex={run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId)}
      label={t('presentation.label')} counterLabel={(current, total) => t('presentation.slideCounter', { current, total })}
      previousLabel={t('presentation.previousSlide')} nextLabel={t('presentation.nextSlide')} exitLabel={t('presentation.exit')}
      onExit={exit} onError={(error) => onError(error instanceof Error ? error : new Error(String(error)))} />}
  </>;
}
