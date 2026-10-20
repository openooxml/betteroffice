import { lazy, Suspense, useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx/editor';
import type { Proposal } from '@betteroffice/pptx';
import type { PptxCommandStore } from '../commands/types';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { commandReason } from '../commands/evaluate';
import { useTranslation } from '../i18n';
import { PeerlessPresentationOverlay } from './PeerlessPresentationOverlay';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { peerlessFrameSource } from './peerlessFrameSource';
import { commitComposition, PeerlessCommandRefusedError } from './peerlessViewStore';

const ProposalChangeDetails = lazy(() => import('../components/ProposalDetails').then((module) => ({ default: module.ProposalChangeDetails })));

function changesKey(proposal: Proposal): string {
  return JSON.stringify(proposal.changes, (_key, value: unknown) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b))) : value);
}

export interface PeerlessReviewState { open: boolean; selectedId: string | null; diff: boolean; presenting: boolean }

function Diff({ run, id, onPreview, onError }: { run: PeerlessPresentation; id: string; onPreview(proposal: Proposal | null): void; onError(error: Error): void }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const reportError = useRef(onError);
  reportError.current = onError;
  const visual = useSyncExternalStore(run.subscribeVisual, run.getVisualSnapshot, run.getVisualSnapshot);
  useEffect(() => {
    let cancelled = false;
    onPreview(null);
    const index = run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId);
    void (async () => {
      const source = await peerlessFrameSource(run, index);
      try {
        const diff = await run.client.access!('layoutProposalDiffSlide', id, index);
        run.require();
        const scale = visual?.fit.scale ?? 1;
        const buffer = document.createElement('canvas');
        const context = buffer.getContext('2d');
        if (!context) throw new Error('Canvas is unavailable');
        sizeCanvasForSlide(buffer, diff.frame, 1, scale);
        await paintSlide(context, diff.frame, 1, scale, { resolveImage: source.resolveImage });
        if (!cancelled && canvas.current && run.current) {
          sizeCanvasForSlide(canvas.current, diff.frame, 1, scale);
          canvas.current.getContext('2d')?.drawImage(buffer, 0, 0);
          onPreview(diff.proposal);
        }
      } finally { source.release(); }
    })().catch((error: unknown) => { if (!cancelled && run.current) reportError.current(error instanceof Error ? error : new Error(String(error))); });
    return () => { cancelled = true; };
  }, [run, id, visual, onPreview]);
  return <canvas ref={canvas} data-testid="pptx-proposal-diff" style={{ position: 'absolute', left: 0, top: 0 }} />;
}

export function PeerlessReview({ run, state, proposals, commands, onError, exit }: {
  run: PeerlessPresentation; state: PeerlessReviewState; proposals: Proposal[]; commands: PptxCommandStore;
  onError(error: Error): void; exit(): void;
}) {
  const { t } = useTranslation();
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [reviewed, setReviewed] = useState<Proposal | null>(null);
  const accepting = useRef(false);
  const [previewGeneration, setPreviewGeneration] = useState(0);
  const selected = useRef(state.selectedId);
  selected.current = state.diff ? state.selectedId : null;
  const onPreview = useCallback((proposal: Proposal | null) => setReviewed(proposal), [run]);
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
  const preview = async (proposal: Proposal) => {
    setPreviewError(null);
    setReviewed(null);
    setPreviewGeneration((generation) => generation + 1);
    try {
      const composition = run.store.getSnapshot().composition;
      if (composition) {
        const ack = await commitComposition(run.store, composition.text, performance.now());
        const refused = ack?.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
        if (refused) throw new PeerlessCommandRefusedError('compositionCommit', refused);
        await run.store.fence();
      }
      run.require();
      const slideId = proposal.changes[0]?.slideId;
      if (slideId && slideId !== run.client.state.activeSlideId) {
        const ack = await run.store.dispatch({ kind: 'activateSlide', payload: { slideId } });
        const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
        if (refused) throw new PeerlessCommandRefusedError('activateSlide', refused);
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
    {state.open && <aside data-testid="pptx-proposals-panel" aria-label={t('proposals.title')} style={{ width: 280, overflowY: 'auto' }}>
      <strong>{t('proposals.title')} <span>{proposals.length}</span></strong>
      {previewError && <div role="alert">{previewError}</div>}
      <button type="button" aria-label={t('proposals.close')} onClick={() => { void commands.execute('proposalsPanel', { open: false }); }}>×</button>
      {proposals.map((proposal) => <section key={proposal.id} data-testid="pptx-proposal" data-proposal-id={proposal.id}>
        <strong>{proposal.agentId}</strong>
        {proposal.note && <p>{proposal.note}</p>}
        <Suspense fallback={null}>
          {(state.diff && state.selectedId === proposal.id && reviewed?.id === proposal.id ? reviewed : proposal).changes.map((change, index) => <ProposalChangeDetails key={index} change={change} t={t} testId={change.shapeId === null ? 'pptx-proposal-notes-diff' : undefined} />)}
          {proposal.staleTargets.length > 0 && state.diff && state.selectedId === proposal.id && reviewed?.id === proposal.id && <>
            <p>{t('proposals.stalePreview')}</p>
            <button type="button" data-testid="pptx-proposal-force" onClick={() => { void acceptPreview(); }}>{t('proposals.force')}</button>
          </>}
        </Suspense>
        {proposal.staleTargets.length > 0 && <p role="alert" data-testid="pptx-proposal-stale">{t('proposals.stale')}</p>}
        <button type="button" data-testid="pptx-proposal-preview" onClick={() => { void preview(proposal); }}>{t('proposals.preview')}</button>
        <button type="button" data-testid="pptx-proposal-accept" disabled={proposal.staleTargets.length > 0} onClick={() => { void commands.execute('proposalAccept', { proposalId: proposal.id }); }}>{t('proposals.accept')}</button>
        <button type="button" data-testid="pptx-proposal-reject" onClick={() => { void commands.execute('proposalReject', { proposalId: proposal.id }); }}>{t('proposals.reject')}</button>
      </section>)}
    </aside>}
    {state.diff && state.selectedId && <Diff key={previewGeneration} run={run} id={state.selectedId} onPreview={onPreview} onError={onError} />}
    {state.presenting && <PeerlessPresentationOverlay frameSource={(index) => peerlessFrameSource(run, index)}
      slideCount={run.client.state.slides.length} startIndex={run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId)}
      label={t('presentation.label')} counterLabel={(current, total) => t('presentation.slideCounter', { current, total })}
      previousLabel={t('presentation.previousSlide')} nextLabel={t('presentation.nextSlide')} exitLabel={t('presentation.exit')}
      onExit={exit} onError={(error) => onError(error instanceof Error ? error : new Error(String(error)))} />}
  </>;
}
