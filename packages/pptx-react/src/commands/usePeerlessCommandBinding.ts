import { useLayoutEffect, useMemo } from 'react';
import { flushSync } from 'react-dom';
import type { Translations } from '@betteroffice/pptx-i18n';
import type { Proposal } from '@betteroffice/pptx';
import type { PresentationCommandAck } from '@betteroffice/pptx/worker';
import type { PeerlessReviewState } from '../worker/PeerlessReview';
import { invokePeerless } from '../worker/peerlessObjectCommands';
import { useTranslation } from '../i18n';
import type { PptxEditorTool } from '../components/toolbarTypes';
import type { PeerlessPresentation } from '../worker/usePeerlessPresentation';
import { formatPeerlessText, peerlessTextContext } from '../worker/peerlessTextCommands';
import { commitComposition, hitConfirmedGeometry } from '../worker/peerlessViewStore';
import { peerlessHistory } from '../worker/peerlessHistory';
import { peerlessObjectCommand } from '../worker/peerlessObjectCommands';
import { peerlessInsertSlide } from '../worker/peerlessSlideCommands';
import { createPptxCommandController, PptxCommandAdmissionError, type PptxCommandBinding, type PptxCommandOrigin } from './createPptxCommandStore';
import { commandReason, DEFAULT_FONT_FAMILIES, DEFAULT_FONT_SIZES, type PptxCommandEnvironment } from './evaluate';
import type { PptxCommandArgs, PptxCommandFailureCode, PptxCommandId, PptxCommandResult } from './types';

interface Origin extends PptxCommandOrigin { run: PeerlessPresentation; slideId: string | null; target: string | null; id: PptxCommandId }

const supported = new Set<PptxCommandId>(['bold', 'italic', 'underline', 'fontFamily', 'fontSize', 'fontSizeStep',
  'textColor', 'alignment', 'tool', 'insertImage', 'save', 'undo', 'redo', 'shapeFill', 'shapeStrokeColor', 'shapeStrokeWidth', 'shapeAdjustment', 'zOrder', 'insertSlide', 'zoom', 'proposalsPanel', 'proposalSelect', 'proposalDiff', 'proposalAccept', 'proposalReject', 'slideshow', 'exportPng']);

function target(run: PeerlessPresentation): string | null {
  const selected = run.store.getSnapshot().selection;
  return selected ? JSON.stringify([selected.slideId, selected.shapeId, selected.storyId, selected.kind, selected.anchor, selected.focus]) : null;
}

function receipt(ack: PresentationCommandAck | null, env: PptxCommandEnvironment): PptxCommandResult {
  if (!ack) return { ok: false, failure: commandReason('unsupported-selection', env) };
  const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
  if (refused) {
    const code: PptxCommandFailureCode = refused.reason === 'read-only' ? 'read-only'
      : refused.reason === 'bad-target' || refused.reason === 'stale-basis' ? 'target-changed'
      : refused.reason === 'invalid-payload' || refused.reason === 'invalid-offset' ? 'invalid-arguments' : 'input-failed';
    return { ok: false, failure: commandReason(code, env) };
  }
  return { ok: true, status: ack.outcomes.some((outcome) => outcome.status === 'applied') ? 'executed' : 'noop' };
}

export function usePeerlessCommandBinding(run: PeerlessPresentation | null, {
  tool, setTool, openPicture, save, focus, i18n, zoom, setZoom, review, proposals, setReview, exportPng,
}: { tool: PptxEditorTool; setTool(tool: PptxEditorTool): void; openPicture(): void; save(): Promise<void>;
  focus(): void; i18n?: Translations; zoom: number | 'fit'; setZoom(value: number | 'fit'): void; review: PeerlessReviewState; proposals: Proposal[]; setReview(next: PeerlessReviewState): void; exportPng(): Promise<void> }) {
  const { t } = useTranslation();
  const controller = useMemo(() => createPptxCommandController(), [run]);
  const binding = useMemo(() => {
    const environment = (executing: boolean): PptxCommandEnvironment => {
      const view = run?.store.getSnapshot();
      let text = view ? peerlessTextContext(view) : null;
      const pending = view?.pending[view.pending.length - 1]?.intent;
      if (!text && executing && run && pending?.kind === 'selectAtPoint' &&
        hitConfirmedGeometry(run.store, pending.payload)?.kind === 'text') {
        text = { kind: 'caret', bold: 'mixed', italic: 'mixed', underline: 'mixed',
          fontFamily: null, fontSize: null, color: null, alignment: null };
      }
      const slideId = view?.confirmed?.frame.slideId;
      const slide = run?.client.state.slides.find((slide) => slide.id === slideId);
      return { status: !run?.current ? 'unavailable' : 'ready',
        readOnly: view?.editState.readOnly ?? true, reviewing: review.diff,
        pendingInput: Boolean((view?.pending.length || view?.waitingInputCount) && (!view?.editState.canUndo ||
          view?.pending.some((record) => record.intent.kind === 'undo' || record.intent.kind === 'redo'))),
        canUndo: view?.editState.canUndo ?? false, canRedo: view?.editState.canRedo ?? false,
        slide: slide ? { id: slide.id, layoutPartPath: slide.layoutPartPath } : null, layouts: [...new Set(run?.client.state.slides.flatMap((slide) => slide.layoutPartPath ? [slide.layoutPartPath] : []))].map((layoutPartPath, index) => ({ value: layoutPartPath, label: t('toolbar.layoutOption', { number: index + 1 }), args: { layoutPartPath } })),
        text, shape: view?.editState.object ?? null, tool, zoom,
        fontFamilies: DEFAULT_FONT_FAMILIES, fontSizes: DEFAULT_FONT_SIZES, proposals: { open: review.open, pending: proposals.map((proposal) => ({ id: proposal.id, label: proposal.note ?? proposal.agentId, stale: proposal.staleTargets.length > 0 })), canvas: { available: proposals.filter((proposal) => proposal.changes.some((change) => change.slideId === slideId)).map((proposal) => ({ id: proposal.id, label: proposal.note ?? proposal.agentId, stale: proposal.staleTargets.length > 0 })), selectedId: review.selectedId, diff: review.diff } },
        hostDisabled: (id) => !supported.has(id), translate: t };
    };
    const value: PptxCommandBinding = {
      environment,
      ordered: (id) => id !== 'tool' && id !== 'insertImage',
      admit(operation) {
        try { run?.require(); return Promise.resolve(operation()); }
        catch { return Promise.reject(new PptxCommandAdmissionError('input-failed')); }
      },
      capture(id) {
        return run?.current ? { generation: 0, run, id, slideId: run.store.getSnapshot().confirmed?.frame.slideId ?? null,
          target: target(run) } as Origin : null;
      },
      resume(captured) {
        const origin = captured as Origin;
        if (!run?.current || origin.run !== run) return 'document-replaced';
        if (origin.slideId !== (run.store.getSnapshot().confirmed?.frame.slideId ?? null)) return 'target-changed';
        if (!['undo', 'redo', 'save', 'tool', 'insertImage'].includes(origin.id) && origin.target !== target(run)) return 'target-changed';
        return null;
      },
      async perform<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K], env: PptxCommandEnvironment) {
        if (!run?.current) return { ok: false, failure: commandReason('editor-unavailable', env) };
        const text = (action: Parameters<typeof formatPeerlessText>[1]) => formatPeerlessText(run.store, action);
        let operation: Promise<PresentationCommandAck | null> | null = null;
        switch (id) {
          case 'zoom': setZoom((args as PptxCommandArgs['zoom']).scale); return { ok: true, status: 'executed' };
          case 'proposalsPanel': setReview({ ...review, open: args === null ? !review.open : (args as PptxCommandArgs['proposalsPanel'])!.open }); return { ok: true, status: 'opened' };
          case 'proposalSelect': flushSync(() => setReview({ ...review, selectedId: (args as PptxCommandArgs['proposalSelect']).proposalId })); return { ok: true, status: 'executed' };
          case 'proposalDiff': {
            const enabled = (args as PptxCommandArgs['proposalDiff']).enabled;
            const composition = run.store.getSnapshot().composition;
            if (enabled && composition) {
              const ack = await commitComposition(run.store, composition.text, performance.now());
              const result = receipt(ack, env);
              if (!result.ok) return result;
              await run.store.fence();
              run.require();
            }
            flushSync(() => setReview({ ...review, diff: enabled }));
            return { ok: true, status: 'executed' };
          }
          case 'slideshow': await run.store.fence(); setReview({ ...review, presenting: true }); return { ok: true, status: 'opened' };
          case 'exportPng': await exportPng(); return { ok: true, status: 'executed' };
          case 'proposalAccept': operation = invokePeerless(run, { method: 'acceptProposal', args: [(args as PptxCommandArgs['proposalAccept']).proposalId, { force: (args as PptxCommandArgs['proposalAccept']).force }] }); setReview({ ...review, diff: false }); break;
          case 'proposalReject': operation = invokePeerless(run, { method: 'rejectProposal', args: [(args as PptxCommandArgs['proposalReject']).proposalId] }); setReview({ ...review, diff: false }); break;
          case 'insertSlide': operation = peerlessInsertSlide(run, (args as PptxCommandArgs['insertSlide']).layoutPartPath); break;
          case 'shapeFill': operation = peerlessObjectCommand(run, { kind: 'fill', color: (args as PptxCommandArgs['shapeFill']).color }); break;
          case 'shapeStrokeColor': operation = peerlessObjectCommand(run, { kind: 'strokeColor', color: (args as PptxCommandArgs['shapeStrokeColor']).color }); break;
          case 'shapeStrokeWidth': operation = peerlessObjectCommand(run, { kind: 'strokeWidth', points: (args as PptxCommandArgs['shapeStrokeWidth']).points }); break;
          case 'shapeAdjustment': operation = peerlessObjectCommand(run, { kind: 'adjust', ...(args as PptxCommandArgs['shapeAdjustment']) }); break;
          case 'zOrder': operation = peerlessObjectCommand(run, { kind: 'zOrder', value: (args as PptxCommandArgs['zOrder']).value }); break;
          case 'tool': setTool((args as PptxCommandArgs['tool']).value); return { ok: true, status: 'executed' };
          case 'insertImage': openPicture(); return { ok: true, status: 'opened' };
          case 'save': await save(); return { ok: true, status: 'executed' };
          case 'undo': case 'redo': operation = peerlessHistory(run.store, id); break;
          case 'bold': case 'italic': case 'underline': operation = text(id); break;
          case 'fontFamily': operation = text({ type: 'fontFamily', value: (args as PptxCommandArgs['fontFamily']).family }); break;
          case 'fontSize': operation = text({ type: 'fontSize', value: (args as PptxCommandArgs['fontSize']).points }); break;
          case 'fontSizeStep': operation = text({ type: 'fontSizeStep', direction: (args as PptxCommandArgs['fontSizeStep']).direction, sizes: env.fontSizes }); break;
          case 'textColor': operation = text({ type: 'textColor', value: (args as PptxCommandArgs['textColor']).color }); break;
          case 'alignment': operation = text({ type: 'align', value: (args as PptxCommandArgs['alignment']).value }); break;
          default: return { ok: false, failure: commandReason('unsupported-command', env) };
        }
        const ack = await operation;
        run.require();
        const result = receipt(ack, env);
        if (result.ok) focus();
        return result;
      },
      chrome: () => ({ i18n }), focusEditor: focus,
    };
    return value;
  }, [run, tool, setTool, openPicture, save, focus, i18n, zoom, setZoom, review, proposals, setReview, exportPng, t]);
  useLayoutEffect(() => {
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [controller, binding]);
  useLayoutEffect(() => {
    if (!run) return;
    let previous = JSON.stringify(binding.environment(false));
    const refresh = () => {
      const next = JSON.stringify(binding.environment(false));
      if (previous === next) return;
      previous = next;
      controller.refresh();
    };
    const releases = [run.store.subscribeEditState(refresh), run.store.subscribePresentation(refresh)];
    return () => { for (const release of releases) release(); };
  }, [controller, run, binding]);
  return controller.store;
}
