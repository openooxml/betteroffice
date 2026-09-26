import { useEffect, useId, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  DeckSnapshot,
  PresentationHandle,
  Proposal,
  ProposalChange,
  ProposalPreview,
  SlideDisplayList,
} from '@betteroffice/pptx';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { commandReason } from '../commands/evaluate';
import { usePptxCommands, usePptxCommandState } from '../commands/hooks';
import { useTranslation } from '../i18n';
import { useDisabledDescription } from './ui/ToolbarPrimitives';

interface Props {
  handle: PresentationHandle;
  proposals: Proposal[];
  snapshot: DeckSnapshot;
  resolveImage: CanvasImageResolver;
  onNavigate: (slideId: string, shapeId: string | null, proposalId?: string) => void;
  /** Called after the panel closed itself, to restore focus. */
  onClose: () => void;
}

function ProposalAction({
  command,
  proposalId,
  label,
  style,
  testId,
}: {
  command: 'proposalAccept' | 'proposalReject';
  proposalId: string;
  label: string;
  style: CSSProperties;
  testId?: string;
}) {
  const store = usePptxCommands();
  const args = { proposalId };
  const state = usePptxCommandState(command, args);
  const described = useDisabledDescription(
    !state.enabled,
    state.enabled ? undefined : state.disabledReason.message
  );
  return (
    <>
      <button
        type="button"
        data-testid={testId}
        {...described.props}
        title={described.reason}
        style={{ ...style, opacity: state.enabled ? 1 : 0.45 }}
        onClick={() => {
          if (state.enabled) void store.execute(command, args);
        }}
      >
        {label}
      </button>
      {described.node}
    </>
  );
}

export function ProposalsPanel({
  handle,
  proposals,
  snapshot,
  resolveImage,
  onNavigate,
  onClose,
}: Props) {
  const { t } = useTranslation();
  const store = usePptxCommands();
  const titleId = useId();
  const [selected, setSelected] = useState<{
    id: string;
    index: number;
  } | null>(null);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  const [preview, setPreview] = useState<{
    data: ProposalPreview;
    before: SlideDisplayList;
    after: SlideDisplayList;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const previewButton = useRef<HTMLButtonElement | null>(null);

  useEffect(() => {
    if (!selected) return;
    if (!proposals.some((proposal) => proposal.id === selected.id)) {
      setSelected(null);
      return;
    }
    try {
      const data = handle.previewProposal(selected.id);
      const change = data.proposal.changes[selected.index];
      const index = data.snapshot.slides.findIndex(
        (slide) => slide.id === change?.slideId
      );
      if (index < 0) throw new Error(t('proposals.targetMissing'));
      const before = handle.layoutSlide(index);
      const after = handle.layoutProposalSlide(selected.id, index);
      setPreview({ data, before, after });
      setError(null);
    } catch (value) {
      setPreview(null);
      setError(value instanceof Error ? value.message : String(value));
    }
  }, [handle, proposals, snapshot, selected, t]);

  useEffect(() => {
    const element = dialog.current;
    if (selected && element && !element.open) element.showModal();
    if (!selected && element?.open) element.close();
  }, [selected]);

  const location = (change: ProposalChange) => {
    const index = snapshot.slides.findIndex(
      (slide) => slide.id === change.slideId
    );
    const label = change.before?.name || t('notes.panelLabel');
    return index < 0
      ? t('proposals.targetMissing')
      : t('proposals.target', { slide: index + 1, name: label });
  };

  const openPreview = (
    proposal: Proposal,
    index: number,
    button: HTMLButtonElement
  ) => {
    previewButton.current = button;
    const change = proposal.changes[index];
    onNavigate(change.slideId, change.shapeId, proposal.id);
    setPreview(null);
    setError(null);
    setSelected({ id: proposal.id, index });
  };

  /**
   * Whether `reviewed` still shows the proposal of `target`. Otherwise the dialog shows the latest
   * preview of `target`, unless the reviewer has selected something else since.
   */
  const stillReviewed = (
    target: { id: string; index: number },
    reviewed: ProposalPreview
  ) => {
    const shown = () =>
      selectedRef.current?.id === target.id &&
      selectedRef.current.index === target.index;
    try {
      const latest = handle.previewProposal(target.id);
      if (
        JSON.stringify(latest.proposal.changes) ===
        JSON.stringify(reviewed.proposal.changes)
      ) {
        return true;
      }
      if (!shown()) return false;
      const change = latest.proposal.changes[target.index];
      const index = latest.snapshot.slides.findIndex(
        (slide) => slide.id === change.slideId
      );
      setPreview({
        data: latest,
        before: handle.layoutSlide(index),
        after: handle.layoutProposalSlide(target.id, index),
      });
      setError(t('proposals.changedAgain'));
    } catch (value) {
      if (!shown()) return false;
      setPreview(null);
      setError(value instanceof Error ? value.message : String(value));
    }
    return false;
  };

  const acceptPreview = (force: boolean) => {
    if (!selected || !preview) return;
    const target = selected;
    const reviewed = preview.data;
    if (!stillReviewed(target, reviewed)) return;
    const accept = pptxCommandController(store)?.defer('proposalAccept', {
      proposalId: target.id,
      force,
    });
    void accept?.complete((perform) =>
      stillReviewed(target, reviewed)
        ? perform()
        : { ok: false, failure: commandReason('target-changed', { translate: t }) }
    );
  };

  return (
    <aside
      data-testid="pptx-proposals-panel"
      aria-label={t('proposals.title')}
      style={styles.panel}
    >
      <header style={styles.header}>
        <strong>
          {t('proposals.title')}{' '}
          <span style={styles.count}>{proposals.length}</span>
        </strong>
        <button
          type="button"
          onClick={() =>
            void store.execute('proposalsPanel', { open: false }).then(() => onClose())
          }
          aria-label={t('proposals.close')}
          style={styles.button}
        >
          ×
        </button>
      </header>
      {proposals.length === 0 && (
        <p style={styles.muted}>{t('proposals.empty')}</p>
      )}
      {proposals.map((proposal) => (
        <section
          key={proposal.id}
          data-testid="pptx-proposal"
          data-proposal-id={proposal.id}
          style={styles.card}
        >
          <strong style={styles.author}>{proposal.agentId}</strong>
          {proposal.note && <p style={styles.note}>{proposal.note}</p>}
          {proposal.changes.map((change, index) => (
            <div
              key={`${change.slideId}:${change.shapeId ?? 'notes'}`}
              style={styles.change}
            >
              <button
                type="button"
                onClick={() => onNavigate(change.slideId, change.shapeId, proposal.id)}
                style={styles.link}
              >
                {location(change)}
              </button>
              {change.oldText !== change.newText ? (
                <div style={styles.textDiff}>
                  <del style={styles.deleted}>
                    {change.oldText || t('proposals.blank')}
                  </del>
                  <ins style={styles.inserted}>
                    {change.newText || t('proposals.blank')}
                  </ins>
                </div>
              ) : null}
              {change.before && change.after && <ShapeChange change={change} />}
              <button
                type="button"
                data-testid="pptx-proposal-preview"
                style={styles.button}
                onClick={(event) =>
                  openPreview(proposal, index, event.currentTarget)
                }
              >
                {t('proposals.preview')}
              </button>
            </div>
          ))}
          {proposal.staleTargets.length > 0 && (
            <p
              role="alert"
              data-testid="pptx-proposal-stale"
              style={styles.warning}
            >
              {t('proposals.stale')}
            </p>
          )}
          <div style={styles.actions}>
            <ProposalAction
              command="proposalAccept"
              proposalId={proposal.id}
              label={t('proposals.accept')}
              style={styles.primary}
              testId="pptx-proposal-accept"
            />
            <ProposalAction
              command="proposalReject"
              proposalId={proposal.id}
              label={t('proposals.reject')}
              style={styles.button}
              testId="pptx-proposal-reject"
            />
          </div>
        </section>
      ))}
      <dialog
        ref={dialog}
        data-testid="pptx-proposal-preview-dialog"
        aria-labelledby={titleId}
        style={styles.dialog}
        onCancel={() => setSelected(null)}
        onClose={() => {
          setSelected(null);
          previewButton.current?.focus();
        }}
      >
        <header style={styles.header}>
          <strong id={titleId}>{t('proposals.previewTitle')}</strong>
          <button
            type="button"
            style={styles.button}
            onClick={() => setSelected(null)}
          >
            {t('proposals.close')}
          </button>
        </header>
        {error && (
          <p role="alert" style={styles.warning}>
            {error}
          </p>
        )}
        {preview && selected && (
          <>
            <p>
              <strong>{preview.data.proposal.agentId}</strong>
              {preview.data.proposal.note && (
                <> · {preview.data.proposal.note}</>
              )}
            </p>
            <label>
              {t('proposals.targetLabel')}{' '}
              <select
                aria-label={t('proposals.targetLabel')}
                value={selected.index}
                onChange={(event) => {
                  const index = Number(event.target.value);
                  const change = preview.data.proposal.changes[index];
                  onNavigate(change.slideId, change.shapeId, selected.id);
                  setSelected({ id: selected.id, index });
                }}
              >
                {preview.data.proposal.changes.map((change, index) => (
                  <option key={index} value={index}>
                    {location(change)}
                  </option>
                ))}
              </select>
            </label>
            {preview.data.proposal.changes[selected.index].oldText !==
              preview.data.proposal.changes[selected.index].newText && (
              <div style={styles.textDiff}>
                <del style={styles.deleted}>
                  {preview.data.proposal.changes[selected.index].oldText ||
                    t('proposals.blank')}
                </del>
                <ins style={styles.inserted}>
                  {preview.data.proposal.changes[selected.index].newText ||
                    t('proposals.blank')}
                </ins>
              </div>
            )}
            <div style={styles.previews}>
              <PreviewCanvas
                label={t('proposals.before')}
                frame={preview.before}
                resolveImage={resolveImage}
              />
              <PreviewCanvas
                label={t('proposals.after')}
                frame={preview.after}
                resolveImage={resolveImage}
              />
            </div>
            {preview.data.proposal.staleTargets.length > 0 && (
              <p role="alert" style={styles.warning}>
                {t('proposals.stalePreview')}
              </p>
            )}
            <div style={styles.actions}>
              {preview.data.proposal.staleTargets.length > 0 ? (
                <button
                  type="button"
                  data-testid="pptx-proposal-force"
                  style={styles.primary}
                  onClick={() => acceptPreview(true)}
                >
                  {t('proposals.force')}
                </button>
              ) : (
                <button
                  type="button"
                  style={styles.primary}
                  onClick={() => acceptPreview(false)}
                >
                  {t('proposals.accept')}
                </button>
              )}
              <ProposalAction
                command="proposalReject"
                proposalId={selected.id}
                label={t('proposals.reject')}
                style={styles.button}
              />
            </div>
          </>
        )}
      </dialog>
    </aside>
  );
}

function ShapeChange({ change }: { change: ProposalChange }) {
  const { t } = useTranslation();
  const before = change.before!;
  const after = change.after!;
  const geometryChanged =
    before.x !== after.x ||
    before.y !== after.y ||
    before.width !== after.width ||
    before.height !== after.height;
  const rectangle = (shape: typeof before) =>
    [shape.x, shape.y, shape.width, shape.height]
      .map((value) => (value / 12700).toFixed(1))
      .join(', ');
  return (
    <>
      {geometryChanged && (
        <p style={styles.detail}>
          {t('proposals.geometry')}
          <br />
          {rectangle(before)} → {rectangle(after)} pt
        </p>
      )}
      {JSON.stringify({ ...before, x: 0, y: 0, width: 0, height: 0 }) !==
        JSON.stringify({ ...after, x: 0, y: 0, width: 0, height: 0 }) &&
        change.oldText === change.newText && (
          <p style={styles.detail}>{t('proposals.appearance')}</p>
        )}
    </>
  );
}

function PreviewCanvas({
  label,
  frame,
  resolveImage,
}: {
  label: string;
  frame: SlideDisplayList;
  resolveImage: CanvasImageResolver;
}) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    const target = canvas.current;
    if (!target) return;
    const scratch = document.createElement('canvas');
    const context = scratch.getContext('2d');
    const destination = target.getContext('2d');
    if (!context || !destination) return;
    const scale = Math.min(1, 560 / frame.width);
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(scratch, frame, dpr, scale);
    setError(null);
    void paintSlide(context, frame, dpr, scale, { resolveImage })
      .then(() => {
        if (cancelled) return;
        sizeCanvasForSlide(target, frame, dpr, scale);
        destination.setTransform(1, 0, 0, 1, 0, 0);
        destination.drawImage(scratch, 0, 0);
      })
      .catch((value) => {
        if (!cancelled)
          setError(value instanceof Error ? value.message : String(value));
      });
    return () => {
      cancelled = true;
    };
  }, [frame, resolveImage]);
  return (
    <figure style={styles.figure}>
      <figcaption style={styles.caption}>{label}</figcaption>
      <canvas
        ref={canvas}
        role="img"
        aria-label={label}
        style={{
          maxWidth: '100%',
          height: 'auto',
          display: 'block',
          background: '#fff',
        }}
      />
      {error && <p role="alert">{error}</p>}
    </figure>
  );
}

const styles: Record<string, CSSProperties> = {
  panel: {
    width: 340,
    maxWidth: '45%',
    flex: '0 0 auto',
    overflowY: 'auto',
    background: '#fff',
    borderLeft: '1px solid #dbe3ee',
    padding: 16,
    fontSize: 13,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 12,
  },
  count: { color: '#64748b', marginLeft: 6 },
  card: {
    border: '1px solid #dbe3ee',
    borderLeft: '3px solid #6366f1',
    borderRadius: 8,
    padding: 12,
    marginTop: 16,
  },
  author: { color: '#4338ca' },
  note: { margin: '6px 0 12px', lineHeight: 1.5 },
  change: { borderTop: '1px solid #eef2f7', padding: '10px 0' },
  textDiff: {
    display: 'grid',
    gap: 6,
    margin: '8px 0',
    whiteSpace: 'pre-wrap',
    overflowWrap: 'anywhere',
    maxHeight: 200,
    overflowY: 'auto',
  },
  deleted: { color: '#9f1239', background: '#fff1f2', padding: '4px 6px' },
  inserted: {
    color: '#166534',
    background: '#f0fdf4',
    padding: '4px 6px',
    textDecoration: 'none',
  },
  actions: { display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 12 },
  button: {
    border: '1px solid #cbd5e1',
    borderRadius: 5,
    background: '#fff',
    color: '#334155',
    padding: '6px 10px',
    cursor: 'pointer',
    font: 'inherit',
  },
  primary: {
    border: '1px solid #4338ca',
    borderRadius: 5,
    background: '#4338ca',
    color: '#fff',
    padding: '6px 10px',
    cursor: 'pointer',
    font: 'inherit',
  },
  link: {
    background: 'none',
    border: 0,
    padding: 0,
    color: '#4338ca',
    textAlign: 'left',
    cursor: 'pointer',
    font: 'inherit',
    fontWeight: 600,
  },
  warning: {
    padding: 10,
    borderRadius: 5,
    background: '#fffbeb',
    color: '#92400e',
    lineHeight: 1.5,
  },
  detail: { color: '#64748b', lineHeight: 1.5 },
  muted: { color: '#64748b', marginTop: 24 },
  dialog: {
    width: 'min(1200px, 92vw)',
    maxHeight: '90vh',
    overflow: 'auto',
    padding: 24,
    border: '1px solid #cbd5e1',
    borderRadius: 12,
    color: '#172033',
    background: '#f8fafc',
  },
  previews: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(min(360px, 100%), 1fr))',
    gap: 16,
    margin: '20px 0',
  },
  figure: { margin: 0, minWidth: 0 },
  caption: { fontWeight: 600, marginBottom: 8 },
};
