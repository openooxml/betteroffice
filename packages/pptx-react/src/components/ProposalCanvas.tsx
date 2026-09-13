import { useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  DeckSnapshot,
  PresentationHandle,
  Proposal,
  ProposalDiffSlide,
  SlideDisplayList,
} from '@betteroffice/pptx';
import { frameBoundsForShape } from '../interactions';
import { useTranslation } from '../i18n';

export function useProposalCanvas(
  handle: PresentationHandle | null,
  proposals: Proposal[],
  snapshot: DeckSnapshot | undefined,
  slideIndex: number
) {
  const [enabled, setEnabled] = useState(true);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const slideId = snapshot?.slides[slideIndex]?.id;
  const available = useMemo(
    () =>
      proposals.filter((proposal) =>
        proposal.changes.some((change) => change.slideId === slideId)
      ),
    [proposals, slideId]
  );
  const selected =
    available.find((proposal) => proposal.id === selectedId) ??
    available[0] ??
    null;
  const result = useMemo(() => {
    if (!enabled || !handle || !selected) return { diff: null, error: null };
    try {
      return {
        diff: handle.layoutProposalDiffSlide(selected.id, slideIndex),
        error: null,
      };
    } catch (error) {
      return {
        diff: null,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }, [enabled, handle, selected, slideIndex, snapshot]);
  useEffect(() => {
    setEnabled(true);
    setSelectedId(null);
  }, [handle]);
  return {
    ...result,
    available,
    selected,
    enabled,
    setEnabled,
    select: setSelectedId,
    reviewing: enabled && selected !== null,
  };
}

type CanvasReview = ReturnType<typeof useProposalCanvas>;

export function ProposalCanvasToolbar({
  review,
  ready,
  onAccept,
  onReject,
  onDetails,
}: {
  review: CanvasReview;
  ready: boolean;
  onAccept: (id: string) => void;
  onReject: (id: string) => void;
  onDetails: () => void;
}) {
  const { t } = useTranslation();
  if (!review.selected) return null;
  const stale = review.selected.staleTargets.length > 0;
  const disabled = stale || Boolean(review.error) || (review.enabled && !ready);
  return (
    <div style={styles.toolbar} data-testid="pptx-canvas-review-toolbar">
      <div style={styles.row}>
        <label style={styles.label}>
          {t('proposals.canvasTitle')}
          <select
            aria-label={t('proposals.canvasTitle')}
            value={review.selected.id}
            onChange={(event) => {
              review.select(event.target.value);
              review.setEnabled(true);
            }}
            style={styles.select}
          >
            {review.available.map((proposal, index) => (
              <option key={proposal.id} value={proposal.id}>
                {index + 1}. {proposal.agentId}
                {proposal.note ? ` · ${proposal.note}` : ''}
              </option>
            ))}
          </select>
        </label>
        <button
          type="button"
          style={styles.button}
          data-testid="pptx-canvas-review-toggle"
          aria-pressed={review.enabled}
          onClick={() => review.setEnabled(!review.enabled)}
        >
          {t(review.enabled ? 'proposals.editSlide' : 'proposals.showDiff')}
        </button>
        <button type="button" style={styles.button} onClick={onDetails}>
          {t('proposals.details')}
        </button>
        <button
          type="button"
          style={{
            ...styles.primary,
            opacity: disabled ? 0.45 : 1,
            cursor: disabled ? 'default' : 'pointer',
          }}
          disabled={disabled}
          data-testid="pptx-canvas-proposal-accept"
          onClick={() => onAccept(review.selected!.id)}
        >
          {t('proposals.accept')}
        </button>
        <button
          type="button"
          style={styles.button}
          data-testid="pptx-canvas-proposal-reject"
          onClick={() => onReject(review.selected!.id)}
        >
          {t('proposals.reject')}
        </button>
      </div>
      {review.enabled && (
        <div style={styles.legend}>
          <span style={styles.removed}>{t('proposals.removed')}</span>
          <span style={styles.added}>{t('proposals.added')}</span>
          <span>{t('proposals.reviewHint')}</span>
        </div>
      )}
      {stale && (
        <p role="status" style={styles.warning}>
          {t('proposals.canvasStale')}
        </p>
      )}
      {review.error && (
        <p role="alert" style={styles.warning}>
          {review.error}
        </p>
      )}
    </div>
  );
}

export function ProposalCanvasOverlay({
  diff,
  current,
  frame,
  slideIndex,
  scale,
  resolveImage,
  onTarget,
  onPainted,
}: {
  diff: ProposalDiffSlide;
  current: DeckSnapshot;
  frame: SlideDisplayList;
  slideIndex: number;
  scale: number;
  resolveImage: CanvasImageResolver;
  onTarget: (slideId: string, shapeId: string) => void;
  onPainted: (diff: ProposalDiffSlide) => void;
}) {
  const { t } = useTranslation();
  const canvas = useRef<HTMLCanvasElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [painted, setPainted] = useState<ProposalDiffSlide | null>(null);
  const slideId = current.slides[slideIndex]?.id;
  const targets = diff.proposal.changes.filter(
    (change) =>
      change.slideId === slideId &&
      change.shapeId &&
      change.before &&
      change.after
  );
  useEffect(() => {
    let cancelled = false;
    const destination = canvas.current;
    if (!destination) return;
    const scratch = document.createElement('canvas');
    const ctx = scratch.getContext('2d');
    const output = destination.getContext('2d');
    if (!ctx || !output) return;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(scratch, diff.frame, dpr, scale);
    destination.dataset.ready = 'false';
    setError(null);
    void paintSlide(ctx, diff.frame, dpr, scale, {
      resolveImage,
      textChanges: diff.textChanges,
    })
      .then(() => {
        if (cancelled) return;
        sizeCanvasForSlide(destination, diff.frame, dpr, scale);
        output.setTransform(1, 0, 0, 1, 0, 0);
        output.drawImage(scratch, 0, 0);
        destination.dataset.ready = 'true';
        setPainted(diff);
        onPainted(diff);
      })
      .catch((value) => {
        if (!cancelled)
          setError(value instanceof Error ? value.message : String(value));
      });
    return () => {
      cancelled = true;
    };
  }, [diff, scale, resolveImage, onPainted]);
  return (
    <div style={styles.overlay} data-testid="pptx-canvas-proposal-diff">
      <canvas
        ref={canvas}
        role="img"
        aria-label={t('proposals.canvasTitle')}
        style={{
          ...styles.canvas,
          visibility: painted === diff ? 'visible' : 'hidden',
        }}
      />
      {painted !== diff && !error && (
        <p role="status" style={styles.paintError}>
          {t('proposals.loadingDiff')}
        </p>
      )}
      {targets.map((change) => {
        const before = frameBoundsForShape(current, frame, change.before!);
        const after = frameBoundsForShape(
          diff.snapshot,
          diff.frame,
          change.after!
        );
        if (!after) return null;
        const moved =
          before &&
          (before.x !== after.x ||
            before.y !== after.y ||
            before.width !== after.width ||
            before.height !== after.height);
        const bounds = (rect: typeof after): CSSProperties => ({
          left: rect.x * scale,
          top: rect.y * scale,
          width: Math.max(1, rect.width * scale),
          height: Math.max(1, rect.height * scale),
        });
        return (
          <div key={change.shapeId}>
            {moved && (
              <span
                data-testid="pptx-proposal-old-bounds"
                style={{
                  ...styles.outline,
                  ...bounds(before),
                  borderColor: '#b91c1c',
                  borderStyle: 'dashed',
                  pointerEvents: 'none',
                }}
                aria-hidden="true"
              >
                <span
                  style={{
                    ...styles.tag,
                    color: '#b91c1c',
                    background: '#fee2e2',
                  }}
                >
                  {t('proposals.previousPosition')}
                </span>
              </span>
            )}
            <button
              type="button"
              data-testid="pptx-proposal-canvas-target"
              aria-label={`${diff.proposal.agentId}: ${change.after!.name}`}
              onClick={() => onTarget(change.slideId, change.shapeId!)}
              style={{
                ...styles.outline,
                ...bounds(after),
                borderColor: '#16a34a',
                cursor: 'pointer',
              }}
            >
              {moved && (
                <span
                  style={{
                    ...styles.tag,
                    left: 'auto',
                    right: -1,
                    color: '#166534',
                    background: '#dcfce7',
                  }}
                >
                  {t('proposals.proposedPosition')}
                </span>
              )}
            </button>
          </div>
        );
      })}
      <div style={styles.srOnly}>
        {targets.map((change) => (
          <p key={change.shapeId}>
            {change.after!.name}: <del>{change.oldText}</del>{' '}
            <ins>{change.newText}</ins>
          </p>
        ))}
      </div>
      {error && (
        <p role="alert" style={styles.paintError}>
          {error}
        </p>
      )}
    </div>
  );
}

export function ProposalNotesDiff({
  diff,
  slideId,
}: {
  diff: ProposalDiffSlide | null;
  slideId: string;
}) {
  const { t } = useTranslation();
  const notes = diff?.proposal.changes.find(
    (change) =>
      change.slideId === slideId &&
      !change.shapeId &&
      change.oldText !== change.newText
  );
  if (!notes) return null;
  return (
    <section
      aria-label={t('notes.panelLabel')}
      style={styles.notes}
      data-testid="pptx-proposal-notes-diff"
    >
      <strong>{t('notes.panelLabel')}</strong>
      <del style={styles.removed}>{notes.oldText || t('proposals.blank')}</del>
      <ins style={styles.added}>{notes.newText || t('proposals.blank')}</ins>
    </section>
  );
}

const styles: Record<string, CSSProperties> = {
  toolbar: {
    flex: '0 0 auto',
    padding: '10px 14px',
    background: '#fff',
    borderBottom: '1px solid #dbe3ee',
    fontSize: 12,
  },
  row: { display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 },
  label: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: 8,
    minWidth: 0,
    fontWeight: 600,
  },
  select: {
    maxWidth: 260,
    minWidth: 80,
    padding: 5,
    border: '1px solid #cbd5e1',
    borderRadius: 4,
    font: 'inherit',
  },
  legend: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: 12,
    marginTop: 8,
    color: '#64748b',
  },
  removed: {
    color: '#b91c1c',
    background: '#fee2e2',
    textDecoration: 'line-through',
    padding: '2px 5px',
  },
  added: {
    color: '#166534',
    background: '#dcfce7',
    textDecoration: 'underline',
    padding: '2px 5px',
  },
  button: {
    padding: '5px 9px',
    border: '1px solid #cbd5e1',
    borderRadius: 5,
    background: '#fff',
    color: '#334155',
    cursor: 'pointer',
    font: 'inherit',
  },
  primary: {
    padding: '5px 9px',
    border: '1px solid #4338ca',
    borderRadius: 5,
    background: '#4338ca',
    color: '#fff',
    cursor: 'pointer',
    font: 'inherit',
  },
  warning: { margin: '8px 0 0', color: '#92400e' },
  overlay: {
    position: 'absolute',
    inset: 0,
    overflow: 'hidden',
    background: '#fff',
  },
  canvas: {
    position: 'absolute',
    inset: 0,
    display: 'block',
    pointerEvents: 'none',
  },
  outline: {
    position: 'absolute',
    border: '1px solid',
    background: 'transparent',
    padding: 0,
  },
  tag: {
    position: 'absolute',
    left: -1,
    top: -18,
    fontSize: 11,
    lineHeight: '16px',
    padding: '0 4px',
    whiteSpace: 'nowrap',
  },
  srOnly: {
    position: 'absolute',
    width: 1,
    height: 1,
    padding: 0,
    margin: -1,
    overflow: 'hidden',
    clipPath: 'inset(50%)',
    whiteSpace: 'nowrap',
  },
  notes: {
    display: 'grid',
    gap: 6,
    padding: '10px 14px',
    maxHeight: 160,
    overflow: 'auto',
    background: '#fff',
    fontSize: 12,
    whiteSpace: 'pre-wrap',
  },
  paintError: {
    position: 'absolute',
    top: 8,
    left: 8,
    right: 8,
    padding: 8,
    background: '#fff1f2',
    color: '#b91c1c',
  },
};
