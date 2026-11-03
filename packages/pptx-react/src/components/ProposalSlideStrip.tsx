import { useEffect, useRef, useState } from 'react';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type { CanvasImageResolver, ProposalDiffSlide, ProposalStripEntry } from '@betteroffice/pptx';
import { useTranslation } from '../i18n';

interface ProposalThumbnailPreview extends ProposalDiffSlide {
  resolveImage?: CanvasImageResolver;
  release?(): void;
}

function Thumbnail({ slideId, load, resolveImage }: {
  slideId: string;
  load(slideId: string): ProposalThumbnailPreview | Promise<ProposalThumbnailPreview>;
  resolveImage?: CanvasImageResolver;
}) {
  const canvas = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    let cancelled = false;
    void Promise.resolve().then(() => load(slideId)).then(async (diff) => {
      try {
        if (cancelled) return;
        const buffer = document.createElement('canvas');
        const context = buffer.getContext('2d');
        if (!context) return;
        const scale = 128 / diff.frame.width;
        sizeCanvasForSlide(buffer, diff.frame, 1, scale);
        await paintSlide(context, diff.frame, 1, scale, { resolveImage: diff.resolveImage ?? resolveImage });
        if (cancelled || !canvas.current) return;
        sizeCanvasForSlide(canvas.current, diff.frame, 1, scale);
        canvas.current.getContext('2d')?.drawImage(buffer, 0, 0);
      } finally { diff.release?.(); }
    }).catch(() => {});
    return () => { cancelled = true; };
  }, [slideId, load, resolveImage]);
  return <canvas ref={canvas} aria-hidden="true" style={{ width: 128, display: 'block' }} />;
}

export function ProposalSlideStrip({ entries, selectedSlideId, load, onSelect, resolveImage }: {
  entries: ProposalStripEntry[];
  selectedSlideId?: string | null;
  load(slideId: string): ProposalThumbnailPreview | Promise<ProposalThumbnailPreview>;
  onSelect(slideId: string): void;
  resolveImage?: CanvasImageResolver;
}) {
  const { t } = useTranslation();
  return <div aria-label={t('proposals.slideStrip')} data-testid="pptx-proposal-slide-strip">
    {entries.map((entry, index) => <button key={entry.slideId} type="button"
      data-testid={`pptx-proposal-slide-${entry.status}`} data-slide-id={entry.slideId}
      aria-label={t('proposals.previewSlide', { number: index + 1 })}
      aria-current={selectedSlideId === entry.slideId ? 'page' : undefined}
      onClick={() => onSelect(entry.slideId)}
      style={{ display: 'block', width: '100%', marginBottom: 12, padding: 4, cursor: 'pointer',
        border: '2px solid #4338ca', borderStyle: entry.status === 'inserted' ? 'dashed' : 'solid',
        borderRadius: 5, opacity: entry.status === 'deleted' ? 0.4 : 1,
        background: selectedSlideId === entry.slideId ? '#e0e7ff' : '#fff', color: '#4338ca' }}>
      <span>{index + 1}</span>
      <Thumbnail slideId={entry.slideId} load={load} resolveImage={resolveImage} />
      {entry.status === 'inserted' && <span aria-label={t('proposals.insertedSlide')}>+</span>}
      {entry.status === 'deleted' && <span aria-label={t('proposals.deletedSlide')}>-</span>}
      {entry.status === 'moved' && <span>{t('proposals.movedFrom', { number: (entry.liveIndex ?? 0) + 1 })}</span>}
    </button>)}
  </div>;
}

export function AsyncProposalSlideStrip({ id, read, ...props }: Omit<Parameters<typeof ProposalSlideStrip>[0], 'entries'> & {
  id: string;
  read(id: string): Promise<ProposalStripEntry[]>;
}) {
  const [result, setResult] = useState<{ id: string; entries: ProposalStripEntry[] } | null>(null);
  useEffect(() => {
    let cancelled = false;
    void read(id).then(entries => { if (!cancelled) setResult({ id, entries }); }).catch(() => {});
    return () => { cancelled = true; };
  }, [id, read]);
  return <ProposalSlideStrip {...props} entries={result?.id === id ? result.entries : []} />;
}
