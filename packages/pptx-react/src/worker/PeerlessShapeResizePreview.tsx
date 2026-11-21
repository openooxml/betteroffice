import { paintSlide, sizeCanvasForSlide, type SlideDisplayList } from '@betteroffice/pptx';
import { useCallback, useEffect, useRef, useSyncExternalStore } from 'react';
import type { PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export const peerlessShapeResizePreviewPaint = {
  paint: paintSlide,
};

interface PreviewJob {
  gesture: string;
  request: NonNullable<NonNullable<ReturnType<PeerlessShapePreviewStore['getSnapshot']>>['request']>;
  scale: number;
}

export function PeerlessShapeResizePreview({ run, preview, scale }: {
  run: PeerlessPresentation;
  preview: PeerlessShapePreviewStore;
  scale: number;
}) {
  const snapshot = useSyncExternalStore(preview.subscribe, preview.getSnapshot, preview.getSnapshot);
  const request = snapshot?.request;
  const canvas = useRef<HTMLCanvasElement>(null);
  const currentScale = useRef(scale);
  currentScale.current = scale;
  const desired = useRef<PreviewJob | null>(null);
  const running = useRef(false);
  const pumpRef = useRef<() => void>(() => {});

  const clear = useCallback(() => {
    const target = canvas.current;
    if (!target) return;
    target.width = 0;
    target.height = 0;
    target.style.width = '0px';
    target.style.height = '0px';
  }, []);

  const pump = useCallback(() => {
    if (running.current) return;
    running.current = true;
    const isCurrent = (job: PreviewJob) => job.gesture === preview.getSnapshot()?.request?.gesture &&
      job.scale === currentScale.current && run.current && canvas.current !== null;
    void (async () => {
      try {
        while (desired.current) {
          const job = desired.current;
          desired.current = null;
          if (!isCurrent(job)) continue;
          try {
            const frame = await run.client.previewShapeRect(
              job.request.frame, job.request.target, job.request.rect,
            );
            run.require();
            if (!isCurrent(job)) continue;
            const images = run.images.pin({ displayList: frame, frame: { media: [] } });
            try {
              await Promise.all([images.ready, run.ensureWholeFonts?.()]);
              if (!isCurrent(job)) continue;
              const target = canvas.current;
              if (!target) continue;
              const buffer = target.ownerDocument.createElement('canvas');
              const context = buffer.getContext('2d');
              if (!context) throw new Error('Slide canvas is unavailable');
              const dpr = target.ownerDocument.defaultView?.devicePixelRatio ?? 1;
              sizeCanvasForSlide(buffer, frame as SlideDisplayList, dpr, job.scale);
              await peerlessShapeResizePreviewPaint.paint(context, frame as SlideDisplayList, dpr, job.scale,
                { resolveImage: images.resolve });
              if (!isCurrent(job) || !canvas.current) continue;
              sizeCanvasForSlide(canvas.current, frame as SlideDisplayList, dpr, job.scale);
              const visible = canvas.current.getContext('2d');
              if (!visible) throw new Error('Slide canvas is unavailable');
              visible.setTransform(1, 0, 0, 1, 0, 0);
              visible.clearRect(0, 0, canvas.current.width, canvas.current.height);
              visible.drawImage(buffer, 0, 0);
            } finally {
              images.release();
            }
          } catch {
            if (isCurrent(job) && !desired.current) clear();
          }
        }
      } finally {
        running.current = false;
        if (desired.current) pumpRef.current();
      }
    })();
  }, [clear, preview, run]);
  pumpRef.current = pump;

  useEffect(() => {
    desired.current = request ? { gesture: request.gesture, request, scale } : null;
    if (request) pump();
  }, [pump, request, scale]);

  useEffect(() => () => {
    desired.current = null;
  }, []);

  if (!request && !snapshot?.retainCanvas) return null;
  return <canvas ref={canvas} data-testid="pptx-shape-resize-preview" aria-hidden="true"
    width={0} height={0} style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none', backgroundColor: 'white' }} />;
}
