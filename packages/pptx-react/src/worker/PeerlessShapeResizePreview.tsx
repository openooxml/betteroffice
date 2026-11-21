import { paintSlide, sizeCanvasForSlide, type SlideDisplayList } from '@betteroffice/pptx';
import { useCallback, useEffect, useRef, useSyncExternalStore } from 'react';
import type { PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export const peerlessShapeResizePreviewPaint = {
  paint: paintSlide,
};

interface PreviewJob {
  generation: number;
  request: NonNullable<NonNullable<ReturnType<PeerlessShapePreviewStore['getSnapshot']>>['request']>;
}

export function PeerlessShapeResizePreview({ run, preview, scale }: {
  run: PeerlessPresentation;
  preview: PeerlessShapePreviewStore;
  scale: number;
}) {
  const snapshot = useSyncExternalStore(preview.subscribe, preview.getSnapshot, preview.getSnapshot);
  const request = snapshot?.request;
  const canvas = useRef<HTMLCanvasElement>(null);
  const generation = useRef(0);
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
    void (async () => {
      try {
        while (desired.current) {
          const job = desired.current;
          desired.current = null;
          try {
            const frame = await run.client.previewShapeRect(
              job.request.frame, job.request.target, job.request.rect,
            );
            run.require();
            if (job.generation !== generation.current || desired.current) continue;
            const images = run.images.pin({ displayList: frame, frame: { media: [] } });
            try {
              await Promise.all([images.ready, run.ensureWholeFonts?.()]);
              if (job.generation !== generation.current || desired.current || !run.current) continue;
              const target = canvas.current;
              if (!target) continue;
              const buffer = target.ownerDocument.createElement('canvas');
              const context = buffer.getContext('2d');
              if (!context) throw new Error('Slide canvas is unavailable');
              const dpr = target.ownerDocument.defaultView?.devicePixelRatio ?? 1;
              sizeCanvasForSlide(buffer, frame as SlideDisplayList, dpr, scale);
              await peerlessShapeResizePreviewPaint.paint(context, frame as SlideDisplayList, dpr, scale,
                { resolveImage: images.resolve });
              if (job.generation !== generation.current || desired.current || !run.current || !canvas.current) continue;
              sizeCanvasForSlide(canvas.current, frame as SlideDisplayList, dpr, scale);
              const visible = canvas.current.getContext('2d');
              if (!visible) throw new Error('Slide canvas is unavailable');
              visible.setTransform(1, 0, 0, 1, 0, 0);
              visible.clearRect(0, 0, canvas.current.width, canvas.current.height);
              visible.drawImage(buffer, 0, 0);
            } finally {
              images.release();
            }
          } catch {
            if (job.generation === generation.current && !desired.current) clear();
          }
        }
      } finally {
        running.current = false;
        if (desired.current) pumpRef.current();
      }
    })();
  }, [clear, run, scale]);
  pumpRef.current = pump;

  useEffect(() => {
    const current = ++generation.current;
    desired.current = request ? { generation: current, request } : null;
    if (!request) clear();
    else pump();
  }, [clear, pump, request]);

  useEffect(() => () => {
    generation.current += 1;
    desired.current = null;
  }, []);

  if (!request) return null;
  return <canvas ref={canvas} data-testid="pptx-shape-resize-preview" aria-hidden="true"
    style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }} />;
}
