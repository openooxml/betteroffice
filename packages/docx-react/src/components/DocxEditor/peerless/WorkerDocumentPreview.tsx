import { useLayoutEffect, useMemo, useRef } from 'react';
import { applyFrameDelta, CANVAS_PAGE_GAP_PX, decodeFrameDelta } from '@betteroffice/docx/layout/render';
import type { ResidentEngineWorkerClient, ResidentEngineWorkerFrame } from '@betteroffice/docx/yrs';

export interface WorkerPreviewJob {
  client: ResidentEngineWorkerClient;
  frame: ResidentEngineWorkerFrame;
  signal: AbortSignal;
  resolve(): void;
  reject(error: unknown): void;
}

export function WorkerDocumentPreview({ job, zoom, onPresented }: {
  job: WorkerPreviewJob;
  zoom: number;
  onPresented(frame: ResidentEngineWorkerFrame): void;
}) {
  const host = useRef<HTMLDivElement>(null);
  const pages = useMemo(() => applyFrameDelta(null, decodeFrameDelta(job.frame.frame)).pages, [job]);
  const width = Math.max(0, ...pages.map(({ page }) => page.width));
  const height = pages.reduce((total, { page }) => total + page.height + CANVAS_PAGE_GAP_PX, 0);
  useLayoutEffect(() => {
    let current = true;
    const root = host.current!;
    try {
      let y = 0;
      const surfaces = pages.filter(({ page }) => !page.unbuilt).map(({ pageId, page }) => {
        const canvas = document.createElement('canvas');
        canvas.dataset.previewPageId = pageId.toString();
        Object.assign(canvas.style, { position: 'absolute', left: `${(width - page.width) / 2}px`, top: `${y}px`,
          width: `${page.width}px`, height: `${page.height}px`, background: 'white' });
        y += page.height + CANVAS_PAGE_GAP_PX;
        root.appendChild(canvas);
        return { pageId: pageId.toString(), canvas: canvas.transferControlToOffscreen() };
      });
      void job.client.attachCanvases(surfaces, surfaces.map(({ pageId }) => pageId), globalThis.devicePixelRatio || 1, 1,
        { color: '#000', width: 2 }).then(() => {
        if (!current || job.signal.aborted) return;
        onPresented(job.frame);
        job.resolve();
      }, job.reject);
    } catch (error) { job.reject(error); }
    return () => { current = false; root.replaceChildren(); };
  }, [job, pages, width, onPresented]);
  return <div aria-label="Document preview" style={{ position: 'relative', width: width * zoom, height: height * zoom,
    marginLeft: 'auto', marginRight: 'auto', pointerEvents: 'none' }}>
    <div ref={host} style={{ position: 'relative', width, height, transform: `scale(${zoom})`, transformOrigin: 'top left' }} />
  </div>;
}
