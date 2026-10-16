import { useLayoutEffect, useRef } from 'react';
import { applyFrameDelta, CANVAS_PAGE_GAP_PX, decodeFrameDelta, type RetainedFramePage } from '@betteroffice/docx/layout/render';
import type { ResidentEngineWorkerClient, ResidentEngineWorkerFrame } from '@betteroffice/docx/yrs';

/** @internal How long the open waits for the preview to paint before it lays out the document anyway. */
export const peerlessPreviewPaint = { timeoutMs: 5_000 };

export interface WorkerPreviewPages {
  /** Every page of the preview frame, built or not. */
  pages: readonly RetainedFramePage[];
  /** The leading pages that paint as the opened document will. */
  presented: readonly RetainedFramePage[];
}

export interface WorkerPreviewJob {
  client: ResidentEngineWorkerClient;
  frame: ResidentEngineWorkerFrame;
  pages: WorkerPreviewPages;
  live(): boolean;
  resolve(): void;
  reject(error: unknown): void;
}

/** Decodes a preview frame; a page is presented only when it and every page before it are built. */
export function workerPreviewPages(frame: ResidentEngineWorkerFrame): WorkerPreviewPages {
  const { pages } = applyFrameDelta(null, decodeFrameDelta(frame.frame));
  const presented: RetainedFramePage[] = [];
  for (const entry of pages) {
    if (entry.pageIndex !== presented.length || entry.page.unbuilt) break;
    presented.push(entry);
  }
  return { pages, presented };
}

export function releaseWorkerPreview(client: ResidentEngineWorkerClient): void {
  void client.attachCanvases([], [], globalThis.devicePixelRatio || 1, 1, { color: '#000', width: 2 }).catch(() => {});
}

export function WorkerDocumentPreview({ job, zoom, onPresented }: {
  job: WorkerPreviewJob;
  zoom: number;
  onPresented(frame: ResidentEngineWorkerFrame): void;
}) {
  const host = useRef<HTMLDivElement>(null);
  const { pages, presented } = job.pages;
  const width = Math.max(0, ...pages.map(({ page }) => page.width));
  const height = pages.reduce((total, { page }) => total + page.height + CANVAS_PAGE_GAP_PX, 0);
  useLayoutEffect(() => {
    let current = true;
    const root = host.current!;
    try {
      let y = 0;
      const surfaces = presented.map(({ pageId, pageIndex, page }) => {
        const element = document.createElement('div');
        element.className = 'paged-editor__page canvas-page';
        element.dataset.pageIndex = String(pageIndex);
        Object.assign(element.style, { position: 'absolute', left: `${(width - page.width) / 2}px`, top: `${y}px`,
          width: `${page.width}px`, height: `${page.height}px`, background: 'white' });
        const canvas = document.createElement('canvas');
        canvas.dataset.previewPageId = pageId.toString();
        canvas.dataset.pageIndex = String(pageIndex);
        Object.assign(canvas.style, { display: 'block', width: '100%', height: '100%' });
        element.appendChild(canvas);
        y += page.height + CANVAS_PAGE_GAP_PX;
        root.appendChild(element);
        return { pageId: pageId.toString(), canvas: canvas.transferControlToOffscreen() };
      });
      void job.client.attachCanvases(surfaces, surfaces.map(({ pageId }) => pageId), globalThis.devicePixelRatio || 1, 1,
        { color: '#000', width: 2 }).then(() => {
        if (!current || !job.live()) return;
        onPresented(job.frame);
        job.resolve();
      }, job.reject);
    } catch (error) { job.reject(error); }
    return () => { current = false; root.replaceChildren(); };
  }, [job, presented, width, onPresented]);
  return <div aria-label="Document preview" style={{ position: 'relative', width: width * zoom, height: height * zoom,
    marginLeft: 'auto', marginRight: 'auto', pointerEvents: 'none' }}>
    <div ref={host} style={{ position: 'relative', width, height, transform: `scale(${zoom})`, transformOrigin: 'top left' }} />
  </div>;
}
