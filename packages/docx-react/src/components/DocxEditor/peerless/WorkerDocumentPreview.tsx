import { useLayoutEffect, useMemo, useRef } from 'react';
import { applyFrameDelta, CANVAS_PAGE_GAP_PX, canvasPageStack, decodeFrameDelta, snapCanvasLength, type RetainedFramePage } from '@betteroffice/docx/layout/render';
import type { ResidentEngineWorkerClient, ResidentEngineWorkerFrame } from '@betteroffice/docx/yrs';
import { useTranslation } from '../../../i18n';

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
  const { t } = useTranslation();
  const host = useRef<HTMLDivElement>(null);
  const { pages, presented } = job.pages;
  const rasterScale = (globalThis.devicePixelRatio || 1) * zoom;
  const scale = useMemo(() => ({ rasterScale, unitScale: rasterScale }), [rasterScale]);
  const boxes = useMemo(() => canvasPageStack(pages.map(({ page }) => page), { scale }), [pages, scale]);
  const width = Math.max(0, ...boxes.map((box) => box.width));
  const last = boxes.at(-1);
  const height = last ? last.top + last.height + snapCanvasLength(CANVAS_PAGE_GAP_PX, scale) : 0;
  useLayoutEffect(() => {
    let current = true;
    const root = host.current!;
    try {
      const surfaces = presented.map(({ pageId, pageIndex }, index) => {
        const box = boxes[index]!;
        const element = document.createElement('div');
        element.className = 'paged-editor__page canvas-page';
        element.dataset.pageIndex = String(pageIndex);
        Object.assign(element.style, { position: 'absolute', left: `${snapCanvasLength((width - box.width) / 2, scale)}px`,
          top: `${box.top}px`, width: `${box.width}px`, height: `${box.height}px`, background: 'white' });
        const canvas = document.createElement('canvas');
        canvas.dataset.previewPageId = pageId.toString();
        canvas.dataset.pageIndex = String(pageIndex);
        Object.assign(canvas.style, { display: 'block', width: '100%', height: '100%' });
        element.appendChild(canvas);
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
  }, [job, presented, boxes, scale, width, onPresented]);
  return <div aria-label={t('editor.documentPreview')} style={{ position: 'relative', width: width * zoom, height: height * zoom,
    marginLeft: 'auto', marginRight: 'auto', pointerEvents: 'none' }}>
    <div ref={host} style={{ position: 'relative', width, height, transform: `scale(${zoom})`, transformOrigin: 'top left' }} />
  </div>;
}
