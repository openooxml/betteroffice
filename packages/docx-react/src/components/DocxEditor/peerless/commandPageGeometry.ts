import { canvasPageStack, snapCanvasLength, type RetainedFrame } from '@betteroffice/docx/layout/render';
import type { DocxCommandFramePayload } from '@betteroffice/docx/worker';
import type { PeerlessFrameGeometry } from './viewStore';

export function commandPageGeometry(
  frame: RetainedFrame,
  payload: Pick<DocxCommandFramePayload, 'finalizedPages' | 'documentWidth'>,
  rasterScale = 1
): PeerlessFrameGeometry['pages'] {
  const certified = payload.finalizedPages === undefined
    ? frame.pages : frame.pages.slice(0, payload.finalizedPages);
  const scale = { rasterScale, unitScale: rasterScale };
  const boxes = canvasPageStack(certified.map(({ page }) => page), { scale });
  const width = Math.max(payload.documentWidth ?? 0, ...boxes.map((box) => box.width));
  return certified.map(({ pageId, pageIndex, page }, index) => {
    const { top, width: pageWidth, height } = boxes[index]!;
    const left = snapCanvasLength((width - pageWidth) / 2, scale);
    return {
      page: { pageId: pageId.toString(), pageIndex },
      bounds: { x: left, y: top, width: pageWidth, height },
      transform: [1, 0, 0, 1, left, top],
      displayPage: page,
    };
  });
}
