import { CANVAS_PAGE_GAP_PX, type RetainedFrame } from '@betteroffice/docx/layout/render';
import type { DocxCommandFramePayload } from '@betteroffice/docx/peerless';
import type { PeerlessFrameGeometry } from './viewStore';

export function commandPageGeometry(
  frame: RetainedFrame,
  payload: Pick<DocxCommandFramePayload, 'finalizedPages' | 'documentWidth'>
): PeerlessFrameGeometry['pages'] {
  const certified = payload.finalizedPages === undefined
    ? frame.pages : frame.pages.slice(0, payload.finalizedPages);
  const width = Math.max(payload.documentWidth ?? 0,
    ...certified.map(({ page }) => page.width));
  let y = 0;
  return certified.map(({ pageId, pageIndex, page }) => {
    const top = y;
    const left = (width - page.width) / 2;
    y += page.height + CANVAS_PAGE_GAP_PX;
    return {
      page: { pageId: pageId.toString(), pageIndex },
      bounds: { x: left, y: top, width: page.width, height: page.height },
      transform: [1, 0, 0, 1, left, top],
      displayPage: page,
    };
  });
}
