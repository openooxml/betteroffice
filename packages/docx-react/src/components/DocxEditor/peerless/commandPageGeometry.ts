import { canvasPageStack, snapCanvasLength, type RetainedFrame } from '@betteroffice/docx/layout/render';
import type { DocxCommandFramePayload } from '@betteroffice/docx/worker';
import type { DocxCommandFrame, YrsResidentCaretSnapshot } from '@betteroffice/docx/yrs';
import type { PeerlessFrameGeometry, PeerlessViewIdentity } from './viewStore';

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

/** A prepared frame with what its geometry is rebuilt from at another raster scale. */
export interface PreparedFrameGeometry {
  frame: DocxCommandFrame;
  retained: RetainedFrame;
  caret: YrsResidentCaretSnapshot;
  identity: PeerlessViewIdentity;
  geometry: PeerlessFrameGeometry;
}

/** Frame geometry with page boxes snapped to the raster scale. */
export function frameGeometry({ frame, retained, caret, identity }: Omit<PreparedFrameGeometry, 'geometry'>, rasterScale: number): PeerlessFrameGeometry {
  const pages = commandPageGeometry(retained, frame.payload, rasterScale);
  return {
    identity,
    coversThroughSeq: frame.coversThroughSeq,
    frameEpoch: retained.frameEpoch,
    layoutRevision: frame.payload.layoutRevision,
    layoutEpoch: retained.layoutEpoch,
    documentVersion: frame.payload.documentVersion,
    objects: frame.payload.geometry,
    documentWidth: frame.payload.documentWidth,
    layoutProvisional: frame.payload.layoutProvisional,
    ...(frame.payload.reservations ? { reservations: frame.payload.reservations } : {}),
    rasterScale,
    pages,
    caret: caret.caretRect ? { ...caret.caretRect,
      x: caret.caretRect.x + (pages[caret.caretRect.pageIndex]?.bounds.x ?? 0), y: caret.caretRect.y + (pages[caret.caretRect.pageIndex]?.bounds.y ?? 0), width: 1 } : null,
    selectionFocus: frame.payload.selectionFocus ? { ...frame.payload.selectionFocus,
      x: frame.payload.selectionFocus.x + (pages[frame.payload.selectionFocus.pageIndex]?.bounds.x ?? 0),
      y: frame.payload.selectionFocus.y + (pages[frame.payload.selectionFocus.pageIndex]?.bounds.y ?? 0), width: 1 } : null,
    parts: frame.payload.parts,
    selectionRects: (frame.payload.selectionRects ?? []).map(({ pageIndex, x, y, width, height }) => ({
      x: x + (pages[pageIndex]?.bounds.x ?? 0), y: y + (pages[pageIndex]?.bounds.y ?? 0), width, height,
    })),
  };
}
