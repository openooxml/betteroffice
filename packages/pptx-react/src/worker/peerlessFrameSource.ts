import { decodeCommandDisplayList } from '@betteroffice/pptx/peerless';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import type { PeerlessConfirmedFrame } from './peerlessViewStore';
import { createCommandFramePainter } from './commandFramePainter';

export async function peerlessFrameSource(run: PeerlessPresentation, index?: number) {
  run.require();
  await run.store.fence();
  run.require();
  const slide = index === undefined ? run.client.state.slides.find((slide) => slide.id === run.client.state.activeSlideId)
    : run.client.state.slides[index];
  if (!slide) throw new Error('Slide is unavailable');
  const reply = await run.client.thumbnail(slide.id);
  run.require();
  if (reply.status !== 'frame') throw new Error('Slide frame is unavailable');
  const frame: PeerlessConfirmedFrame = { frame: reply.frame, displayList: decodeCommandDisplayList(reply.frame) };
  const lease = run.images.pin(frame, reply.frame.media);
  await lease.ready;
  if (!run.current) { lease.release(); run.require(); }
  return { frame: frame.displayList, resolveImage: lease.resolve, release: lease.release, confirmed: frame };
}

export async function peerlessExportPng(run: PeerlessPresentation): Promise<Blob> {
  const source = await peerlessFrameSource(run);
  const canvas = document.createElement('canvas');
  const painter = createCommandFramePainter(canvas, { images: run.images, scale: 1, dpr: 1, current: () => run.current });
  try {
    painter.paintFrame(source.confirmed, { ...run.store.getSnapshot(), confirmed: source.confirmed, echo: null, caret: null,
      selection: null, selectionRects: [], composition: null });
    if (!await painter.whenPainted()) throw new Error('Slide painting failed');
    run.require();
    return await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('PNG encoding failed')), 'image/png'));
  } finally { painter.dispose(); source.release(); }
}
