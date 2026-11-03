import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import type { PeerlessConfirmedFrame } from './peerlessViewStore';
import { createCommandFramePainter } from './commandFramePainter';

export async function peerlessFrameSource(run: PeerlessPresentation, index?: number) {
  run.require();
  await run.store.fence();
  run.require();
  const position = index ?? run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId);
  const slide = run.client.state.slides[position];
  if (!slide) throw new Error('Slide is unavailable');
  const reply = await run.client.thumbnail(slide.id);
  run.require();
  if (reply.status !== 'frame') throw new Error('Slide frame is unavailable');
  const frame: PeerlessConfirmedFrame = { frame: reply.frame, displayList: decodeCommandDisplayList(reply.frame) };
  const lease = run.images.pin(frame, reply.frame.media);
  await lease.ready;
  if (!run.current) { lease.release(); run.require(); }
  return { frame: frame.displayList, resolveImage: lease.resolve, release: lease.release, confirmed: frame, index: position };
}

export async function peerlessProposalFrameSource(run: PeerlessPresentation, id: string, slideId?: string | null) {
  run.require();
  const index = run.client.state.slides.findIndex(slide => slide.id === run.client.state.activeSlideId);
  const targetId = slideId ?? (index < 0 ? (await run.client.access!('proposalSlideStrip', id))[0]?.slideId : null);
  const diff = targetId
    ? await run.client.access!('layoutProposalDiffSlideById', id, targetId)
    : await run.client.access!('layoutProposalDiffSlide', id, index);
  const live = await run.client.access!('snapshot');
  const target = diff.snapshot.slides[0];
  const original = live.slides.find(slide => slide.id === target?.id ||
    Boolean(target?.sourcePartPath && slide.sourcePartPath === target.sourcePartPath));
  const sourceIndex = run.client.state.slides.findIndex(slide => slide.id === original?.id);
  if (!run.client.state.slides.length) return { ...diff, resolveImage: () => null, release() {} };
  const source = await peerlessFrameSource(run, sourceIndex >= 0 ? sourceIndex : Math.max(0, index));
  return { ...diff, resolveImage: source.resolveImage, release: source.release };
}

export async function peerlessExportPng(run: PeerlessPresentation): Promise<{ png: Blob; index: number }> {
  const source = await peerlessFrameSource(run);
  const canvas = document.createElement('canvas');
  const painter = createCommandFramePainter(canvas, { images: run.images, scale: 1, dpr: 1, current: () => run.current });
  try {
    painter.paintFrame(source.confirmed, { ...run.store.getSnapshot(), confirmed: source.confirmed, echo: null, caret: null,
      selection: null, selectionRects: [], composition: null });
    if (!await painter.whenPainted()) throw new Error('Slide painting failed');
    run.require();
    const png = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('PNG encoding failed')), 'image/png'));
    return { png, index: source.index };
  } finally { painter.dispose(); source.release(); }
}
