import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import type { ShapeSnapshot } from '@betteroffice/pptx';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import type { PeerlessConfirmedFrame } from './peerlessViewStore';

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
  run.require();
  const prepareImages = async (shapes: ShapeSnapshot[]): Promise<void> => {
    for (const shape of shapes) {
      if (shape.pendingMedia?.contentType === 'image/tiff') {
        const binary = atob(shape.pendingMedia.base64);
        const bytes = Uint8Array.from(binary, character => character.charCodeAt(0));
        const converted = await run.client.preparePicture?.(bytes.buffer);
        run.require();
        if (!converted) throw new Error('Picture conversion unavailable');
        const png = new Uint8Array(converted);
        const chunks: string[] = [];
        for (let start = 0; start < png.length; start += 0x8000) {
          chunks.push(String.fromCharCode(...png.subarray(start, start + 0x8000)));
        }
        shape.pendingMedia = { contentType: 'image/png', base64: btoa(chunks.join('')) };
      }
      await prepareImages(shape.children);
    }
  };
  for (const slide of diff.snapshot.slides) await prepareImages(slide.shapes);
  run.require();
  const lease = run.images.pin({ displayList: diff.frame, frame: { media: diff.media ?? [] } });
  try {
    await lease.ready;
    run.require();
    return { ...diff, resolveImage: lease.resolve, release: lease.release };
  } catch (error) { lease.release(); throw error; }
}

export async function peerlessExportPng(run: PeerlessPresentation): Promise<{ png: Blob; index: number }> {
  run.require();
  await run.store.fence();
  run.require();
  const slideId = run.client.state.activeSlideId;
  const index = run.client.state.slides.findIndex((slide) => slide.id === slideId);
  if (index < 0 || !slideId || !run.client.renderPng) throw new Error('Slide is unavailable');
  const bytes = await run.client.renderPng({ slide: slideId }, { scale: 1 });
  run.require();
  return { png: new Blob([new Uint8Array(bytes)], { type: 'image/png' }), index };
}
