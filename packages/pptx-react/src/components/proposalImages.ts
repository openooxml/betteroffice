import { decodeBrowserPresentationImage } from '@betteroffice/pptx/worker';
import type { CanvasImageResolver, ProposalDiffSlide, ShapeSnapshot } from '@betteroffice/pptx';

export function proposalPaintChanges(diff: ProposalDiffSlide) {
  const changes = [...(diff.shapeChanges ?? [])];
  const deleted = (shapes: ShapeSnapshot[]) => {
    for (const shape of shapes) {
      changes.push({ shapeId: shape.id, kind: 'deleted' });
      deleted(shape.children);
    }
  };
  for (const change of diff.proposal.changes) {
    if (change.before && !change.after) deleted([change.before]);
  }
  return changes;
}

export function proposalImages(
  diff: Pick<ProposalDiffSlide, 'snapshot' | 'media'>,
  base: CanvasImageResolver,
  errorMessage: string,
  decodeImage: (bytes: Uint8Array, message: string) => Promise<CanvasImageSource> = decodeBrowserPresentationImage,
) {
  const pending = new Map<string, string>();
  const collect = (shapes: ShapeSnapshot[]) => {
    for (const shape of shapes) {
      if (shape.pendingMedia) pending.set(`pending-media:${shape.id}`, shape.pendingMedia.base64);
      collect(shape.children);
    }
  };
  for (const slide of diff.snapshot.slides) collect(slide.shapes);
  const media = new Map(diff.media?.map(({ assetId, bytes }) => [assetId, new Uint8Array(bytes)]));
  const lease = base.acquire?.() ?? base;
  const decoded = new Map<string, Promise<CanvasImageSource>>();
  const resolveImage: CanvasImageResolver = (id) => {
    const encoded = pending.get(id);
    const supplied = media.get(id);
    const decode = () => {
      let image = decoded.get(id);
      if (!image) {
        const binary = encoded ? atob(encoded) : '';
        const bytes = supplied ?? Uint8Array.from(binary, character => character.charCodeAt(0));
        image = decodeImage(bytes, errorMessage);
        decoded.set(id, image);
      }
      return image;
    };
    if (encoded) return decode();
    if (!supplied) return lease(id);
    try {
      return Promise.resolve(lease(id)).then(source => source ?? decode(), decode);
    } catch {
      return decode();
    }
  };
  return {
    resolveImage,
    release() {
      lease.release?.();
      for (const image of decoded.values()) {
        void image.then(source => {
          if ('close' in source && typeof source.close === 'function') source.close();
        }, () => {});
      }
      decoded.clear();
    },
  };
}
