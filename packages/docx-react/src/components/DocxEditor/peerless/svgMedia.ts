/** Longest edge of one decoded picture, in pixels. */
const MAX_EDGE = 4096;
/** Pixels all decoded pictures of one document may hold together. */
const MAX_PIXELS = 48 * 1024 * 1024;

export interface DecodedSvgMedia {
  source: string;
  image: ImageBitmap;
}

export interface SvgMediaDecode {
  /** The pictures decoded since the last call, now owned by the caller. */
  take(): DecodedSvgMedia[];
  /** Settles once every picture is decoded or left out. */
  done: Promise<void>;
  /** Closes the pictures nobody took and drops the ones still decoding. */
  close(): void;
}

/**
 * Rasterizes SVG pictures for the document worker, which has no SVG decoder,
 * in document order: at twice their size times the device pixel ratio, within
 * per-picture and per-document pixel caps. A picture that fails to decode is
 * left out.
 */
export function decodeSvgMedia(
  media: readonly { source: string; url: string }[] | undefined,
  pixelRatio = globalThis.devicePixelRatio || 1,
): SvgMediaDecode {
  const decoded: DecodedSvgMedia[] = [];
  let closed = false;
  const take = () => decoded.splice(0);
  const close = () => {
    closed = true;
    for (const { image } of take()) image.close();
  };
  if (!media?.length || typeof Image === 'undefined' || typeof createImageBitmap !== 'function') {
    return { take, done: Promise.resolve(), close };
  }
  const loading = media.map(async ({ source, url }) => {
    try {
      const image = new Image();
      image.src = url;
      await image.decode();
      return image.naturalWidth > 0 && image.naturalHeight > 0 ? { source, image } : null;
    } catch {
      return null;
    }
  });
  const done = (async () => {
    let budget = MAX_PIXELS;
    for (const pending of loading) {
      const entry = await pending;
      if (!entry || closed) continue;
      const { naturalWidth: width, naturalHeight: height } = entry.image;
      const scale = Math.min(2 * Math.max(1, pixelRatio), MAX_EDGE / Math.max(width, height), Math.sqrt(budget / (width * height)));
      const target = { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) };
      if (target.width * target.height > budget) continue;
      try {
        const canvas = document.createElement('canvas');
        canvas.width = target.width;
        canvas.height = target.height;
        const context = canvas.getContext('2d');
        if (!context) continue;
        context.drawImage(entry.image, 0, 0, target.width, target.height);
        const image = await createImageBitmap(canvas);
        if (closed) {
          image.close();
          continue;
        }
        decoded.push({ source: entry.source, image });
        budget -= target.width * target.height;
      } catch {
        // A picture the host cannot rasterize stays undecoded.
      }
    }
  })();
  return { take, done, close };
}
