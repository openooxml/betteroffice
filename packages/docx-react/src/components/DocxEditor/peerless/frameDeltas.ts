import { decodeFrameDelta, type DecodedFrameDelta } from '@betteroffice/docx/layout/render';
import type { DocxCommandFrame } from '@betteroffice/docx/peerless';

interface FrameDecodeCache {
  readonly revision: number;
  readonly buffers: readonly ArrayBuffer[];
  readonly byteLengths: readonly number[];
  readonly deltas: Map<ArrayBuffer, DecodedFrameDelta>;
}

const decodedFrames = new WeakMap<DocxCommandFrame, FrameDecodeCache>();

export function decodeCommandFrameDeltas(
  frame: DocxCommandFrame,
  skipEmpty = false
): readonly DecodedFrameDelta[] {
  const buffers = [...(frame.payload.pageFrames ?? []), frame.payload.frame];
  const previous = decodedFrames.get(frame);
  const cache: FrameDecodeCache = previous && previous.revision === frame.revision &&
    previous.buffers.length === buffers.length && buffers.every((bytes, index) =>
      bytes === previous.buffers[index] && bytes.byteLength === previous.byteLengths[index])
    ? previous
    : { revision: frame.revision, buffers, byteLengths: buffers.map((bytes) => bytes.byteLength), deltas: new Map() };
  if (cache !== previous) decodedFrames.set(frame, cache);
  return buffers.filter((bytes) => !skipEmpty || bytes.byteLength > 0).map((bytes) => {
    let delta = cache.deltas.get(bytes);
    if (!delta) {
      delta = decodeFrameDelta(bytes);
      cache.deltas.set(bytes, delta);
    }
    return delta;
  });
}
