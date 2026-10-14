import { transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import { graphemeBoundaries } from '../../../../shared/textSelection';
import type { SlideDisplayList, SlidePrimitive } from '../types';
import {
  decodeTiffImage, presentationDisplayListJson, type PresentationHandle,
} from '../wasm/loader';
import type { PresentationCommandFrame } from './commandProtocol';
import { frameAssetIds } from './frame';

type FrameTags = Omit<PresentationCommandFrame, 'displayList' | 'media'>;

export function buildCommandFrame(
  handle: PresentationHandle, tags: FrameTags, sentMedia: Set<string>, parsed?: (list: SlideDisplayList) => void, cached?: SlideDisplayList
): TransferResult<PresentationCommandFrame> {
  const list = cached ?? JSON.parse(presentationDisplayListJson(handle, tags.slideIndex)) as SlideDisplayList;
  function visit(primitives: SlidePrimitive[]): void {
    for (const primitive of primitives) {
      if (primitive.kind === 'textBox') {
        const text = primitive.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
        const boundaries = [...graphemeBoundaries(text)];
        let boundaryIndex = 0;
        for (const line of primitive.lines) {
          const stops = [...line.caretStops].sort((left, right) => left.position - right.position);
          while (boundaryIndex < boundaries.length && boundaries[boundaryIndex]! < line.start) boundaryIndex++;
          const normalized: typeof stops = [];
          let stopIndex = 0;
          for (let index = boundaryIndex; index < boundaries.length && boundaries[index]! <= line.end; index++) {
            const position = boundaries[index]!;
            while (stopIndex < stops.length && stops[stopIndex]!.position < position) stopIndex++;
            const after = stops[stopIndex];
            const before = stops[stopIndex - 1];
            if (after?.position === position) normalized.push(after);
            else if (before && after) normalized.push({ position,
              x: before.x + (after.x - before.x) * (position - before.position) / (after.position - before.position) });
          }
          line.caretStops = normalized;
        }
      } else if (primitive.kind === 'table' || primitive.kind === 'chart') visit(primitive.primitives);
    }
  }
  if (!cached) visit(list.primitives);
  parsed?.(list);
  const displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
  const media: PresentationCommandFrame['media'] = [];
  for (const assetId of frameAssetIds(list)) {
    if (sentMedia.has(assetId)) continue;
    let bytes: Uint8Array;
    try { bytes = handle.mediaBytes(assetId); } catch { continue; }
    if (isTiff(bytes) && bytes.byteLength <= MAX_TIFF_BYTES) {
      try { bytes = decodeTiffImage(bytes); } catch {}
    }
    media.push({ assetId, bytes: new Uint8Array(bytes).buffer });
  }
  for (const { assetId } of media) sentMedia.add(assetId);
  return transferable({ ...tags, displayList, media },
    [displayList, ...media.map(({ bytes }) => bytes)]);
}
