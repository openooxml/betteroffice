import { readFontAsset } from './asset';
import type { FontDecoder } from './decode';
import { openProgressiveFace, type ProgressiveFontFace, type SliceLocation } from './progressive';

export { readFontAsset } from './asset';
export {
  openProgressiveFace,
  type ProgressiveFontFace,
  type SliceLocation,
  type SliceRef,
  type SliceTransport,
} from './progressive';

const faces = new Map<string, Promise<ProgressiveFontFace>>();
const engineFaces = new WeakMap<FontDecoder, Map<string, Promise<ProgressiveFontFace>>>();

/**
 * Opens the sliced face a loader's `slices` locates. One face is shared per
 * sliced face, asset base and decoder; a failed open is dropped so a later
 * call retries.
 * @public
 */
export function openSlicedFace(location: SliceLocation): Promise<ProgressiveFontFace> {
  const { ref, base, url, timeoutMs, wasmUrl, decoder } = location;
  let cache = faces;
  if (decoder !== undefined) {
    cache = engineFaces.get(decoder) ?? new Map();
    engineFaces.set(decoder, cache);
  }
  const key = `${ref.sha256}|${base}|${timeoutMs ?? ''}|${wasmUrl?.href ?? ''}`;
  let pending = cache.get(key);
  if (pending === undefined) {
    pending = openProgressiveFace(ref, {
      read: async (file) => readFontAsset(file, await url(file), timeoutMs),
      decode: async (encoded, maxOutput) => (decoder === undefined
        ? await (await import('./decode')).decodeFontFace(encoded, undefined, { codec: 'zstd', maxOutput, wasmUrl })
        : await decoder(encoded, { codec: 'zstd', maxOutput })).font,
    });
    cache.set(key, pending);
    const current = pending;
    pending.catch(() => {
      if (cache.get(key) === current) cache.delete(key);
    });
  }
  return pending;
}
