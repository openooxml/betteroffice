import { readFontAsset } from './asset';
import type { FontDecoder } from './decode';
import { openProgressiveFace, type ProgressiveFontFace, type SliceLocation } from './progressive';

export { readFontAsset } from './asset';
export {
  openProgressiveFace,
  sliceSourceOf,
  type ProgressiveFontFace,
  type SliceLocation,
  type SliceRef,
  type SliceSource,
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
  return location.prepareUrls
    ? location.prepareUrls().then((assets) => openWithUrls({ ...location, ...assets }))
    : openWithUrls(location);
}

function openWithUrls(location: SliceLocation): Promise<ProgressiveFontFace> {
  const { ref, base, url, timeoutMs, requestTimeoutMs, decode, decoder, retain } = location;
  let cache = faces;
  if (decoder !== undefined) {
    cache = engineFaces.get(decoder) ?? new Map();
    engineFaces.set(decoder, cache);
  }
  const key = `${ref.indexSha256}|${base}|${timeoutMs ?? ''}|${url.implicitCdn?.(ref.index) ?? false}`;
  let pending = cache.get(key);
  if (pending === undefined) {
    pending = (async () => {
      const indexUrl = await url(ref.index);
      const transportTimeout = requestTimeoutMs ? requestTimeoutMs(indexUrl) : timeoutMs;
      return openProgressiveFace(ref, {
        read: async (file) => {
          const asset = file === ref.index ? indexUrl : await url(file);
          const request = url.implicitCdn?.(file) ? 'cjk' : requestTimeoutMs ? requestTimeoutMs(asset) : timeoutMs;
          return readFontAsset(file, asset, request);
        },
        url,
        implicitCdn: url.implicitCdn,
        timeoutMs: transportTimeout,
        decode,
      });
    })();
    cache.set(key, pending);
    const current = pending;
    const forget = () => {
      if (cache.get(key) === current) cache.delete(key);
    };
    pending.then(retain === false ? forget : undefined, forget);
  }
  return pending;
}
