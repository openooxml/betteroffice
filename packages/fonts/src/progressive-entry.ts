import { readFontAsset } from './asset';
import type { FontDecoder } from './decode';
import { openProgressiveFace, type ProgressiveFontFace, type SliceLocation } from './progressive';
import type { FontCdnReporter } from '../../../shared/fontCdnDiagnostics';

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

type Faces = Map<string, Promise<ProgressiveFontFace>>;
type ScopedSliceLocation = SliceLocation & { reportCdnFailure?: FontCdnReporter };
const faceCaches = () => ({ faces: new Map<string, Promise<ProgressiveFontFace>>(), engines: new WeakMap<FontDecoder, Faces>() });
const shared = faceCaches();
const scoped = new WeakMap<FontCdnReporter, ReturnType<typeof faceCaches>>();

/**
 * Opens the sliced face a loader's `slices` locates. One face is shared per
 * sliced face, asset base and decoder; a failed open is dropped so a later
 * call retries.
 * @public
 */
export function openSlicedFace(location: SliceLocation): Promise<ProgressiveFontFace>;
export function openSlicedFace(location: ScopedSliceLocation): Promise<ProgressiveFontFace> {
  return location.prepareUrls
    ? location.prepareUrls().then((assets) => openWithUrls({ ...location, ...assets }))
    : openWithUrls(location);
}

function openWithUrls(location: ScopedSliceLocation): Promise<ProgressiveFontFace> {
  const { ref, base, url, timeoutMs, requestTimeoutMs, decode, decoder, retain, reportCdnFailure } = location;
  let caches = shared;
  if (reportCdnFailure) {
    const existing = scoped.get(reportCdnFailure);
    caches = existing ?? faceCaches();
    if (!existing) scoped.set(reportCdnFailure, caches);
  }
  let cache = caches.faces;
  if (decoder !== undefined) {
    cache = caches.engines.get(decoder) ?? new Map();
    caches.engines.set(decoder, cache);
  }
  const key = `${ref.indexSha256}|${base}|${timeoutMs ?? ''}`;
  let pending = cache.get(key);
  if (pending === undefined) {
    pending = (async () => {
      const indexUrl = await url(ref.index);
      const transportTimeout = requestTimeoutMs ? requestTimeoutMs(indexUrl) : timeoutMs;
      const transport = {
        read: async (file: string) => {
          const asset = file === ref.index ? indexUrl : await url(file);
          return readFontAsset(file, asset, requestTimeoutMs ? requestTimeoutMs(asset) : timeoutMs);
        },
        url,
        timeoutMs: transportTimeout,
        decode,
        reportCdnFailure,
      };
      return openProgressiveFace(ref, transport);
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
