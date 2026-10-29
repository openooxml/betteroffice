import { readFontBytes } from './bytes';
import { decodeFontBytes } from './decode';
import { BUNDLED_FONTS, type BundledFontFace } from './manifest';

const decodedCache = new Map<string, Promise<ArrayBuffer>>();

export async function loadPackedFontBytes(
  face: BundledFontFace,
  resolveUrl: (face: BundledFontFace) => Promise<URL>,
  timeoutMs?: number,
): Promise<ArrayBuffer> {
  const url = await resolveUrl(face);
  const key = `${url.href}|${face.byteLength}|${timeoutMs ?? ''}`;
  let pending = decodedCache.get(key);
  if (pending === undefined) {
    pending = (async () => {
      const regular = face.file.endsWith('-Bold.ttf.zst')
        ? BUNDLED_FONTS.find((candidate) =>
            candidate.file === face.file.replace(/-Bold\.ttf\.zst$/, '-Regular.ttf.zst'))
        : undefined;
      if (face.file.endsWith('-Bold.ttf.zst') && regular === undefined) {
        throw new Error(`Missing bundled font dictionary for ${face.file}`);
      }
      const [encoded, dictionary] = await Promise.all([
        readFontBytes(face.file, url, timeoutMs),
        regular === undefined
          ? undefined
          : loadPackedFontBytes(regular, resolveUrl, timeoutMs),
      ]);
      const decoded = await decodeFontBytes(encoded, dictionary, { codec: 'zstd' });
      if (decoded.byteLength !== face.byteLength) {
        throw new Error(
          `Bundled font ${face.file} has ${decoded.byteLength} bytes; expected ${face.byteLength}`,
        );
      }
      return decoded;
    })();
    decodedCache.set(key, pending);
    const current = pending;
    pending.catch(() => {
      if (decodedCache.get(key) === current) decodedCache.delete(key);
    });
  }
  return (await pending).slice(0);
}
