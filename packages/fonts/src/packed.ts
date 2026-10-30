import latinManifest from '../../fonts-latin/assets/manifest.json';
import intlManifest from '../../fonts-intl/assets/manifest.json';
import { readFontBytes } from './bytes';
import { decodeFontFace } from './decode';
import type { BundledFontFace } from './manifest';
import { sha256 } from './sha256';

interface PackFace {
  file: string;
  source: string;
  family: string;
  codec: 'zstd' | 'woff2t+zstd';
  dict: string | null;
  encodedLength: number;
  byteLength: number;
  sha256: string;
}

interface PackManifest {
  format: number;
  schema?: number;
  pack: string;
  version: string;
  faces: PackFace[];
}

export function readPackManifest(value: unknown): PackManifest {
  const manifest = value as PackManifest | null;
  if (manifest?.format !== 1 || (manifest.schema !== undefined && manifest.schema !== 1)
    || !Array.isArray(manifest.faces)) {
    throw new Error('Unsupported font pack manifest');
  }
  for (const face of manifest.faces) {
    if (face.codec !== 'zstd' && face.codec !== 'woff2t+zstd') {
      throw new Error(`Unsupported font face codec: ${face.codec}`);
    }
  }
  return manifest;
}

const latin = readPackManifest(latinManifest);
const intl = readPackManifest(intlManifest);
const latinFiles = new Set(latin.faces.map((face) => face.file));

export function isLatinFont(file: string): boolean {
  return latinFiles.has(file);
}

type DecodedFace = Awaited<ReturnType<typeof decodeFontFace>>;
const decodedCache = new Map<string, Promise<DecodedFace>>();

async function loadFace(
  face: BundledFontFace,
  manifest: PackManifest,
  resolveUrl: (face: BundledFontFace) => Promise<URL>,
  timeoutMs?: number,
  wasmUrl?: URL,
): Promise<DecodedFace> {
  const url = await resolveUrl(face);
  const entry = manifest.faces.find((entry) => entry.file === face.file);
  if (entry === undefined) throw new Error(`Unknown bundled font asset: ${face.file}`);
  const key = `${url.href}|${entry.codec}|${entry.dict}|${entry.encodedLength}|${face.byteLength}|${entry.sha256}|${timeoutMs ?? ''}|${wasmUrl?.href ?? ''}`;
  let pending = decodedCache.get(key);
  if (pending === undefined) {
    pending = (async () => {
      const regular = entry.dict === null ? undefined
        : manifest.faces.find((candidate) => candidate.file === entry.dict);
      if (entry.dict !== null && (regular === undefined || regular.dict !== null
        || regular.codec !== entry.codec || regular.family !== entry.family)) {
        throw new Error(`Missing bundled font dictionary for ${face.file}`);
      }
      const [encoded, prerequisite] = await Promise.all([
        readFontBytes(face.file, url, timeoutMs),
        regular === undefined ? undefined : loadFace(
          { ...face, file: regular.file, byteLength: regular.byteLength },
          manifest, resolveUrl, timeoutMs, wasmUrl,
        ),
      ]);
      if (encoded.byteLength !== entry.encodedLength) {
        throw new Error(
          `Bundled font ${face.file} has ${encoded.byteLength} encoded bytes; expected ${entry.encodedLength}`,
        );
      }
      const dictionary = prerequisite === undefined ? undefined
        : (prerequisite.payload ?? prerequisite.font).slice(0);
      const decoded = await decodeFontFace(encoded, dictionary, { codec: entry.codec, wasmUrl });
      if (decoded.font.byteLength !== face.byteLength) {
        throw new Error(
          `Bundled font ${face.file} has ${decoded.font.byteLength} bytes; expected ${face.byteLength}`,
        );
      }
      if (await sha256(decoded.font) !== entry.sha256) {
        throw new Error(`Bundled font ${face.file} SHA-256 mismatch`);
      }
      if (entry.codec === 'woff2t+zstd' && decoded.payload === undefined) {
        throw new Error(`Missing bundled font dictionary payload for ${face.file}`);
      }
      return decoded;
    })();
    decodedCache.set(key, pending);
    const current = pending;
    pending.catch(() => {
      if (decodedCache.get(key) === current) decodedCache.delete(key);
    });
  }
  return pending;
}

export async function loadPackedFontBytes(
  face: BundledFontFace,
  resolveUrl: (face: BundledFontFace) => Promise<URL>,
  timeoutMs?: number,
  wasmUrl?: URL,
  manifest: PackManifest = isLatinFont(face.file) ? latin : intl,
): Promise<ArrayBuffer> {
  return (await loadFace(face, manifest, resolveUrl, timeoutMs, wasmUrl)).font.slice(0);
}
