import latinManifest from '../../fonts-latin/assets/manifest.json';
import intlManifest from '../../fonts-intl/assets/manifest.json';
import cjkManifest from '../../fonts-cjk/assets/manifest.json';
import baseLatinManifest from '../assets/base-latin.json';
import { readFontBytes } from './bytes';
import type { FontDecoder } from './decode';
import type { BundledFontFace } from './manifest';
import { openProgressiveFace, type ProgressiveFontFace, type SliceRef } from './progressive';
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
  slices?: SliceRef;
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
const cjk = readPackManifest(cjkManifest);
const baseLatin = readPackManifest(baseLatinManifest);
const baseLatinFiles = new Set(baseLatin.faces.map((face) => face.file));
const latinFiles = new Set(latin.faces.map((face) => face.file));
const cjkFiles = new Set(cjk.faces.map((face) => face.file));

export function isLatinFont(file: string): boolean {
  return latinFiles.has(file);
}

export function isBaseLatinFont(file: string): boolean {
  return baseLatinFiles.has(file);
}

export function isCjkFont(file: string): boolean {
  return cjkFiles.has(file);
}

type DecodedFace = Awaited<ReturnType<FontDecoder>>;
const decodedCache = new Map<string, Promise<DecodedFace>>();
const engineCaches = new WeakMap<FontDecoder, Map<string, Promise<DecodedFace>>>();

async function loadFace(
  face: BundledFontFace,
  manifest: PackManifest,
  resolveUrl: (face: BundledFontFace) => Promise<URL>,
  timeoutMs?: number,
  wasmUrl?: URL,
  runtime?: { decoder?: FontDecoder; /** @internal Entries this call creates are dropped once settled; concurrent requests still share one decode. */ retain?: boolean },
): Promise<DecodedFace> {
  const decoder = runtime?.decoder;
  let cache = decodedCache;
  if (decoder !== undefined) {
    cache = engineCaches.get(decoder) ?? new Map();
    engineCaches.set(decoder, cache);
  }
  const url = await resolveUrl(face);
  const entry = manifest.faces.find((entry) => entry.file === face.file);
  if (entry === undefined) throw new Error(`Unknown bundled font asset: ${face.file}`);
  const key = `${manifest.pack}|${manifest.version}|${url.href}|${entry.codec}|${entry.dict}|${entry.encodedLength}|${face.byteLength}|${entry.sha256}|${timeoutMs ?? ''}|${wasmUrl?.href ?? ''}`;
  let pending = cache.get(key);
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
          manifest, resolveUrl, timeoutMs, wasmUrl, runtime,
        ),
      ]);
      if (encoded.byteLength !== entry.encodedLength) {
        throw new Error(
          `Bundled font ${face.file} has ${encoded.byteLength} encoded bytes; expected ${entry.encodedLength}`,
        );
      }
      const dictionary = prerequisite === undefined ? undefined
        : (prerequisite.payload ?? prerequisite.font).slice(0);
      const decoded = decoder === undefined
        ? await (await import('./decode')).decodeFontFace(encoded, dictionary, { codec: entry.codec, wasmUrl })
        : await decoder(encoded, { codec: entry.codec, dict: dictionary, maxOutput: 64 * 1024 * 1024 });
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
    cache.set(key, pending);
    const current = pending;
    const forget = () => {
      if (cache.get(key) === current) cache.delete(key);
    };
    pending.then(runtime?.retain === false ? forget : undefined, forget);
  }
  return pending;
}

const progressiveFaces = new Map<string, Promise<ProgressiveFontFace>>();
const engineProgressiveFaces = new WeakMap<FontDecoder, Map<string, Promise<ProgressiveFontFace>>>();

/** @internal The slice pointer of a sliced pack face. */
export function sliceRefOf(file: string): SliceRef | undefined {
  return cjk.faces.find((entry) => entry.file === file)?.slices;
}

/**
 * @internal One shared progressive face per sliced face, asset base and decoder;
 * a failed open is dropped so a later call retries.
 */
export function loadProgressiveFace(
  face: BundledFontFace,
  resolveFile: (file: string) => Promise<URL>,
  base: string,
  timeoutMs?: number,
  wasmUrl?: URL,
  runtime?: { decoder?: FontDecoder },
  readFile: typeof readFontBytes = readFontBytes,
): Promise<ProgressiveFontFace> {
  const ref = sliceRefOf(face.file);
  if (ref === undefined) return Promise.reject(new Error(`Bundled font ${face.file} is not sliced`));
  const decoder = runtime?.decoder;
  let cache = progressiveFaces;
  if (decoder !== undefined) {
    cache = engineProgressiveFaces.get(decoder) ?? new Map();
    engineProgressiveFaces.set(decoder, cache);
  }
  const key = `${cjk.version}|${face.file}|${ref.sha256}|${base}|${timeoutMs ?? ''}|${wasmUrl?.href ?? ''}`;
  let pending = cache.get(key);
  if (pending === undefined) {
    pending = openProgressiveFace(ref, {
      read: async (file) => readFile(file, await resolveFile(file), timeoutMs),
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

export async function loadPackedFontBytes(
  face: BundledFontFace,
  resolveUrl: (face: BundledFontFace) => Promise<URL>,
  timeoutMs?: number,
  wasmUrl?: URL,
  manifest: PackManifest = isBaseLatinFont(face.file) ? baseLatin : isLatinFont(face.file) ? latin : isCjkFont(face.file) ? cjk : intl,
  runtime?: { decoder?: FontDecoder; /** @internal Entries this call creates are dropped once settled; concurrent requests still share one decode. */ retain?: boolean },
): Promise<ArrayBuffer> {
  return (await loadFace(face, manifest, resolveUrl, timeoutMs, wasmUrl, runtime)).font.slice(0);
}
