import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, readdir, rename, stat, unlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export function resolveAssetCacheDir(environment = process.env) {
  const directory = environment.QUALITY_ASSET_CACHE?.trim();
  return directory ? directory : null;
}

export function assetCachePath(directory, entry) {
  if (
    !/^[a-f0-9]{64}$/.test(entry?.sha256 ?? '') ||
    !Number.isInteger(entry?.bytes) ||
    entry.bytes < 1
  )
    throw new Error('Invalid corpus asset metadata');
  return join(directory, `${entry.sha256}-${entry.bytes}`);
}

export async function readCachedAsset(directory, entry) {
  if (!directory) return null;
  const path = assetCachePath(directory, entry);
  try {
    if ((await stat(path)).size !== entry.bytes) {
      await unlink(path).catch(() => {});
      return null;
    }
    const bytes = await readFile(path);
    if (
      bytes.length !== entry.bytes ||
      createHash('sha256').update(bytes).digest('hex') !== entry.sha256
    ) {
      await unlink(path).catch(() => {});
      return null;
    }
    return bytes;
  } catch {
    return null;
  }
}

export async function writeCachedAsset(directory, entry, bytes) {
  if (!directory) return;
  try {
    await mkdir(directory, { recursive: true });
    const staging = join(directory, `.tmp-${randomUUID()}`);
    try {
      await writeFile(staging, bytes);
      await rename(staging, assetCachePath(directory, entry));
    } catch {
      await unlink(staging).catch(() => {});
    }
  } catch {}
}

export function isCacheBlobName(name) {
  if (typeof name !== 'string') return false;
  const separator = name.lastIndexOf('-');
  if (separator < 0) return false;
  const sha = name.slice(0, separator);
  const size = name.slice(separator + 1);
  return /^[a-f0-9]{64}$/.test(sha) && /^[1-9][0-9]*$/.test(size);
}

export function digestCacheNames(names) {
  const valid = names.filter(isCacheBlobName).sort();
  if (!valid.length) return '';
  return createHash('sha256').update(valid.join('\n')).digest('hex');
}

export async function digestAssetCache(directory) {
  try {
    const entries = await readdir(directory, { withFileTypes: true });
    return digestCacheNames(entries.filter((entry) => entry.isFile()).map((entry) => entry.name));
  } catch {
    return '';
  }
}

export async function fetchAsset(entry, sample, downloadFn, options = {}) {
  const { cacheDir = null, origin, maximum = 32 * 1024 * 1024 } = options;
  const url = new URL(entry?.url);
  if (
    url.origin !== origin ||
    !url.pathname.startsWith(`/${sample}/`) ||
    !Number.isInteger(entry.bytes) ||
    entry.bytes < 1 ||
    entry.bytes > maximum ||
    !/^[a-f0-9]{64}$/.test(entry.sha256)
  )
    throw new Error('Invalid corpus asset metadata');
  const cached = await readCachedAsset(cacheDir, entry);
  if (cached) return cached;
  const bytes = await downloadFn(url, entry.bytes);
  if (
    bytes.length !== entry.bytes ||
    createHash('sha256').update(bytes).digest('hex') !== entry.sha256
  )
    throw new Error(`Corpus hash mismatch: ${url}`);
  await writeCachedAsset(cacheDir, entry, bytes);
  return bytes;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  process.stdout.write(await digestAssetCache(process.argv[2] ?? ''));
}
