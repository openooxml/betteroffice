import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, test } from 'bun:test';
import {
  assetCachePath,
  digestAssetCache,
  digestCacheNames,
  fetchAsset,
  isCacheBlobName,
  readCachedAsset,
  resolveAssetCacheDir,
  writeCachedAsset,
} from './asset-cache.mjs';

const origin = 'https://corpus.betteroffice.dev';
const sample = 'demo-sample';

function entryFor(payload: Buffer) {
  return {
    url: `${origin}/${sample}/reference/page_0001.png`,
    bytes: payload.length,
    sha256: createHash('sha256').update(payload).digest('hex'),
  };
}

function corruptSameLength(payload: Buffer) {
  const corrupt = Buffer.from(payload);
  corrupt[0] ^= 0xff;
  if (corrupt.equals(payload)) corrupt[0] ^= 0x01;
  return corrupt;
}

async function scratch() {
  const directory = await mkdtemp(join(tmpdir(), 'fidelity-assets-'));
  return directory;
}

test('warm cache serves verified bytes with no network call', async () => {
  const directory = await scratch();
  try {
    const payload = Buffer.from('reference-png-bytes');
    const entry = entryFor(payload);
    let calls = 0;
    const fetched = await fetchAsset(entry, sample, async () => {
      calls++;
      return payload;
    }, { cacheDir: directory, origin });
    expect(fetched).toEqual(payload);
    expect(calls).toBe(1);
    const warm = await fetchAsset(entry, sample, async () => {
      throw new Error('Unexpected network call');
    }, { cacheDir: directory, origin });
    expect(warm).toEqual(payload);
    expect(calls).toBe(1);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('same-length hash-corrupt cache is discarded and refetched', async () => {
  const directory = await scratch();
  try {
    const payload = Buffer.from('genuine-bytes-1234');
    const entry = entryFor(payload);
    await mkdir(directory, { recursive: true });
    await writeFile(assetCachePath(directory, entry), corruptSameLength(payload));
    let calls = 0;
    const fetched = await fetchAsset(entry, sample, async () => {
      calls++;
      return payload;
    }, { cacheDir: directory, origin });
    expect(fetched).toEqual(payload);
    expect(calls).toBe(1);
    expect(await readCachedAsset(directory, entry)).toEqual(payload);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('truncated cache is discarded and refetched', async () => {
  const directory = await scratch();
  try {
    const payload = Buffer.from('full-payload-bytes');
    const entry = entryFor(payload);
    await mkdir(directory, { recursive: true });
    await writeFile(assetCachePath(directory, entry), payload.subarray(0, 4));
    let calls = 0;
    const fetched = await fetchAsset(entry, sample, async () => {
      calls++;
      return payload;
    }, { cacheDir: directory, origin });
    expect(fetched).toEqual(payload);
    expect(calls).toBe(1);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('network hash mismatch fails instead of scoring stale data', async () => {
  const directory = await scratch();
  try {
    const expected = Buffer.from('expected-bytes!!');
    const entry = entryFor(expected);
    const tampered = corruptSameLength(expected);
    expect(tampered.length).toBe(entry.bytes);
    let calls = 0;
    await expect(
      fetchAsset(entry, sample, async () => {
        calls++;
        return tampered;
      }, { cacheDir: directory, origin })
    ).rejects.toThrow('Corpus hash mismatch');
    expect(calls).toBe(1);
    expect(await readCachedAsset(directory, entry)).toBeNull();
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('changed SHA for the same URL does not reuse the old blob', async () => {
  const directory = await scratch();
  try {
    const oldPayload = Buffer.from('old-revision-payload-1');
    const newPayload = Buffer.from('new-revision-payload-2');
    expect(newPayload.length).toBe(oldPayload.length);
    const oldEntry = entryFor(oldPayload);
    const newEntry = { ...entryFor(newPayload), url: oldEntry.url };
    await mkdir(directory, { recursive: true });
    await writeFile(assetCachePath(directory, oldEntry), oldPayload);
    let calls = 0;
    const fetched = await fetchAsset(newEntry, sample, async () => {
      calls++;
      return newPayload;
    }, { cacheDir: directory, origin });
    expect(fetched).toEqual(newPayload);
    expect(calls).toBe(1);
    expect(await readCachedAsset(directory, newEntry)).toEqual(newPayload);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('invalid asset metadata fails before any network call', async () => {
  const payload = Buffer.from('payload');
  const entry = entryFor(payload);
  const cases = [
    { ...entry, url: `https://evil.example/${sample}/reference/page_0001.png` },
    { ...entry, url: `${origin}/other-sample/reference/page_0001.png` },
    { ...entry, bytes: -1 },
    { ...entry, sha256: 'z'.repeat(64) },
    { ...entry, bytes: 40 * 1024 * 1024 },
  ];
  for (const invalid of cases) {
    await expect(
      fetchAsset(invalid, sample, async () => {
        throw new Error('Unexpected network call');
      }, { cacheDir: null, origin })
    ).rejects.toThrow('Invalid corpus asset metadata');
  }
  await expect(
    fetchAsset({ ...entry, bytes: payload.length + 1 }, sample, async () => payload, {
      cacheDir: null,
      origin,
    })
  ).rejects.toThrow('Corpus hash mismatch');
});

test('cache storage failures still return verified bytes', async () => {
  const directory = await scratch();
  try {
    const blocker = join(directory, 'blocker');
    await writeFile(blocker, Buffer.from('not-a-directory'));
    const payload = Buffer.from('payload-bytes');
    const entry = entryFor(payload);
    const fetched = await fetchAsset(entry, sample, async () => payload, {
      cacheDir: join(blocker, 'nested'),
      origin,
    });
    expect(fetched).toEqual(payload);
    await writeCachedAsset(join(blocker, 'nested'), entry, payload);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('cache digest is stable regardless of listing order', () => {
  const a = `${'a'.repeat(64)}-12`;
  const b = `${'b'.repeat(64)}-34`;
  expect(isCacheBlobName(a)).toBe(true);
  expect(isCacheBlobName(b)).toBe(true);
  expect(digestCacheNames([a, b])).toBe(digestCacheNames([b, a]));
  expect(digestCacheNames([a, b])).toBe(
    createHash('sha256').update([a, b].sort().join('\n')).digest('hex')
  );
});

test('cache digest ignores staging files and invalid names', () => {
  const valid = `${'c'.repeat(64)}-56`;
  const names = [
    valid,
    '.tmp-0193abcd-xyz',
    '.DS_Store',
    'garbage',
    `${'Z'.repeat(64)}-56`,
    `${'c'.repeat(64)}-0`,
    `${'c'.repeat(63)}-56`,
  ];
  expect(digestCacheNames(names)).toBe(digestCacheNames([valid]));
  expect(digestCacheNames([])).toBe('');
  expect(digestCacheNames(['.tmp-abc', 'nope'])).toBe('');
});

test('cache digest changes when assets are added', async () => {
  const directory = await scratch();
  try {
    expect(await digestAssetCache(directory)).toBe('');
    expect(await digestAssetCache(join(directory, 'missing'))).toBe('');
    const first = Buffer.from('first-asset-payload');
    const second = Buffer.from('second-asset-payload');
    const firstEntry = entryFor(first);
    const secondEntry = entryFor(second);
    await writeCachedAsset(directory, firstEntry, first);
    const one = await digestAssetCache(directory);
    expect(one).toBe(digestCacheNames([`${firstEntry.sha256}-${firstEntry.bytes}`]));
    await writeFile(join(directory, '.tmp-staging'), Buffer.from('staging'));
    expect(await digestAssetCache(directory)).toBe(one);
    await writeCachedAsset(directory, secondEntry, second);
    const two = await digestAssetCache(directory);
    expect(two).not.toBe(one);
    expect(two).toBe(
      digestCacheNames([
        `${firstEntry.sha256}-${firstEntry.bytes}`,
        `${secondEntry.sha256}-${secondEntry.bytes}`,
      ])
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('asset cache directory is optional and blank-safe', () => {
  expect(resolveAssetCacheDir({})).toBeNull();
  expect(resolveAssetCacheDir({ QUALITY_ASSET_CACHE: '  ' })).toBeNull();
  expect(resolveAssetCacheDir({ QUALITY_ASSET_CACHE: '/tmp/fidelity' })).toBe('/tmp/fidelity');
});
