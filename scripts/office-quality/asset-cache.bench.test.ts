import { createHash, randomBytes } from 'node:crypto';
import { mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, test } from 'bun:test';
import { fetchAsset } from './asset-cache.mjs';
import { download } from './download.mjs';

test('cold fetch populates the cache and warm fetch reuses identical bytes', async () => {
  const payloads = Array.from({ length: 4 }, () => randomBytes(128 * 1024));
  const requests = new Map<string, number>();
  const server = createServer((request, response) => {
    requests.set(request.url ?? '', (requests.get(request.url ?? '') ?? 0) + 1);
    const index = Number((request.url ?? '').match(/asset-(\d+)/)?.[1]);
    response.writeHead(200, { 'Content-Type': 'application/octet-stream' });
    response.end(payloads[index]);
  });
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done));
  const origin = `http://127.0.0.1:${(server.address() as { port: number }).port}`;
  const directory = await mkdtemp(join(tmpdir(), 'fidelity-bench-'));
  try {
    const entries = payloads.map((payload, index) => ({
      url: `${origin}/sample/asset-${index}.bin`,
      bytes: payload.length,
      sha256: createHash('sha256').update(payload).digest('hex'),
    }));
    const fetched = (entry: (typeof entries)[number]) =>
      fetchAsset(entry, 'sample', download, { cacheDir: directory, origin });

    const coldStart = performance.now();
    const cold = [];
    for (const entry of entries) cold.push(await fetched(entry));
    const coldMs = performance.now() - coldStart;
    const coldRequests = [...requests.values()].reduce((sum, count) => sum + count, 0);

    const warmStart = performance.now();
    const warm = [];
    for (const entry of entries)
      warm.push(
        await fetchAsset(entry, 'sample', async () => {
          throw new Error('Unexpected network call');
        }, { cacheDir: directory, origin })
      );
    const warmMs = performance.now() - warmStart;
    const totalRequests = [...requests.values()].reduce((sum, count) => sum + count, 0);

    expect(coldRequests).toBe(entries.length);
    expect(totalRequests).toBe(entries.length);
    for (const [index, bytes] of warm.entries()) {
      expect(bytes).toEqual(cold[index]);
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(entries[index].sha256);
    }
    console.log(
      `fidelity asset bench: cold ${coldRequests} requests in ${coldMs.toFixed(0)}ms, ` +
        `warm ${totalRequests - coldRequests} requests in ${warmMs.toFixed(0)}ms, ` +
        `${payloads.length} assets verified identical`
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
    await new Promise((done) => server.close(done));
  }
});
