import { expect, test } from 'bun:test';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PRIVATE_BUCKET } from './private-corpus.mjs';
import { bundleKey, bundlePaths, createStore, parseListing } from './private-store.mjs';
import { sha256Hex } from './sigv4.mjs';

const ENDPOINT = `https://${'c'.repeat(32)}.r2.cloudflarestorage.com`;
const BASE = `${ENDPOINT}/${PRIVATE_BUCKET}`;
const ENV = {
  GITHUB_RUN_ID: '424242',
  PRIVATE_R2_ENDPOINT: ENDPOINT,
  PRIVATE_R2_READ_KEY_ID: 'read-id',
  PRIVATE_R2_READ_SECRET: 'read-secret',
  PRIVATE_R2_WRITE_KEY_ID: 'write-id',
  PRIVATE_R2_WRITE_SECRET: 'write-secret',
};

function bucket({ pageSize = 1000, transient = 0 } = {}) {
  const objects = new Map<string, { body: Buffer; meta: string }>();
  const calls: { method: string; url: string; credential: string }[] = [];
  let failures = transient;
  const request = async (url: string, init: any) => {
    const credential = init.headers.authorization.match(/Credential=([^/]+)\//)[1];
    calls.push({ method: init.method, url, credential });
    expect(init.redirect).toBe('error');
    expect(url.startsWith(`${BASE}/`) || url.startsWith(`${BASE}?`)).toBe(true);
    if (failures-- > 0) return new Response('busy', { status: 503 });
    const target = new URL(url);
    if (init.method === 'PUT') {
      expect(credential).toBe('write-id');
      const body = Buffer.from(init.body);
      expect(init.headers['x-amz-content-sha256']).toBe(sha256Hex(body));
      objects.set(decodeURIComponent(target.pathname.slice(PRIVATE_BUCKET.length + 2)), { body, meta: init.headers['x-amz-meta-sha256'] });
      return new Response(null, { status: 200 });
    }
    expect(credential).toBe('read-id');
    if (target.searchParams.get('list-type') === '2') {
      const prefix = target.searchParams.get('prefix') ?? '';
      const keys = [...objects.keys()].filter((key) => key.startsWith(prefix)).sort();
      const start = Number(target.searchParams.get('continuation-token') ?? 0);
      const page = keys.slice(start, start + pageSize);
      const more = start + pageSize < keys.length;
      return new Response(`<?xml version="1.0"?><ListBucketResult>${page.map((key) => `<Contents><Key>${key.replace(/&/g, '&amp;')}</Key></Contents>`).join('')}` +
        `<IsTruncated>${more}</IsTruncated>${more ? `<NextContinuationToken>${start + pageSize}</NextContinuationToken>` : ''}</ListBucketResult>`);
    }
    const object = objects.get(decodeURIComponent(target.pathname.slice(PRIVATE_BUCKET.length + 2)));
    if (!object) return new Response('missing', { status: 404 });
    return new Response(object.body, { headers: { 'x-amz-meta-sha256': object.meta } });
  };
  return { objects, calls, request };
}

async function withTemp(task: (root: string) => Promise<void>) {
  const root = await mkdtemp(join(tmpdir(), 'private-store-test-'));
  try {
    await task(root);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

const noSleep = async () => {};

test('bundles are keyed by run and name and stay inside the working directory', () => {
  expect(bundleKey(ENV, 'docx-benchmark-report-3')).toBe('runs/424242/docx-benchmark-report-3.tar.gz');
  expect(() => bundleKey(ENV, '../x')).toThrow('Invalid private bundle name');
  expect(() => bundleKey(ENV, 'Report')).toThrow('Invalid private bundle name');
  expect(() => bundleKey({}, 'plan')).toThrow('GITHUB_RUN_ID');
  expect(bundlePaths(['report.json', 'a/timing/', 'a/timing/x.json', './report.json', 'b/lo'])).toEqual(['a/timing', 'b/lo', 'report.json']);
  for (const bad of [['/etc/passwd'], ['../x'], ['a/../../x'], ['.']]) expect(() => bundlePaths(bad)).toThrow('inside the working directory');
});

test('a bundle round-trips with the write key up and the read key down', async () => {
  await withTemp(async (root) => {
    const { calls, request, objects } = bucket({ transient: 1 });
    const store = createStore(ENV, { request, sleep: noSleep });
    await mkdir(join(root, 'src/s1/timing'), { recursive: true });
    await writeFile(join(root, 'src/report.json'), '{"ok":true}');
    await writeFile(join(root, 'src/s1/timing/t.json'), '{"ms":1}');
    expect(await store.put('xlsx-benchmark-report-0', ['report.json', 's1/timing'], { cwd: join(root, 'src') }))
      .toEqual({ name: 'xlsx-benchmark-report-0', bytes: expect.any(Number), skipped: false });
    expect([...objects.keys()]).toEqual(['runs/424242/xlsx-benchmark-report-0.tar.gz']);
    expect(calls.map((call) => call.method)).toEqual(['PUT', 'PUT']);
    await store.get('xlsx-benchmark-report-0', join(root, 'out'));
    expect(await readFile(join(root, 'out/report.json'), 'utf8')).toBe('{"ok":true}');
    expect(await readFile(join(root, 'out/s1/timing/t.json'), 'utf8')).toBe('{"ms":1}');
  });
});

test('a pattern fetches every matching bundle into its own folder, across listing pages', async () => {
  await withTemp(async (root) => {
    const { request, objects } = bucket({ pageSize: 2 });
    const store = createStore(ENV, { request, sleep: noSleep });
    for (const shard of [0, 1, 2]) {
      await mkdir(join(root, `s${shard}`), { recursive: true });
      await writeFile(join(root, `s${shard}/report.json`), `{"shard":${shard}}`);
      await store.put(`docx-benchmark-report-${shard}`, ['report.json'], { cwd: join(root, `s${shard}`) });
    }
    await store.put('docx-benchmark-diagnostics-0', ['report.json'], { cwd: join(root, 's0') });
    objects.set('runs/999/docx-benchmark-report-9.tar.gz', objects.get('runs/424242/docx-benchmark-report-0.tar.gz')!);
    const fetched = await store.get('docx-benchmark-report-*', join(root, 'parts'));
    expect(fetched.map((item) => item.name)).toEqual(['docx-benchmark-report-0', 'docx-benchmark-report-1', 'docx-benchmark-report-2']);
    for (const shard of [0, 1, 2])
      expect(await readFile(join(root, `parts/docx-benchmark-report-${shard}/report.json`), 'utf8')).toBe(`{"shard":${shard}}`);
    await expect(store.get('roundtrip-report-*', join(root, 'none'))).rejects.toThrow('No private bundles match roundtrip-report-*');
  });
});

test('missing paths, missing bundles and tampered bundles fail', async () => {
  await withTemp(async (root) => {
    const { request, objects } = bucket();
    const store = createStore(ENV, { request, sleep: noSleep });
    await expect(store.put('plan', ['plan.json'], { cwd: root })).rejects.toThrow('Missing path for private bundle plan');
    expect(await store.put('logs-measure-docx', ['private-logs'], { cwd: root, optional: true }))
      .toEqual({ name: 'logs-measure-docx', bytes: 0, skipped: true });
    expect(objects.size).toBe(0);
    await expect(store.get('plan', join(root, 'out'))).rejects.toThrow('Private bundle runs/424242/plan.tar.gz is unavailable (HTTP 404)');
    await writeFile(join(root, 'plan.json'), '{}');
    await store.put('plan', ['plan.json'], { cwd: root });
    const object = objects.get('runs/424242/plan.tar.gz')!;
    object.body = Buffer.concat([object.body, Buffer.from('x')]);
    await expect(store.get('plan', join(root, 'out'))).rejects.toThrow('failed its integrity check');
  });
});

test('the store fails closed without the matching key pair', async () => {
  const { request } = bucket();
  const readOnly = { ...ENV, PRIVATE_R2_WRITE_KEY_ID: '', PRIVATE_R2_WRITE_SECRET: '' };
  await withTemp(async (root) => {
    await writeFile(join(root, 'plan.json'), '{}');
    await expect(createStore(readOnly, { request, sleep: noSleep }).put('plan', ['plan.json'], { cwd: root }))
      .rejects.toThrow('PRIVATE_R2_WRITE_KEY_ID and PRIVATE_R2_WRITE_SECRET');
  });
  await expect(createStore({ ...ENV, PRIVATE_R2_ENDPOINT: '' }, { request, sleep: noSleep }).get('plan', '/nonexistent'))
    .rejects.toThrow('Private corpus runs need PRIVATE_R2_ENDPOINT');
});

test('listings decode keys and continuation tokens', () => {
  expect(parseListing('<ListBucketResult><Contents><Key>runs/1/a&amp;b.tar.gz</Key></Contents><IsTruncated>true</IsTruncated><NextContinuationToken>t+/=&amp;</NextContinuationToken></ListBucketResult>'))
    .toEqual({ keys: ['runs/1/a&b.tar.gz'], next: 't+/=&' });
  expect(parseListing('<ListBucketResult><IsTruncated>false</IsTruncated></ListBucketResult>')).toEqual({ keys: [], next: '' });
});

test('archive failures name the bundle and nothing inside it', async () => {
  await withTemp(async (root) => {
    const { request } = bucket();
    const leaky = async (args: string[]) => {
      throw new Error(`Command failed: tar ${args.join(' ')}\ntar: leakprobe-sample-x/timing/t.json: No space left on device`);
    };
    await mkdir(join(root, 'leakprobe-sample-x/timing'), { recursive: true });
    await writeFile(join(root, 'leakprobe-sample-x/timing/t.json'), '{}');
    const store = createStore(ENV, { request, sleep: noSleep, tar: leaky });
    const failure = await store.put('docx-benchmark-diagnostics-0', ['leakprobe-sample-x/timing'], { cwd: root }).catch((error) => error);
    expect(failure.message).toBe('Packing private bundle docx-benchmark-diagnostics-0 failed');
    await writeFile(join(root, 'plan.json'), '{}');
    await createStore(ENV, { request, sleep: noSleep }).put('plan', ['plan.json'], { cwd: root });
    const unpack = await store.get('plan', join(root, 'leakprobe-sample-x/out')).catch((error) => error);
    expect(unpack.message).toBe('Unpacking private bundle plan failed');
  });
});
