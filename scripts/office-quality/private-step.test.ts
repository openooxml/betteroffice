import { expect, test } from 'bun:test';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createPlan } from './plan.mjs';
import { PRIVATE_BUCKET } from './private-corpus.mjs';
import { privateLogDirectory, privateMode } from './private-step.mjs';

const script = (name: string) => fileURLToPath(new URL(`./${name}`, import.meta.url));
const wrapper = script('private-step.mjs');

function run(args: string[], env: Record<string, string>) {
  return new Promise<{ code: number; output: string }>((done) => {
    execFile('node', [wrapper, ...args], { env: { ...process.env, QUALITY_CORPUS_ORIGIN: '', ...env } }, (error, stdout, stderr) =>
      done({ code: error ? Number((error as any).code ?? 1) : 0, output: stdout + stderr }));
  });
}

async function withTemp(task: (root: string) => Promise<void>) {
  const root = await mkdtemp(join(tmpdir(), 'private-step-test-'));
  try {
    await task(root);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

const ENDPOINT = `https://${'d'.repeat(32)}.r2.cloudflarestorage.com`;
const BUCKET = `${ENDPOINT}/${PRIVATE_BUCKET}`;
const signed: { url: string; authorization: string }[] = [];

async function withOrigin(routes: Record<string, (response: any) => void>, task: (origin: string) => Promise<void>) {
  const server = createServer((request, response) => {
    const path = request.url ?? '';
    const inBucket = path.startsWith(`/${PRIVATE_BUCKET}/`);
    if (inBucket) signed.push({ url: path, authorization: String(request.headers.authorization ?? '') });
    const route = routes[inBucket ? path.slice(PRIVATE_BUCKET.length + 1) : path];
    if (route) route(response);
    else response.writeHead(404).end();
  });
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done));
  try {
    await task(`http://127.0.0.1:${(server.address() as { port: number }).port}`);
  } finally {
    server.close();
  }
}

async function privateLogs(directory: string) {
  const names = await readdir(directory);
  return (await Promise.all(names.map((name) => readFile(join(directory, name), 'utf8')))).join('\n');
}

const leaks = (output: string, secrets: string[]) => secrets.filter((secret) => output.includes(secret));

async function privateBucket(root: string, origin: string) {
  const preload = join(root, 'route-r2.mjs');
  await writeFile(preload, `const real = globalThis.fetch;
globalThis.fetch = (input, init) => {
  const url = String(input);
  return real(url.startsWith(${JSON.stringify(ENDPOINT)}) ? ${JSON.stringify(origin)} + url.slice(${ENDPOINT.length}) : input, init);
};
`);
  signed.length = 0;
  return {
    QUALITY_PRIVATE: 'true',
    QUALITY_PRIVATE_LOGS: join(root, 'logs'),
    PRIVATE_R2_ENDPOINT: ENDPOINT,
    PRIVATE_R2_READ_KEY_ID: 'leakprobe-read-id',
    PRIVATE_R2_READ_SECRET: 'leakprobe-read-secret',
    NODE_OPTIONS: `--import=${pathToFileURL(preload).href}`,
  };
}

function expectSignedReads() {
  expect(signed.length).toBeGreaterThan(0);
  for (const { authorization } of signed)
    expect(authorization).toMatch(/^AWS4-HMAC-SHA256 Credential=leakprobe-read-id\/\d{8}\/auto\/s3\/aws4_request, /);
}

test('private mode is explicit and needs a log directory', () => {
  expect(privateMode({})).toBe(false);
  expect(privateMode({ QUALITY_PRIVATE: 'false' })).toBe(false);
  expect(privateMode({ QUALITY_PRIVATE: 'true' })).toBe(true);
  expect(() => privateMode({ QUALITY_PRIVATE: 'yes' })).toThrow('QUALITY_PRIVATE');
  expect(privateLogDirectory({ RUNNER_TEMP: '/runner/tmp' })).toBe('/runner/tmp/private-logs');
  expect(privateLogDirectory({ QUALITY_PRIVATE_LOGS: '/logs', RUNNER_TEMP: '/runner/tmp' })).toBe('/logs');
  expect(() => privateLogDirectory({})).toThrow('QUALITY_PRIVATE_LOGS');
});

test('a private step keeps output, errors and stack traces out of the job log and keeps the exit code', async () => {
  await withTemp(async (root) => {
    const child = ['-e', "console.log('leakprobe-sample-one'); console.error('https://private.example/leakprobe-sample-two/source.docx'); throw new Error('leakprobe-sample-three')"];
    const secrets = ['leakprobe-sample-one', 'leakprobe-sample-two', 'leakprobe-sample-three', 'private.example'];
    const shown = await run(['probe', '--', 'node', ...child], { QUALITY_PRIVATE: 'false' });
    expect(shown.code).toBe(1);
    expect(leaks(shown.output, secrets)).toEqual(secrets);
    const logs = join(root, 'logs');
    const hidden = await run(['probe', '--', 'node', ...child], { QUALITY_PRIVATE: 'true', QUALITY_PRIVATE_LOGS: logs });
    expect(hidden.code).toBe(1);
    expect(leaks(hidden.output, secrets)).toEqual([]);
    expect(hidden.output).toBe('probe: exit 1; output kept in the private log\n');
    expect(leaks(await readFile(join(logs, 'probe.log'), 'utf8'), secrets)).toEqual(secrets);
    const ok = await run(['probe-ok', '--', 'node', '-e', "console.log('leakprobe-sample-one')"], { QUALITY_PRIVATE: 'true', QUALITY_PRIVATE_LOGS: logs });
    expect(ok).toEqual({ code: 0, output: 'probe-ok: exit 0; output kept in the private log\n' });
  });
});

test('a private step rejects bad labels and missing commands without running anything', async () => {
  for (const args of [['../escape', '--', 'node', '-e', ''], ['Label', '--', 'node'], ['probe', 'node', '-e', ''], ['probe', '--']]) {
    const result = await run(args, { QUALITY_PRIVATE: 'true', QUALITY_PRIVATE_LOGS: tmpdir() });
    expect(result.code).toBe(2);
  }
  expect((await run(['probe', '--', 'node', '-e', ''], { QUALITY_PRIVATE: 'maybe' })).code).toBe(2);
});

test('a planning failure on a fixture corpus names its samples publicly but not in private mode', async () => {
  const ids = ['leakprobe-docx-a1', 'leakprobe-pptx-b2'];
  await withOrigin({
    '/collections/leakprobe-collection.json': (response) =>
      response.writeHead(200).end(JSON.stringify({ schema_version: 1, id: 'leakprobe-collection', samples: ids })),
    [`/${ids[0]}/metadata.json`]: (response) => response.writeHead(200).end(JSON.stringify({ format: 'odt' })),
    [`/${ids[1]}/metadata.json`]: (response) => response.writeHead(503, { 'Retry-After': '0' }).end(),
  }, async (origin) => {
    await withTemp(async (root) => {
      const secrets = [...ids, 'leakprobe-collection', origin.replace('http://', ''), 'metadata.json', 'd'.repeat(32), 'leakprobe-read'];
      const env = { QUALITY_COLLECTION: 'leakprobe-collection', QUALITY_OUTPUT: join(root, 'plan') };
      const plan = ['plan', '--', 'node', script('plan.mjs')];
      const shown = await run(plan, { ...env, QUALITY_CORPUS_ORIGIN: origin, QUALITY_PRIVATE: 'false' });
      expect(shown.code).not.toBe(0);
      expect(leaks(shown.output, secrets).length).toBeGreaterThan(0);
      const bucket = await privateBucket(root, origin);
      const hidden = await run(plan, { ...env, ...bucket });
      expect(hidden.code).not.toBe(0);
      expect(leaks(hidden.output, secrets)).toEqual([]);
      expect(leaks(await privateLogs(bucket.QUALITY_PRIVATE_LOGS), ids).length).toBeGreaterThan(0);
      expectSignedReads();
    });
  });
}, 30_000);

test('a staging failure on a fixture corpus names its sample and URL publicly but not in private mode', async () => {
  const id = 'leakprobe-docx-c3';
  const bytes = Buffer.from('synthetic private-mode fixture');
  await withOrigin({ [`/${id}/source.docx`]: (response) => response.writeHead(200).end(Buffer.from('tampered private-mode fixture!')) }, async (origin) => {
    await withTemp(async (root) => {
      const sha256 = createHash('sha256').update(bytes).digest('hex');
      const planAt = async (base: string) => {
        const fixture = createPlan({ samples: [{ id, format: 'docx', metadata: { format: 'docx',
          source: { url: `${base}/${id}/source.docx`, bytes: bytes.length, sha256 },
          reference: { status: 'ok', dpi: 150, sha256, pages: 1 },
          reference_pages: [{ url: `${base}/${id}/reference/page_0001.png`, bytes: 1, sha256 }] } }],
          source_sha: 'b'.repeat(40), commit: 'b'.repeat(40), react_version: '1.0.0', versions: { docx: '1.0.0', pptx: '1.0.0', xlsx: '1.0.0' },
          docx_published_source_sha: 'c'.repeat(40), formats: ['docx'] });
        const path = join(root, `plan-${base === origin ? 'public' : 'private'}.json`);
        await writeFile(path, JSON.stringify(fixture));
        return path;
      };
      const secrets = [id, origin.replace('http://', ''), 'source.docx', 'd'.repeat(32), 'leakprobe-read'];
      const env = { QUALITY_FORMAT: 'docx', QUALITY_SHARD: '0', QUALITY_OUTPUT: join(root, 'output') };
      const stage = ['roundtrip-stage', '--', 'node', script('roundtrip.mjs')];
      const shown = await run(stage, { ...env, QUALITY_CORPUS_ORIGIN: origin, QUALITY_PLAN: await planAt(origin), QUALITY_PRIVATE: 'false' });
      expect(shown.code).not.toBe(0);
      expect(leaks(shown.output, secrets).length).toBeGreaterThan(0);
      const bucket = await privateBucket(root, origin);
      const hidden = await run(stage, { ...env, ...bucket, QUALITY_PLAN: await planAt(BUCKET) });
      expect(hidden.code).not.toBe(0);
      expect(leaks(hidden.output, secrets)).toEqual([]);
      const logs = await privateLogs(bucket.QUALITY_PRIVATE_LOGS);
      expect(leaks(logs, [id]).length).toBeGreaterThan(0);
      expect(logs).toContain('Corpus hash mismatch');
      expectSignedReads();
    });
  });
}, 30_000);
