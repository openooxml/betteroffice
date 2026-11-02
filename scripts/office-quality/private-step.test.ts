import { expect, test } from 'bun:test';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createPlan } from './plan.mjs';
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

async function withOrigin(routes: Record<string, (response: any) => void>, task: (origin: string) => Promise<void>) {
  const server = createServer((request, response) => {
    const route = routes[request.url ?? ''];
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
      const secrets = [...ids, 'leakprobe-collection', origin.replace('http://', ''), 'metadata.json'];
      const env = { QUALITY_COLLECTION: 'leakprobe-collection', QUALITY_CORPUS_ORIGIN: origin, QUALITY_OUTPUT: join(root, 'plan') };
      const plan = ['plan', '--', 'node', script('plan.mjs')];
      const shown = await run(plan, { ...env, QUALITY_PRIVATE: 'false' });
      expect(shown.code).not.toBe(0);
      expect(leaks(shown.output, secrets).length).toBeGreaterThan(0);
      const logs = join(root, 'logs');
      const hidden = await run(plan, { ...env, QUALITY_PRIVATE: 'true', QUALITY_PRIVATE_LOGS: logs });
      expect(hidden.code).not.toBe(0);
      expect(leaks(hidden.output, secrets)).toEqual([]);
      expect(leaks(await privateLogs(logs), ids).length).toBeGreaterThan(0);
    });
  });
}, 30_000);

test('a staging failure on a fixture corpus names its sample and URL publicly but not in private mode', async () => {
  const id = 'leakprobe-docx-c3';
  const bytes = Buffer.from('synthetic private-mode fixture');
  await withOrigin({ [`/${id}/source.docx`]: (response) => response.writeHead(200).end(Buffer.from('tampered private-mode fixture!')) }, async (origin) => {
    await withTemp(async (root) => {
      const sha256 = createHash('sha256').update(bytes).digest('hex');
      const source = { url: `${origin}/${id}/source.docx`, bytes: bytes.length, sha256 };
      const formats = ['docx'];
      const fixture = createPlan({ samples: [{ id, format: 'docx', metadata: { format: 'docx', source,
        reference: { status: 'ok', dpi: 150, sha256, pages: 1 },
        reference_pages: [{ url: `${origin}/${id}/reference/page_0001.png`, bytes: 1, sha256 }] } }],
        source_sha: 'b'.repeat(40), commit: 'b'.repeat(40), react_version: '1.0.0', versions: { docx: '1.0.0', pptx: '1.0.0', xlsx: '1.0.0' },
        docx_published_source_sha: 'c'.repeat(40), formats });
      const planPath = join(root, 'plan.json');
      await writeFile(planPath, JSON.stringify(fixture));
      const secrets = [id, origin.replace('http://', ''), 'source.docx'];
      const env = { QUALITY_CORPUS_ORIGIN: origin, QUALITY_PLAN: planPath, QUALITY_FORMAT: 'docx', QUALITY_SHARD: '0', QUALITY_OUTPUT: join(root, 'output') };
      const stage = ['roundtrip-stage', '--', 'node', script('roundtrip.mjs')];
      const shown = await run(stage, { ...env, QUALITY_PRIVATE: 'false' });
      expect(shown.code).not.toBe(0);
      expect(leaks(shown.output, secrets).length).toBeGreaterThan(0);
      const logs = join(root, 'logs');
      const hidden = await run(stage, { ...env, QUALITY_PRIVATE: 'true', QUALITY_PRIVATE_LOGS: logs });
      expect(hidden.code).not.toBe(0);
      expect(leaks(hidden.output, secrets)).toEqual([]);
      expect(leaks(await privateLogs(logs), [id]).length).toBeGreaterThan(0);
    });
  });
}, 30_000);
