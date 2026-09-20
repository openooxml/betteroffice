import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { promisify } from 'node:util';
import { RENDER_PREFIX, latestManifest, planRenders, reportKey } from './renders.mjs';

const execute = promisify(execFile);
const output = resolve(process.env.QUALITY_OUTPUT ?? '.source/office-quality/run');
const bucket = process.env.QUALITY_RENDER_BUCKET ?? '';
const concurrency = Number(process.env.QUALITY_RENDER_CONCURRENCY || 8);

if (!/^[a-z0-9][a-z0-9-]{1,62}$/.test(bucket))
  throw new Error('QUALITY_RENDER_BUCKET must name an R2 bucket');
if (!Number.isInteger(concurrency) || concurrency < 1)
  throw new Error('QUALITY_RENDER_CONCURRENCY must be a positive integer');

async function put(key, file, contentType) {
  if (!key.startsWith(`${RENDER_PREFIX}/`))
    throw new Error(`Refusing to write outside the prefix: ${key}`);
  await execute(
    'bunx',
    [
      '--no-install',
      'wrangler',
      'r2',
      'object',
      'put',
      `${bucket}/${key}`,
      '--file',
      file,
      '--content-type',
      contentType,
      '--force',
      '--remote',
    ],
    { timeout: 300_000, maxBuffer: 16 * 1024 * 1024, env: process.env }
  );
}

async function pooled(items, worker) {
  const queue = [...items];
  const workers = Array.from({ length: Math.min(concurrency, queue.length) }, async () => {
    for (let item = queue.shift(); item !== undefined; item = queue.shift()) await worker(item);
  });
  await Promise.all(workers);
}

const reportFile = resolve(output, 'report.json');
const plan = planRenders(JSON.parse(await readFile(reportFile, 'utf8')));
const sources = plan.uploads.map((upload) => ({
  ...upload,
  file: resolve(output, upload.sample, 'commit', upload.name),
}));
for (const source of sources) await stat(source.file);

const scratch = await mkdtemp(join(tmpdir(), 'fidelity-renders-'));
try {
  let done = 0;
  await pooled(sources, async (source) => {
    await put(source.key, source.file, 'image/png');
    done += 1;
    if (done % 50 === 0 || done === sources.length) console.log(`${done}/${sources.length} pages`);
  });
  await put(reportKey(plan.sha), reportFile, 'application/json');

  const latestFile = join(scratch, 'latest.json');
  const manifest = latestManifest(plan, new Date().toISOString());
  await writeFile(latestFile, JSON.stringify(manifest, null, 2) + '\n');
  await put(`${RENDER_PREFIX}/latest.json`, latestFile, 'application/json');
  console.log(
    `Published ${sources.length} pages for ${plan.sha} across ${Object.keys(plan.samples).length} samples`
  );
} finally {
  await rm(scratch, { recursive: true, force: true });
}
