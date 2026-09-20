import { execFile, spawn } from 'node:child_process';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { createServer } from 'node:net';
import { resolve } from 'node:path';
import { promisify } from 'node:util';
import { FORMATS, renderSection } from './readme.mjs';
import { measureSamples } from './results.mjs';
import { validateReferenceMetadata } from './reference.mjs';
import { download } from './download.mjs';
import { fetchAsset, resolveAssetCacheDir } from './asset-cache.mjs';
import { CORPUS_ORIGIN as corpus } from './samples.mjs';
import { preparePlan, selectPlanSamples, validatePlan } from './plan.mjs';

const execute = promisify(execFile);
const output = resolve(process.env.QUALITY_OUTPUT ?? '.source/office-quality/run');
const assetCache = resolveAssetCacheDir(process.env);
const python = process.env.QUALITY_PYTHON ?? 'python3';
if (
  (
    await command('git', [
      'status',
      '--porcelain',
      '--untracked-files=normal',
      '--',
      '.',
      ':!README.md',
    ])
  ).trim()
)
  throw new Error('Commit source changes before generating a revision-pinned report');
if ((await readdir(output).catch(() => [])).length)
  throw new Error('QUALITY_OUTPUT must be empty');
await mkdir(output, { recursive: true });

async function command(program, args, options = {}) {
  const result = await execute(program, args, {
    timeout: 660_000,
    maxBuffer: 16 * 1024 * 1024,
    ...options,
  });
  if (result.stderr) process.stderr.write(result.stderr);
  return result.stdout;
}

async function registry(name) {
  return JSON.parse(
    await download(
      `https://registry.npmjs.org/${encodeURIComponent(name)}/latest`,
      2 * 1024 * 1024
    )
  );
}

async function checkoutPlan() {
  const file = process.env.QUALITY_PLAN?.trim();
  const plan = file
    ? validatePlan(JSON.parse(await readFile(resolve(file), 'utf8')))
    : await preparePlan(process.env, { download, command, registry });
  const [sourceSha, commit] = await Promise.all([
    command('git', ['rev-parse', 'HEAD']),
    command('git', ['log', '-1', '--format=%H', '--', '.', ':!README.md']),
  ]);
  if (sourceSha.trim() !== plan.source_sha || commit.trim() !== plan.commit)
    throw new Error('Fidelity plan does not match this checkout');
  return plan;
}

async function asset(entry, destination, sample, maximum = 32 * 1024 * 1024) {
  const bytes = await fetchAsset(entry, sample, download, {
    cacheDir: assetCache,
    origin: corpus,
    maximum,
  });
  await writeFile(destination, bytes);
}

async function reference({ id, format, metadata }) {
  const metadataUrl = `${corpus}/${id}/metadata.json`;
  if (!FORMATS.includes(format) || metadata.format !== format)
    throw new Error(`Unsupported frozen sample format: ${id}`);
  validateReferenceMetadata(metadata, id);
  const directory = resolve(output, id);
  await mkdir(resolve(directory, 'reference'), { recursive: true });
  const source = resolve(directory, `source.${metadata.format}`);
  await asset(metadata.source, source, id, 128 * 1024 * 1024);
  for (const [index, page] of metadata.reference_pages.entries()) {
    await asset(
      page,
      resolve(directory, 'reference', `page_${String(index + 1).padStart(4, '0')}.png`),
      id
    );
  }
  await writeFile(
    resolve(directory, 'reference/result.json'),
    JSON.stringify(metadata.reference)
  );
  return {
    id,
    format,
    metadata_url: metadataUrl,
    source_sha256: metadata.source.sha256,
    source,
    directory,
    capture_profile: metadata.reference.capture_profile ?? null,
    comparisons: [],
  };
}

async function packageRoot(name, version) {
  const directory = resolve(output, 'packages', name.replace('@betteroffice/', ''));
  await mkdir(directory, { recursive: true });
  const result = JSON.parse(
    await command('npm', [
      'pack',
      `${name}@${version}`,
      '--ignore-scripts',
      '--json',
      '--pack-destination',
      directory,
    ])
  );
  const archive = result[0].filename;
  if (!/^[\w.-]+\.tgz$/.test(archive)) throw new Error('Invalid npm tarball filename');
  await command('tar', [
    '-xzf',
    resolve(directory, archive),
    '-C',
    directory,
    '--strip-components=1',
  ]);
  return directory;
}

async function viewer(overrides = {}) {
  const socket = createServer();
  await new Promise((done, fail) => {
    socket.once('error', fail);
    socket.listen(0, '127.0.0.1', done);
  });
  const port = socket.address().port;
  await new Promise((done) => socket.close(done));
  const env = { ...process.env, QUALITY_PORT: String(port) };
  delete env.QUALITY_PACKAGE_ROOT;
  delete env.QUALITY_REACT_ROOT;
  Object.assign(env, overrides);
  const child = spawn(process.execPath, ['scripts/docx-quality/server.mjs'], {
    env,
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  child.stderr.on('data', (data) => process.stderr.write(data));
  try {
    await new Promise((done, fail) => {
      const timer = setTimeout(() => fail(new Error('Viewer startup timed out')), 30_000);
      let text = '';
      child.stdout.on('data', (data) => {
        text += data;
        if (text.includes(`http://127.0.0.1:${port}/`)) {
          clearTimeout(timer);
          done();
        }
      });
      child.once('error', (error) => {
        clearTimeout(timer);
        fail(error);
      });
      child.once('exit', (code) => {
        clearTimeout(timer);
        fail(new Error(`Viewer exited: ${code}`));
      });
    });
    return { child, url: `http://127.0.0.1:${port}` };
  } catch (error) {
    child.kill();
    throw error;
  }
}

const plan = await checkoutPlan();
const requestedFormat = process.env.QUALITY_FORMAT?.trim() || null;
const { formats: selectedFormats, samples: selectedPlanSamples } = selectPlanSamples(
  plan,
  requestedFormat
);
const samples = [];
for (const sample of selectedPlanSamples) samples.push(await reference(sample));
for (const format of selectedFormats) {
  const reactVersion = format === 'docx' ? plan.react_version : null;
  const roots = {
    QUALITY_PACKAGE_ROOT: await packageRoot(
      `@betteroffice/${format}`,
      plan.versions[format]
    ),
    ...(reactVersion
      ? {
          QUALITY_REACT_ROOT: await packageRoot(
            '@betteroffice/docx-react',
            plan.react_version
          ),
        }
      : {}),
  };
  for (const channel of ['published', 'commit']) {
    const server = await viewer({
      QUALITY_FORMAT: format,
      ...(channel === 'published' ? roots : {}),
    });
    try {
      await measureSamples(
        samples.filter((sample) => sample.format === format),
        {
          channel,
          version: channel === 'published' ? plan.versions[format] : undefined,
          renderer_source_commit: channel === 'commit' ? plan.commit : undefined,
        },
        {
          capture: (sample) =>
            command(
              process.execPath,
              [
                'scripts/docx-quality/browser-task.mjs',
                sample.source,
                resolve(sample.directory, channel),
                'cdn',
                `${server.url}/?format=${format}`,
              ],
              {
                env: {
                  ...process.env,
                  QUALITY_CAPTURE_CONFIG: JSON.stringify(sample.capture_profile),
                  QUALITY_ENGINE_LABEL:
                    channel === 'published'
                      ? `@betteroffice/${format}@${plan.versions[format]}${
                          reactVersion
                            ? `; @betteroffice/docx-react@${plan.react_version}`
                            : ''
                        }`
                      : plan.commit,
                },
              }
            ),
          compare: async (sample) => {
            const difference = resolve(sample.directory, `${channel}-diff`);
            await command(python, [
              'scripts/office-quality/compare.py',
              resolve(sample.directory, 'reference'),
              resolve(sample.directory, channel),
              '--out',
              difference,
            ]);
            return JSON.parse(await readFile(resolve(difference, 'score.json'), 'utf8'));
          },
        }
      );
    } finally {
      server.child.kill();
    }
  }
}
const report = {
  source_sha: plan.source_sha,
  commit: plan.commit,
  versions: plan.versions,
  react_version: plan.react_version,
  ...(requestedFormat ? { format: requestedFormat } : {}),
  samples: samples.map(({ source, directory, ...sample }) => sample),
};
await writeFile(resolve(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
await writeFile(resolve(output, 'section.md'), renderSection(report));
console.log(`Generated ${resolve(output, 'section.md')}`);
