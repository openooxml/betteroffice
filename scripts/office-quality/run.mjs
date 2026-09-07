import { createHash } from 'node:crypto';
import { execFile, spawn } from 'node:child_process';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { createServer } from 'node:net';
import { resolve } from 'node:path';
import { promisify } from 'node:util';
import { FORMATS, renderSection } from './readme.mjs';
import { CORPUS_ORIGIN as corpus, selectSamples } from './samples.mjs';

const execute = promisify(execFile);
const output = resolve(process.env.QUALITY_OUTPUT ?? '.source/office-quality/run');
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
const ids = await selectSamples(process.env, download);

async function command(program, args, options = {}) {
  const result = await execute(program, args, {
    timeout: 660_000,
    maxBuffer: 16 * 1024 * 1024,
    ...options,
  });
  if (result.stderr) process.stderr.write(result.stderr);
  return result.stdout;
}

async function download(url, maximum = 32 * 1024 * 1024) {
  const response = await fetch(url, {
    signal: AbortSignal.timeout(30_000),
    credentials: 'omit',
    referrerPolicy: 'no-referrer',
  });
  if (!response.ok) throw new Error(`Download failed (${response.status}): ${url}`);
  const chunks = [];
  let size = 0;
  for await (const chunk of response.body) {
    size += chunk.length;
    if (size > maximum) throw new Error(`Download exceeds byte limit: ${url}`);
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function registry(name) {
  return JSON.parse(
    await download(
      `https://registry.npmjs.org/${encodeURIComponent(name)}/latest`,
      2 * 1024 * 1024
    )
  );
}

async function asset(entry, destination, sample) {
  const url = new URL(entry.url);
  if (
    url.origin !== corpus ||
    !url.pathname.startsWith(`/${sample}/`) ||
    !Number.isInteger(entry.bytes) ||
    entry.bytes < 1 ||
    entry.bytes > 32 * 1024 * 1024 ||
    !/^[a-f0-9]{64}$/.test(entry.sha256)
  )
    throw new Error('Invalid corpus asset metadata');
  const bytes = await download(url, entry.bytes);
  if (
    bytes.length !== entry.bytes ||
    createHash('sha256').update(bytes).digest('hex') !== entry.sha256
  ) {
    throw new Error(`Corpus hash mismatch: ${url}`);
  }
  await writeFile(destination, bytes);
}

async function reference(id) {
  const metadataUrl = `${corpus}/${id}/metadata.json`;
  const metadata = JSON.parse(await download(metadataUrl, 2 * 1024 * 1024));
  if (!FORMATS.includes(metadata.format))
    throw new Error(`Unsupported sample format: ${id}`);
  if (
    metadata.reference.status !== 'ok' ||
    metadata.reference.dpi !== 150 ||
    metadata.reference.sha256 !== metadata.source.sha256 ||
    metadata.reference.pages !== metadata.reference_pages.length ||
    metadata.reference.pages < 1 ||
    metadata.reference.pages > 100
  )
    throw new Error(`Invalid Office reference: ${id}`);
  const directory = resolve(output, id);
  await mkdir(resolve(directory, 'reference'), { recursive: true });
  const source = resolve(directory, `source.${metadata.format}`);
  await asset(metadata.source, source, id);
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
    format: metadata.format,
    metadata_url: metadataUrl,
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

const commit = (
  await command('git', ['log', '-1', '--format=%H', '--', '.', ':!README.md'])
).trim();
const versions = Object.fromEntries(
  await Promise.all(
    FORMATS.map(async (format) => [
      format,
      (await registry(`@betteroffice/${format}`)).version,
    ])
  )
);
const samples = [];
for (const id of ids) samples.push(await reference(id));
for (const format of FORMATS.filter((format) =>
  samples.some((sample) => sample.format === format)
)) {
  const reactVersion =
    format === 'docx' ? (await registry('@betteroffice/docx-react')).version : null;
  const roots = {
    QUALITY_PACKAGE_ROOT: await packageRoot(`@betteroffice/${format}`, versions[format]),
    ...(reactVersion
      ? {
          QUALITY_REACT_ROOT: await packageRoot('@betteroffice/docx-react', reactVersion),
        }
      : {}),
  };
  for (const channel of ['published', 'commit']) {
    const server = await viewer({
      QUALITY_FORMAT: format,
      ...(channel === 'published' ? roots : {}),
    });
    try {
      for (const sample of samples.filter((sample) => sample.format === format)) {
        const candidate = resolve(sample.directory, channel);
        const difference = resolve(sample.directory, `${channel}-diff`);
        await command(
          process.execPath,
          [
            'scripts/docx-quality/browser-task.mjs',
            sample.source,
            candidate,
            'cdn',
            `${server.url}/?format=${format}`,
          ],
          {
            env: {
              ...process.env,
              QUALITY_CAPTURE_CONFIG: JSON.stringify(sample.capture_profile),
              QUALITY_ENGINE_LABEL:
                channel === 'published'
                  ? `@betteroffice/${format}@${versions[format]}${
                      reactVersion ? `; @betteroffice/docx-react@${reactVersion}` : ''
                    }`
                  : commit,
            },
          }
        );
        await command(python, [
          'scripts/office-quality/compare.py',
          resolve(sample.directory, 'reference'),
          candidate,
          '--out',
          difference,
        ]);
        const comparison = JSON.parse(
          await readFile(resolve(difference, 'score.json'), 'utf8')
        );
        sample.comparisons.push({
          ...comparison,
          channel,
          version: channel === 'published' ? versions[format] : undefined,
          renderer_source_commit: channel === 'commit' ? commit : undefined,
        });
        console.log(`${sample.id} ${channel}: ${comparison.penalized_ssim.toFixed(4)}`);
      }
    } finally {
      server.child.kill();
    }
  }
}
const report = {
  commit,
  versions,
  samples: samples.map(({ source, directory, ...sample }) => sample),
};
await writeFile(resolve(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
await writeFile(resolve(output, 'section.md'), renderSection(report));
console.log(`Generated ${resolve(output, 'section.md')}`);
