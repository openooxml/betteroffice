import { open, mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { createServer } from 'node:net';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { killGroup, spawnGroup } from './process-group.mjs';
import { MAX_LOCAL_REFERENCE_PAGES } from './reference.mjs';

const { values } = parseArgs({ options: Object.fromEntries([
  'source', 'reference-pdf', 'reference-dir', 'package-root', 'react-root',
  'fonts-dist', 'out', 'python', 'jobs', 'cache', 'baseline', 'timeout-seconds',
].map((key) => [key, { type: 'string' }])) });
for (const key of ['source', 'reference-pdf', 'reference-dir', 'package-root', 'react-root', 'fonts-dist', 'out']) {
  if (!values[key]) throw new Error(`--${key} is required`);
  values[key] = resolve(values[key]);
}
for (const key of ['cache', 'baseline']) if (values[key]) values[key] = resolve(values[key]);
const jobs = Number(values.jobs ?? 4);
const seconds = Number(values['timeout-seconds'] ?? 3600);
if (!Number.isInteger(jobs) || jobs < 1 || !Number.isFinite(seconds) || seconds <= 0)
  throw new Error('jobs must be a positive integer and timeout-seconds must be positive');
const repo = fileURLToPath(new URL('../..', import.meta.url));
const script = (name) => resolve(repo, 'scripts/office-quality', name);
const python = values.python ?? 'python3';
const node = process.execPath;
const out = values.out;
await mkdir(out, { recursive: true });
if ((await readdir(out)).length) throw new Error('--out must be empty');
const children = new Set();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function start(command, args, name, env = {}, capture = false, timeout = null) {
  const log = await open(resolve(out, `${name}.log`), 'w');
  const child = spawnGroup(command, args, {
    cwd: repo,
    env: { ...process.env, ...env },
    stdio: ['ignore', capture ? 'pipe' : log.fd, log.fd],
  });
  children.add(child);
  let stdout = '';
  const childLogWrites = [];
  if (capture) child.stdout.on('data', (data) => {
    stdout += data.toString();
    childLogWrites.push(log.write(data));
  });
  let failure;
  child.on('error', (error) => { failure = error; });
  let timedOut = false;
  let forced;
  const timer = timeout === null ? null : setTimeout(() => {
    timedOut = true;
    killGroup(child, 'SIGTERM');
    forced = setTimeout(() => killGroup(child, 'SIGKILL'), 2000);
  }, timeout * 1000);
  const done = new Promise((resolve) => child.on('close', async (code, signal) => {
    clearTimeout(timer);
    clearTimeout(forced);
    children.delete(child);
    await Promise.all(childLogWrites);
    await log.close();
    resolve({ code, signal, stdout, failure, timedOut });
  }));
  return { child, done };
}

async function run(command, args, name, env = {}, capture = false) {
  const task = await start(command, args, name, env, capture, seconds);
  const result = await task.done;
  if (result.failure || result.timedOut || result.code !== 0)
    throw new Error(`${name} failed${result.timedOut ? ' (timeout)' : ''}; see ${resolve(out, `${name}.log`)}`);
  return result.stdout;
}

async function freePort() {
  const server = createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const port = server.address().port;
  await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  return port;
}

async function ready(task, url) {
  let exited = false;
  task.done.then(() => { exited = true; });
  const deadline = Date.now() + Math.min(seconds * 1000, 60_000);
  while (!exited && Date.now() < deadline) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(1000) });
      await response.body?.cancel();
      if (response.ok) return;
    } catch {}
    await sleep(100);
  }
  throw new Error(`Capture server did not become ready; see ${resolve(out, 'server.log')}`);
}

async function stop(task) {
  killGroup(task.child, 'SIGTERM');
  const forced = setTimeout(() => killGroup(task.child, 'SIGKILL'), 2000);
  await task.done;
  clearTimeout(forced);
}

for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => {
  for (const child of children) killGroup(child, signal);
  process.exitCode = signal === 'SIGINT' ? 130 : 143;
});

try {
  await run(python, [script('private_ssim.py'), 'rasterize', values['reference-pdf'], '--out', values['reference-dir']], 'rasterize');
  const profile = JSON.parse(await run(python, [script('private_ssim.py'), 'profile', values['reference-dir']], 'profile', {}, true));
  if (!profile.pages.length || profile.pages.length > MAX_LOCAL_REFERENCE_PAGES)
    throw new Error(`Local references must have 1–${MAX_LOCAL_REFERENCE_PAGES} pages`);
  const port = await freePort();
  const url = `http://127.0.0.1:${port}/?local=1&maxPages=${profile.pages.length}`;
  const env = {
    QUALITY_FORMAT: 'docx',
    QUALITY_PACKAGE_ROOT: values['package-root'],
    QUALITY_REACT_ROOT: values['react-root'],
    QUALITY_FONT_ASSETS: resolve(repo, 'packages/fonts/assets'),
    QUALITY_FONT_ASSETS_CJK: resolve(repo, 'packages/fonts-cjk/assets'),
    QUALITY_PORT: String(port),
    QUALITY_LOCAL_ONLY: '1',
    QUALITY_TIMEOUT_SECONDS: String(seconds),
    QUALITY_ENGINE_LABEL: values['package-root'],
  };
  const server = await start(node, [resolve(repo, 'scripts/docx-quality/server.mjs')], 'server', env);
  const actual = resolve(out, 'betteroffice');
  try {
    await ready(server, url);
    const config = JSON.stringify(profile);
    const configFile = resolve(out, 'capture-profile.json');
    await writeFile(configFile, config + '\n');
    const captureEnv = {
      ...env,
      QUALITY_CAPTURE_CONFIG: Buffer.byteLength(config) <= 100_000 ? config : '',
      QUALITY_CAPTURE_CONFIG_FILE: Buffer.byteLength(config) > 100_000 ? configFile : '',
    };
    await run(node, [resolve(repo, 'scripts/docx-quality/browser-task.mjs'), values.source, actual, 'cdn', url], 'capture', captureEnv);
  } finally {
    await stop(server);
  }
  await run(node, [script('private-layout-text.mjs'), values['package-root'], values['fonts-dist'], values.source, resolve(actual, 'pages_text.json')], 'layout-text');
  const capture = JSON.parse(await readFile(resolve(actual, 'result.json'), 'utf8'));
  const text = JSON.parse(await readFile(resolve(actual, 'pages_text.json'), 'utf8'));
  const summary = { reference_pages: profile.pages.length, capture_pages: capture.pages, layout_pages: text.length };
  await writeFile(resolve(out, 'summary.json'), JSON.stringify(summary, null, 2) + '\n');
  if (capture.status !== 'ok' || capture.pages !== text.length)
    throw new Error(`Capture/layout page counts differ (${capture.pages} vs ${text.length}); see ${resolve(out, 'summary.json')}`);
  const args = [script('private_ssim.py'), 'compare', values['reference-dir'], actual, '--out', resolve(out, 'compare'), '--jobs', String(jobs)];
  for (const key of ['cache', 'baseline']) if (values[key]) args.push(`--${key}`, values[key]);
  const comparison = await start(python, args, 'compare', {}, true, seconds);
  const result = await comparison.done;
  if (result.failure || result.timedOut)
    throw new Error(`Compare failed; see ${resolve(out, 'compare.log')}`);
  if (result.stdout.trim()) {
    const score = JSON.parse(result.stdout.trim());
    console.log(JSON.stringify(score));
    if (result.code === 1 && score.regression === true) process.exitCode = 1;
    else if (result.code !== 0)
      throw new Error(`Compare failed; see ${resolve(out, 'compare.log')}`);
  } else {
    throw new Error(`Compare failed; see ${resolve(out, 'compare.log')}`);
  }
} catch (error) {
  console.error(error.message);
  process.exitCode ||= 1;
}
