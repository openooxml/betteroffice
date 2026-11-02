import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, normalize, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import { backoffMs, isTransientStatus } from './download.mjs';
import { privateCorpus, signedHeaders } from './private-corpus.mjs';
import { sha256Hex, uriEncode } from './sigv4.mjs';

const execute = promisify(execFile);
const NAME = /^[a-z0-9][a-z0-9-]{0,127}$/;
export const MAX_BUNDLE_BYTES = 2 * 1024 ** 3;
const ATTEMPTS = 5;

export function runPrefix(environment = process.env) {
  const run = environment.GITHUB_RUN_ID?.trim() ?? '';
  if (!/^[1-9][0-9]{0,19}$/.test(run)) throw new Error('GITHUB_RUN_ID is required');
  return `runs/${run}/`;
}

export function bundleKey(environment, name) {
  if (!NAME.test(name ?? '')) throw new Error('Invalid private bundle name');
  return `${runPrefix(environment)}${name}.tar.gz`;
}

export function objectUrl(corpus, key) {
  return `${corpus.base}/${key.split('/').map((part) => uriEncode(part)).join('/')}`;
}

export function bundlePaths(paths) {
  const unique = [...new Set(paths.map((path) => normalize(path).replace(/\/+$/, '')))];
  for (const path of unique)
    if (!path || path === '.' || path.startsWith('/') || path === '..' || path.startsWith('../'))
      throw new Error('Bundle paths must stay inside the working directory');
  return unique.filter((path) => !unique.some((other) => other !== path && path.startsWith(`${other}/`))).sort();
}

async function send(corpus, method, url, options, request, sleep) {
  for (let attempt = 0; ; attempt++) {
    let response;
    try {
      response = await request(url, {
        method,
        headers: signedHeaders(corpus, method, url, options),
        body: options.body,
        redirect: 'error',
        signal: AbortSignal.timeout(15 * 60_000),
      });
    } catch (error) {
      if (attempt + 1 >= ATTEMPTS) throw new Error(`Private bucket ${method} failed: ${error.message}`);
      await sleep(backoffMs(attempt));
      continue;
    }
    if (response.ok || !isTransientStatus(response.status) || attempt + 1 >= ATTEMPTS) return response;
    await response.body?.cancel?.();
    await sleep(backoffMs(attempt));
  }
}

export function parseListing(xml) {
  const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  const text = (value) => value.replace(/&(amp|lt|gt|quot|apos);/g, (_, entity) => entities[entity]);
  return {
    keys: [...xml.matchAll(/<Key>([^<]*)<\/Key>/g)].map((match) => text(match[1])),
    next: /<IsTruncated>true<\/IsTruncated>/.test(xml)
      ? text(xml.match(/<NextContinuationToken>([^<]*)<\/NextContinuationToken>/)?.[1] ?? '')
      : '',
  };
}

async function runTar(args) {
  await execute('tar', args, { maxBuffer: 16 * 1024 * 1024 });
}

export function createStore(environment = process.env, {
  request = fetch,
  sleep = (ms) => new Promise((done) => setTimeout(done, ms)),
  tar = runTar,
} = {}) {
  async function archive(args, failure) {
    try {
      await tar(args);
    } catch {
      throw new Error(failure);
    }
  }

  const reader = () => privateCorpus(environment, 'read');
  const writer = () => privateCorpus(environment, 'write');

  async function list(prefix) {
    const corpus = reader();
    const keys = [];
    let token = '';
    do {
      const query = `list-type=2&prefix=${uriEncode(prefix)}${token ? `&continuation-token=${uriEncode(token)}` : ''}`;
      const response = await send(corpus, 'GET', `${corpus.base}?${query}`, {}, request, sleep);
      if (!response.ok) throw new Error(`Private bucket listing failed (HTTP ${response.status})`);
      const page = parseListing(await response.text());
      keys.push(...page.keys);
      if (page.next && page.next === token) throw new Error('Private bucket listing did not advance');
      token = page.next;
    } while (token);
    return keys;
  }

  async function fetchBundle(key, file) {
    const corpus = reader();
    const response = await send(corpus, 'GET', objectUrl(corpus, key), {}, request, sleep);
    if (!response.ok) throw new Error(`Private bundle ${key} is unavailable (HTTP ${response.status})`);
    const bytes = Buffer.from(await response.arrayBuffer());
    const expected = response.headers.get('x-amz-meta-sha256');
    if (!expected || sha256Hex(bytes) !== expected) throw new Error(`Private bundle ${key} failed its integrity check`);
    await writeFile(file, bytes);
    return bytes.length;
  }

  return {
    async put(name, paths, { cwd = process.cwd(), optional = false } = {}) {
      const key = bundleKey(environment, name);
      const present = [];
      for (const path of bundlePaths(paths)) {
        if (await stat(resolve(cwd, path)).then(() => true, () => false)) present.push(path);
        else if (!optional) throw new Error(`Missing path for private bundle ${name}`);
      }
      if (!present.length) {
        if (optional) return { name, bytes: 0, skipped: true };
        throw new Error(`Nothing to keep for private bundle ${name}`);
      }
      const corpus = writer();
      const scratch = await mkdtemp(join(tmpdir(), 'private-store-'));
      try {
        const bundle = join(scratch, 'bundle.tar.gz');
        await archive(['-czf', bundle, '-C', cwd, '--', ...present], `Packing private bundle ${name} failed`);
        if ((await stat(bundle)).size > MAX_BUNDLE_BYTES) throw new Error(`Private bundle ${name} exceeds ${MAX_BUNDLE_BYTES} bytes`);
        const body = await readFile(bundle);
        const digest = sha256Hex(body);
        const response = await send(corpus, 'PUT', objectUrl(corpus, key), {
          body,
          payloadHash: digest,
          headers: { 'content-type': 'application/gzip', 'x-amz-meta-sha256': digest },
        }, request, sleep);
        if (!response.ok) throw new Error(`Private bundle ${name} upload failed (HTTP ${response.status})`);
        await response.body?.cancel?.();
        return { name, bytes: body.length, skipped: false };
      } finally {
        await rm(scratch, { recursive: true, force: true });
      }
    },

    async get(name, destination) {
      const pattern = name?.endsWith('*');
      const stem = pattern ? name.slice(0, -1) : name;
      if (!NAME.test(stem ?? '')) throw new Error('Invalid private bundle name');
      const prefix = runPrefix(environment);
      const names = pattern
        ? (await list(`${prefix}${stem}`)).filter((key) => key.endsWith('.tar.gz'))
            .map((key) => key.slice(prefix.length, -'.tar.gz'.length)).filter((found) => NAME.test(found)).sort()
        : [stem];
      if (!names.length) throw new Error(`No private bundles match ${name}`);
      const scratch = await mkdtemp(join(tmpdir(), 'private-store-'));
      const fetched = [];
      try {
        for (const found of names) {
          const bundle = join(scratch, `${found}.tar.gz`);
          const bytes = await fetchBundle(bundleKey(environment, found), bundle);
          const target = pattern ? resolve(destination, found) : resolve(destination);
          await mkdir(target, { recursive: true });
          await archive(['-xzf', bundle, '-C', target, '--no-same-owner'], `Unpacking private bundle ${found} failed`);
          fetched.push({ name: found, bytes });
        }
      } finally {
        await rm(scratch, { recursive: true, force: true });
      }
      return fetched;
    },
  };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [mode, ...rest] = process.argv.slice(2);
  const optional = rest[0] === '--optional';
  const [name, ...paths] = optional ? rest.slice(1) : rest;
  try {
    const store = createStore();
    if (mode === 'put' && (paths.length || optional)) {
      const result = await store.put(name, paths, { optional });
      process.stdout.write(result.skipped ? `${name}: nothing to keep\n` : `${name}: kept ${result.bytes} bytes\n`);
    } else if (mode === 'get' && paths.length === 1 && !optional) {
      for (const result of await store.get(name, paths[0])) process.stdout.write(`${result.name}: fetched ${result.bytes} bytes\n`);
    } else {
      process.stderr.write('usage: private-store.mjs put [--optional] <name> <path...> | get <name[*]> <destination>\n');
      process.exitCode = 2;
    }
  } catch (error) {
    process.stderr.write(`private-store: ${error.message}\n`);
    process.exitCode = 1;
  }
}
