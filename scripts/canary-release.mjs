import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, renameSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { NATIVE_PACKAGES } from './native-bindings.mjs';
import { workspaceEntries } from './published-packages.mjs';

export const CANARY_TAG = 'canary';
const ROOT = fileURLToPath(new URL('..', import.meta.url));
const FIELDS = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'];
const RUNTIME_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'];
const PLAN = 'plan.json';

export function isCanaryVersion(version) {
  return typeof version === 'string' && new RegExp(`^\\d+\\.\\d+\\.\\d+-${CANARY_TAG}\\.\\d+$`).test(version);
}

function internalDependencies(manifest, names) {
  return RUNTIME_FIELDS.flatMap((field) => Object.keys(manifest[field] ?? {}).filter((name) => names.has(name)));
}

/** Packages that need a native package to resolve, directly or through another workspace package. */
export function nativeDependents(workspaces, nativeNames) {
  const byName = new Map(workspaces.map(({ manifest }) => [manifest.name, manifest]));
  const names = new Set(byName.keys());
  const needs = new Map();
  const visit = (name, seen = new Set()) => {
    if (needs.has(name)) return needs.get(name);
    if (nativeNames.has(name)) return true;
    if (seen.has(name)) return false;
    seen.add(name);
    const result = internalDependencies(byName.get(name), names).some((dependency) => visit(dependency, seen));
    needs.set(name, result);
    return result;
  };
  return new Set([...names].filter((name) => !nativeNames.has(name) && visit(name)));
}

/** The canary set, dependencies before dependents: public, versioned as a canary, and installable without a native package. */
export function canarySet(workspaces, nativeNames) {
  const dependents = nativeDependents(workspaces, nativeNames);
  const chosen = workspaces.filter(({ manifest }) =>
    manifest.name && !manifest.private && isCanaryVersion(manifest.version) &&
    !nativeNames.has(manifest.name) && !dependents.has(manifest.name));
  const names = new Set(chosen.map(({ manifest }) => manifest.name));
  const byName = new Map(chosen.map((entry) => [entry.manifest.name, entry]));
  const ordered = [];
  const done = new Set();
  const place = (name) => {
    if (done.has(name)) return;
    done.add(name);
    for (const dependency of internalDependencies(byName.get(name).manifest, names)) place(dependency);
    ordered.push(byName.get(name));
  };
  for (const name of [...names].sort()) place(name);
  return ordered.map(({ directory, manifest }) => ({ name: manifest.name, version: manifest.version, directory }));
}

/** Every dependency on a canary set member points at that member's exact canary version. */
export function pinToCanaries(manifest, set) {
  const versions = new Map(set.map(({ name, version }) => [name, version]));
  const pinned = structuredClone(manifest);
  for (const field of FIELDS) {
    for (const name of Object.keys(pinned[field] ?? {})) {
      if (versions.has(name)) pinned[field][name] = versions.get(name);
    }
  }
  return pinned;
}

/** Puts back every dependency range Changesets rewrote, so the snapshot changes versions only. */
export function keepDependencyRanges(before, after) {
  const kept = { ...after };
  for (const field of FIELDS) {
    if (before[field]) kept[field] = { ...before[field] };
    else delete kept[field];
  }
  return kept;
}

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

function writeJson(path, value) {
  writeFileSync(path, JSON.stringify(value, null, 2) + '\n');
}

function nativeNames() {
  return new Set(NATIVE_PACKAGES.map((path) => readJson(join(ROOT, path, 'package.json')).name));
}

function npm(args, options = {}) {
  return execFileSync('npm', args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit'], ...options });
}

function npmView(spec, field) {
  try {
    const out = npm(['view', spec, field, '--json'], { stdio: ['ignore', 'pipe', 'pipe'], timeout: 60_000 }).trim();
    return out ? JSON.parse(out) : undefined;
  } catch (error) {
    if (/\bE404\b/.test(`${error.stdout ?? ''}${error.stderr ?? ''}`)) return undefined;
    throw error;
  }
}

function writeOutput(values) {
  const lines = Object.entries(values).map(([key, value]) => `${key}=${value}`);
  for (const line of lines) console.log(line);
  if (process.env.GITHUB_OUTPUT) writeFileSync(process.env.GITHUB_OUTPUT, lines.join('\n') + '\n', { flag: 'a' });
}

function prepare(destination) {
  const before = new Map(workspaceEntries().map(({ directory, manifest }) => [directory, manifest]));
  execFileSync('bunx', ['changeset', 'version', '--snapshot', CANARY_TAG], { cwd: ROOT, stdio: 'inherit' });
  for (const [directory, manifest] of before) {
    const path = join(directory, 'package.json');
    writeJson(path, keepDependencyRanges(manifest, readJson(path)));
  }
  const set = canarySet(workspaceEntries(), nativeNames());
  mkdirSync(destination, { recursive: true });
  writeJson(join(destination, PLAN), set.map(({ name, version, directory }) => ({ name, version, directory })));
  if (set.length) {
    console.log(`Canary ${[...new Set(set.map(({ version }) => version))].join(', ')}:`);
    for (const { name, version } of set) console.log(`  ${name}@${version}`);
  } else {
    console.log('No package has a canary version: nothing is pending for npm.');
  }
  writeOutput({ count: set.length });
}

function plan(destination) {
  return readJson(join(destination, PLAN));
}

/** Whether every package in the plan already has a canary built from this commit. */
function published(destination, sha) {
  const set = plan(destination);
  const heads = set.map(({ name }) => [name, npmView(`${name}@${CANARY_TAG}`, 'gitHead')]);
  for (const [name, head] of heads) console.log(`${name}@${CANARY_TAG} was built from ${head ?? 'nothing yet'}`);
  const current = set.length > 0 && heads.every(([, head]) => head === sha);
  if (current) console.log(`Every canary was already built from ${sha}.`);
  writeOutput({ fresh: !current, build: !current || process.env.DRY_RUN === 'true' });
}

function tarball(destination, { name, version }) {
  return join(destination, 'tarballs', `${name.replace(/^@/, '').replace('/', '-')}-${version}.tgz`);
}

function pack(destination, sha) {
  const set = plan(destination);
  const members = new Set(set.map(({ name }) => name));
  execFileSync('bun', ['scripts/rewrite-workspace-deps.ts'], { cwd: ROOT, stdio: 'inherit' });
  for (const { directory } of workspaceEntries()) {
    const path = join(directory, 'package.json');
    const pinned = pinToCanaries(readJson(path), set);
    if (sha && members.has(pinned.name)) pinned.gitHead = sha;
    writeJson(path, pinned);
  }
  mkdirSync(join(destination, 'tarballs'), { recursive: true });
  for (const entry of set) {
    const [packed] = JSON.parse(npm(['pack', '--json', '--pack-destination', join(destination, 'tarballs')], { cwd: entry.directory }));
    if (packed?.name !== entry.name || packed.version !== entry.version) {
      throw new Error(`npm pack did not produce ${entry.name}@${entry.version}`);
    }
    renameSync(join(destination, 'tarballs', packed.filename), tarball(destination, entry));
    console.log(`Packed ${entry.name}@${entry.version} (${packed.entryCount} files, ${packed.size} bytes)`);
  }
}

function installedVersions(directory) {
  const scope = join(directory, 'node_modules', '@betteroffice');
  const found = [];
  const walk = (root) => {
    const at = join(root, 'node_modules', '@betteroffice');
    if (!existsSync(at)) return;
    for (const entry of readdirSync(at)) {
      const path = join(at, entry);
      const manifest = readJson(join(path, 'package.json'));
      found.push({ name: manifest.name, version: manifest.version, nested: at !== scope, gitHead: manifest.gitHead });
      walk(path);
    }
  };
  walk(directory);
  return found;
}

/** Installs the packed set alone in an empty project and checks what resolved. */
function verify(destination, sha) {
  const set = plan(destination);
  const project = mkdtempSync(join(tmpdir(), 'canary-install-'));
  writeJson(join(project, 'package.json'), { name: 'canary-install-check', private: true, version: '0.0.0' });
  npm(['install', '--no-audit', '--no-fund', '--ignore-scripts', ...set.map((entry) => tarball(destination, entry))], {
    cwd: project, stdio: ['ignore', 'inherit', 'inherit']
  });
  npm(['ls', '--all'], { cwd: project, stdio: ['ignore', 'ignore', 'inherit'] });
  const versions = new Map(set.map(({ name, version }) => [name, version]));
  const problems = [];
  for (const installed of installedVersions(project)) {
    const expected = versions.get(installed.name);
    if (installed.nested && expected) problems.push(`${installed.name}@${installed.version} is installed twice`);
    if (expected && installed.version !== expected) problems.push(`${installed.name} resolved to ${installed.version}, not ${expected}`);
    if (!expected && installed.version.includes('-')) problems.push(`${installed.name} resolved to the prerelease ${installed.version}`);
    if (expected && sha && installed.gitHead !== sha) problems.push(`${installed.name} records gitHead ${installed.gitHead}, not ${sha}`);
    console.log(`  ${installed.name}@${installed.version}${expected ? ' (canary)' : ' (published)'}${installed.nested ? ' nested' : ''}`);
  }
  for (const name of versions.keys()) {
    if (!installedVersions(project).some((installed) => installed.name === name)) problems.push(`${name} is not installed`);
  }
  if (problems.length) throw new Error(`The canary set does not install cleanly:\n${problems.join('\n')}`);
  console.log(`The ${set.length} canary tarballs install cleanly on their own.`);
}

function distTags(name) {
  return npmView(name, 'dist-tags') ?? {};
}

function packedManifest(path) {
  return JSON.parse(execFileSync('tar', ['-xzOf', path, 'package/package.json'], { encoding: 'utf8' }));
}

/** Refuses any tarball whose manifest is not the canary the plan names. */
export function checkPublishable(set, manifestOf) {
  const problems = [];
  for (const entry of set) {
    const manifest = manifestOf(entry);
    if (!isCanaryVersion(entry.version)) problems.push(`${entry.name}@${entry.version} is not a canary version`);
    if (manifest.name !== entry.name || manifest.version !== entry.version) {
      problems.push(`the ${entry.name} tarball holds ${manifest.name}@${manifest.version}, not ${entry.version}`);
    }
  }
  if (problems.length) throw new Error(problems.join('\n'));
}

function pause(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function publish(destination) {
  const set = plan(destination);
  checkPublishable(set, (entry) => packedManifest(tarball(destination, entry)));
  const latest = new Map(set.map(({ name }) => [name, distTags(name).latest]));
  for (const entry of set) {
    if (npmView(`${entry.name}@${entry.version}`, 'version') === entry.version) {
      console.log(`${entry.name}@${entry.version} is already published.`);
      continue;
    }
    npm(['publish', tarball(destination, entry), '--tag', CANARY_TAG, '--access', 'public', '--provenance'], {
      cwd: ROOT, stdio: ['ignore', 'inherit', 'inherit']
    });
  }
  let problems = [];
  const deadline = Date.now() + 5 * 60_000;
  for (let attempt = 0; attempt < 6 && (attempt === 0 || Date.now() < deadline); attempt += 1) {
    if (attempt) pause(10_000);
    problems = [];
    for (const { name, version } of set) {
      let tags;
      try {
        tags = distTags(name);
      } catch (error) {
        problems.push(`${name}: ${error.message}`);
        continue;
      }
      const before = latest.get(name);
      if (String(tags.latest ?? '').includes('-') || (tags.latest === undefined && before !== undefined)) {
        problems.push(`${name} latest is ${tags.latest}, was ${before}`);
      } else if (tags.latest !== before && attempt === 0) {
        console.log(`${name} latest moved from ${before} to ${tags.latest}, a stable release published meanwhile.`);
      }
      if (tags[CANARY_TAG] !== version) problems.push(`${name}@${CANARY_TAG} is ${tags[CANARY_TAG]}, not ${version}`);
    }
    if (!problems.length) break;
  }
  if (problems.length) throw new Error(problems.join('\n'));
  console.log(`Published ${set.length} canaries; latest points at no canary.`);
}

const COMMANDS = { prepare, published, pack, verify, publish };

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [command, destination, sha] = process.argv.slice(2);
  if (!COMMANDS[command] || !destination) {
    console.error(`usage: node scripts/canary-release.mjs <${Object.keys(COMMANDS).join('|')}> <directory> [sha]`);
    process.exit(2);
  }
  COMMANDS[command](resolve(destination), sha);
}
