// Fail a release before it publishes half a set.
//
// Trusted Publishing attaches only to a name that already exists, so a
// brand-new package or crate cannot be created by the workflow. `changeset
// publish` attempts every unpublished package in parallel, tags the ones that
// worked, and only then throws — leaving live dependents pointing at a name
// that 404s. This runs first and names what has to be bootstrapped by hand.
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { publishedPackageVersions, workspaceEntries } from './published-packages.mjs';
import { NATIVE_PACKAGES } from './native-bindings.mjs';
import { RUST_PUBLISH_CRATES } from './rust-crates.mjs';

const require = createRequire(import.meta.url);
const { satisfies } = createRequire(require.resolve('@changesets/cli'))('semver');
const ROOT = fileURLToPath(new URL('..', import.meta.url));
const NPM_REGISTRY = process.env.NPM_REGISTRY_URL ?? 'https://registry.npmjs.org';
const CRATES_REGISTRY = process.env.CRATES_REGISTRY_URL ?? 'https://crates.io/api/v1/crates';
const CDN_BASE = process.env.CDN_BASE_URL ?? 'https://cdn.jsdelivr.net/npm';
const VERIFY_WAIT_MS = 10 * 60_000;
// crates.io rejects a request without one.
const USER_AGENT = 'betteroffice-release (https://github.com/openooxml/betteroffice)';
const ATTEMPTS = 5;
const REQUEST_TIMEOUT_MS = Number(process.env.REGISTRY_TIMEOUT_MS ?? 15_000);
// A registry may ask for an hour; a release gate waits a minute.
const RETRY_AFTER_CAP_MS = 60_000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Half the backoff plus a random half, so parallel releases don't retry in lockstep. */
function backoffMs(attempt) {
  const base = 2 ** attempt * 500;
  return base / 2 + Math.random() * (base / 2);
}

/** `Retry-After` in ms — seconds or an HTTP-date — or undefined if there is none. */
function retryAfterMs(response) {
  const header = response.headers.get('Retry-After');
  if (!header) return undefined;
  const seconds = Number(header);
  const ms = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header) - Date.now();
  if (!Number.isFinite(ms)) return undefined;
  return Math.min(Math.max(ms, 0), RETRY_AFTER_CAP_MS);
}

/** The parsed body, `null` when the name is unclaimed, or a failure to weigh. */
async function attemptRegistry(url, headers, timeoutMs, method = 'GET') {
  const response = await fetch(url, {
    method,
    headers: { 'User-Agent': USER_AGENT, ...headers },
    cache: 'no-store',
    signal: AbortSignal.timeout(timeoutMs)
  });
  if (response.status === 404) return { body: null };
  if (response.ok) return { body: method === 'HEAD' ? true : await response.json() };
  const error = new Error(`${url} returned ${response.status}`);
  const retry = response.status === 429 || response.status >= 500;
  return { error, retry, wait: retry ? retryAfterMs(response) : undefined };
}

async function fetchRegistry(url, headers = {}, timeoutMs = REQUEST_TIMEOUT_MS) {
  let lastError;
  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    let result;
    try {
      result = await attemptRegistry(url, headers, timeoutMs);
    } catch (error) {
      // A refused connection, a request past REQUEST_TIMEOUT_MS, or a body that stops mid-JSON.
      result = { error, retry: true };
    }
    if (!result.error) return result.body;
    if (!result.retry) throw result.error;
    lastError = result.error;
    if (attempt < ATTEMPTS - 1) await sleep(result.wait ?? backoffMs(attempt));
  }
  throw lastError;
}

function npmPackages(registry, timeoutMs) {
  const cache = new Map();
  return (name) => {
    if (!cache.has(name)) {
      cache.set(name, fetchRegistry(
        `${registry.replace(/\/+$/, '')}/${encodeURIComponent(name)}`,
        { Accept: 'application/vnd.npm.install-v1+json' },
        timeoutMs
      ));
    }
    return cache.get(name);
  };
}

async function auditPackages(packages, fetchPackage) {
  const audit = [];
  for (const { name, version } of packages) {
    const body = await fetchPackage(name);
    if (body === null) {
      audit.push({ name, version, state: 'missing' });
      continue;
    }
    const versions = body.versions ?? {};
    audit.push({ name, version, state: version in versions ? 'published' : 'new' });
  }
  return audit;
}

/** `missing` (the name is unclaimed), `new` (this version is), or `published`. */
export async function auditNpmPackages(packages, registry = NPM_REGISTRY, timeoutMs = REQUEST_TIMEOUT_MS) {
  return auditPackages(packages, npmPackages(registry, timeoutMs));
}

/** The native roots the release's native job publishes after this one, at their manifest versions. */
export function nativeReleaseVersions(workspaces = workspaceEntries()) {
  const roots = new Set(NATIVE_PACKAGES.map((path) => resolve(ROOT, path)));
  return workspaces
    .filter(({ directory }) => roots.has(resolve(directory)))
    .map(({ manifest }) => ({ name: manifest.name, version: manifest.version }));
}

function releaseEntries(packages, workspaces) {
  const byName = new Map(workspaces.map((entry) => [entry.manifest.name, entry]));
  return packages.map(({ name, version }) => {
    const entry = byName.get(name);
    if (!entry) throw new Error(`${name} is not a workspace package.`);
    return { ...entry, name, version };
  });
}

function dependencyRange(range, version) {
  if (!range.startsWith('workspace:')) return range;
  const suffix = range.slice('workspace:'.length);
  if (suffix === '*' || suffix === '') return version;
  if (suffix.startsWith('^') || suffix.startsWith('~')) return suffix[0] + version;
  return suffix;
}

function workspaceDependencies(packages, workspaces) {
  const byName = new Map(workspaces.map(({ manifest }) => [manifest.name, manifest]));
  return releaseEntries(packages, workspaces).flatMap(({ name, version, manifest }) => {
    const dependencies = [];
    for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
      for (const [dependency, range] of Object.entries(manifest[field] ?? {})) {
        const sibling = byName.get(dependency);
        if (sibling) dependencies.push({ name, version, dependency, range: dependencyRange(range, sibling.version) });
      }
    }
    return dependencies;
  });
}

function* sourceFiles(directory) {
  if (!existsSync(directory)) return;
  for (const entry of readdirSync(directory, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) yield* sourceFiles(path);
    if (entry.isFile() && /\.(?:ts|tsx|js|mjs)$/.test(entry.name) && !/\.test\./.test(entry.name)) yield path;
  }
}

export function siblingReferences(packages = publishedPackageVersions(), workspaces = workspaceEntries()) {
  const byDirectory = new Map(workspaces.map((entry) => [resolve(entry.directory), entry.manifest]));
  const references = [];
  for (const { name, directory } of releaseEntries(packages, workspaces)) {
    for (const file of sourceFiles(join(directory, 'src'))) {
      const source = ts.createSourceFile(file, readFileSync(file, 'utf8'), ts.ScriptTarget.Latest);
      function visit(node) {
        const literal = ts.isImportDeclaration(node) || ts.isExportDeclaration(node)
          ? node.moduleSpecifier
          : ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword
            ? node.arguments[0] : undefined;
        if (literal && ts.isStringLiteralLike(literal) && literal.text.startsWith('.') && literal.text.endsWith('/package.json')) {
          const sibling = byDirectory.get(dirname(resolve(dirname(file), literal.text)));
          if (sibling?.version && sibling.name !== name) {
            references.push({
              dependent: name,
              name: sibling.name,
              version: sibling.version,
              file: relative(ROOT, file),
              line: source.getLineAndCharacterOfPosition(literal.getStart(source)).line + 1
            });
          }
        }
        ts.forEachChild(node, visit);
      }
      visit(source);
    }
  }
  return references;
}

/** `missing` (the crate is unclaimed) or `present`. Versions are the publisher's job. */
export async function auditCrates(names, registry = CRATES_REGISTRY, timeoutMs = REQUEST_TIMEOUT_MS) {
  const audit = [];
  for (const name of names) {
    const body = await fetchRegistry(`${registry}/${encodeURIComponent(name)}`, {}, timeoutMs);
    audit.push({ name, state: body === null ? 'missing' : 'present' });
  }
  return audit;
}

export async function checkNpm({
  registry = NPM_REGISTRY,
  out = console,
  timeoutMs = REQUEST_TIMEOUT_MS,
  packages = publishedPackageVersions(),
  workspaces = workspaceEntries(),
  later = nativeReleaseVersions(workspaces)
} = {}) {
  const fetchPackage = npmPackages(registry, timeoutMs);
  const audit = await auditPackages(packages, fetchPackage);
  for (const { name, version, state } of audit) {
    if (state === 'published') out.log(`${name}@${version} is on npm.`);
    if (state === 'new') out.log(`${name}@${version} is a new version of a package on npm.`);
  }

  const missing = audit.filter((entry) => entry.state === 'missing');
  let failed = missing.length > 0;
  if (failed) {
    for (const { name } of missing) out.error(`${name} is not on npm.`);
    out.error(
      'release.yml publishes with OIDC Trusted Publishing, which cannot create a package. `changeset publish` would publish the rest of the set before it fails on these.'
    );
    out.error('Publish each by hand first: RELEASING.md, "Initial npm release of a new package".');
  }

  const releaseVersions = new Map([...later, ...packages].map(({ name, version }) => [name, version]));
  const available = async (name, range, exact = false) => {
    const matches = (version) => exact ? version === range : satisfies(version, range);
    const version = releaseVersions.get(name);
    if (version && matches(version)) return true;
    const body = await fetchPackage(name);
    return Object.keys(body?.versions ?? {}).some(matches);
  };
  for (const { name, version, dependency, range } of workspaceDependencies(packages, workspaces)) {
    if (await available(dependency, range)) continue;
    out.error(`${name}@${version} depends on ${dependency}@${range}, which is neither in this release nor on npm`);
    failed = true;
  }
  for (const { dependent, name, version, file, line } of siblingReferences(packages, workspaces)) {
    if (await available(name, version, true)) continue;
    out.error(`${dependent} builds references to ${name}@${version} (${file}:${line}), which is neither in this release nor on npm`);
    failed = true;
  }
  return failed ? 1 : 0;
}

async function waitForPublished(url, method, timeoutMs, deadline) {
  for (let attempt = 0; Date.now() < deadline; attempt++) {
    let result;
    try {
      result = await attemptRegistry(url, {}, Math.max(1, Math.min(timeoutMs, deadline - Date.now())), method);
    } catch {
      result = { retry: true };
    }
    if (result.body != null) return true;
    if (result.error && !result.retry) return false;
    const remaining = deadline - Date.now();
    if (remaining <= 0) break;
    await sleep(Math.min(result.wait ?? backoffMs(Math.min(attempt, 7)), remaining));
  }
  return false;
}

export async function verifyPublished({
  registry = NPM_REGISTRY,
  cdnBase = CDN_BASE,
  out = console,
  timeoutMs = REQUEST_TIMEOUT_MS,
  maxWaitMs = VERIFY_WAIT_MS,
  packages = publishedPackageVersions(),
  workspaces = workspaceEntries()
} = {}) {
  const cdnPackages = new Map();
  for (const { name, version } of [...packages, ...siblingReferences(packages, workspaces)]) {
    cdnPackages.set(`${name}@${version}`, { name, version });
  }
  const targets = [
    ...packages.map(({ name, version }) => ({
      name, version, service: 'npm', method: 'GET',
      url: `${registry.replace(/\/+$/, '')}/${encodeURIComponent(name)}/${encodeURIComponent(version)}`
    })),
    ...[...cdnPackages.values()].map(({ name, version }) => ({
      name, version, service: 'the CDN', method: 'HEAD',
      url: `${cdnBase.replace(/\/+$/, '')}/${name}@${version}/package.json`
    }))
  ];
  const deadline = Date.now() + Math.min(maxWaitMs, VERIFY_WAIT_MS);
  const results = await Promise.all(targets.map(async (target) => ({
    ...target,
    found: await waitForPublished(target.url, target.method, timeoutMs, deadline)
  })));
  for (const { name, version, service, found } of results) {
    const message = `${name}@${version} is ${found ? '' : 'not '}on ${service} after publishing.`;
    if (found) out.log(message);
    else out.error(message);
  }
  return results.some(({ found }) => !found) ? 1 : 0;
}

const BUMP_ORDER = { none: 0, patch: 1, minor: 2, major: 3 };

/** Major bumps in a changesets release plan that no changeset asked for, each with the peer range that forced it. */
export function unaskedMajors(plan, workspaces = workspaceEntries()) {
  const asked = new Map();
  for (const changeset of plan.changesets) {
    for (const { name, type } of changeset.releases) {
      if (BUMP_ORDER[type] > BUMP_ORDER[asked.get(name) ?? 'none']) asked.set(name, type);
    }
  }
  const releases = new Map(plan.releases.map((release) => [release.name, release]));
  const manifests = new Map(workspaces.map(({ manifest }) => [manifest.name, manifest]));
  const unasked = plan.releases.filter(({ name, type }) => type === 'major' && asked.get(name) !== 'major');
  return unasked.map(({ name, oldVersion, newVersion }) => {
    const causes = [];
    for (const [dependency, range] of Object.entries(manifests.get(name)?.peerDependencies ?? {})) {
      const release = releases.get(dependency);
      if (!release || release.type === 'none' || release.type === 'patch') continue;
      const resolved = dependencyRange(range, release.oldVersion);
      if (!satisfies(release.newVersion, resolved)) {
        causes.push(`peer ${dependency}@${range} does not cover its next version ${release.newVersion}`);
      }
    }
    return { name, oldVersion, newVersion, causes };
  });
}

export async function checkPlan({ out = console, plan, workspaces = workspaceEntries() } = {}) {
  const resolved = plan ?? await createRequire(require.resolve('@changesets/cli'))('@changesets/get-release-plan').default(ROOT);
  const majors = unaskedMajors(resolved, workspaces);
  for (const { name, oldVersion, newVersion, causes } of majors) {
    out.error(`${name} ${oldVersion} -> ${newVersion} is a major bump no changeset asked for${causes.length ? `: ${causes.join('; ')}` : ' (raised with its fixed group)'}`);
  }
  return majors.length ? 1 : 0;
}

export async function checkCrates({
  registry = CRATES_REGISTRY,
  out = console,
  timeoutMs = REQUEST_TIMEOUT_MS,
  bootstrap = Boolean(process.env.CRATES_IO_BOOTSTRAP_TOKEN)
} = {}) {
  const audit = await auditCrates(RUST_PUBLISH_CRATES.map((crate) => crate.name), registry, timeoutMs);
  for (const { name, state } of audit) {
    if (state === 'present') out.log(`${name} is on crates.io.`);
    if (state === 'missing' && bootstrap) {
      out.log(`${name} is not on crates.io; will be created with the bootstrap token.`);
    }
  }

  const missing = bootstrap ? [] : audit.filter((entry) => entry.state === 'missing');
  if (missing.length === 0) return 0;
  for (const { name } of missing) out.error(`${name} is not on crates.io.`);
  out.error(
    'The OIDC token from crates-io-auth-action cannot create a crate; publish-crates.mjs publishes present crates with OIDC and creates each missing one with CRATES_IO_BOOTSTRAP_TOKEN, stopping at the first failure.'
  );
  out.error(
    'Set CRATES_IO_BOOTSTRAP_TOKEN so the missing crates are created with it, or publish each by hand first: RELEASING.md, "Initial crates.io release".'
  );
  return 1;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [mode, ...rest] = process.argv.slice(2);
  const modes = { '--npm': checkNpm, '--crates': checkCrates, '--verify-published': verifyPublished, '--plan': checkPlan };
  if (rest.length > 0 || !Object.hasOwn(modes, mode)) {
    console.error('check-publish-targets.mjs: expected exactly one of --npm, --crates, --verify-published or --plan');
    process.exit(2);
  }
  process.exit(await modes[mode]());
}
