// Fail a release before it publishes half a set.
//
// Trusted Publishing attaches only to a name that already exists, so a
// brand-new package or crate cannot be created by the workflow. `changeset
// publish` attempts every unpublished package in parallel, tags the ones that
// worked, and only then throws — leaving live dependents pointing at a name
// that 404s. This runs first and names what has to be bootstrapped by hand.
import { appendFileSync, existsSync, readdirSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { publishedPackageVersions, workspaceEntries } from './published-packages.mjs';
import { NATIVE_PACKAGES, platformPackageVersions } from './native-bindings.mjs';
import { REQUEST_TIMEOUT_MS, attemptRegistry, backoffMs, fetchRegistry, sleep } from './registry-fetch.mjs';
import { RUST_PUBLISH_CRATES } from './rust-crates.mjs';

const require = createRequire(import.meta.url);
const { satisfies } = createRequire(require.resolve('@changesets/cli'))('semver');
const ROOT = fileURLToPath(new URL('..', import.meta.url));
const NPM_REGISTRY = process.env.NPM_REGISTRY_URL ?? 'https://registry.npmjs.org';
const CRATES_REGISTRY = process.env.CRATES_REGISTRY_URL ?? 'https://crates.io/api/v1/crates';
const CDN_BASE = process.env.CDN_BASE_URL ?? 'https://cdn.jsdelivr.net/npm';
const FONT_PACKAGES = new Set(['@betteroffice/fonts-latin', '@betteroffice/fonts-intl', '@betteroffice/fonts-cjk']);
const VERIFY_WAIT_MS = 10 * 60_000;

/** Bodies by name; with `readBody` false, `true` for a name that exists and `null` for one that does not. */
function npmPackages(registry, timeoutMs, readBody = true) {
  const cache = new Map();
  return (name) => {
    if (!cache.has(name)) {
      cache.set(name, fetchRegistry(
        `${registry.replace(/\/+$/, '')}/${encodeURIComponent(name)}`,
        { headers: { Accept: 'application/vnd.npm.install-v1+json' }, timeoutMs, readBody }
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

/** Whether the release plan skipped native publishing. */
function skipsNative(env = process.env) {
  return env.RELEASE_SKIP_NATIVE?.toLowerCase() === 'true';
}

/** The native roots the release's native job publishes after this one, at their manifest versions. */
export function nativeReleaseVersions(workspaces = workspaceEntries()) {
  const roots = new Set(NATIVE_PACKAGES.map((path) => resolve(ROOT, path)));
  return workspaces
    .filter(({ directory }) => roots.has(resolve(directory)))
    .map(({ manifest }) => ({ name: manifest.name, version: manifest.version }));
}

export function releaseEntries(packages, workspaces) {
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
  if (suffix === '^' || suffix === '~') return suffix + version;
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

function nativeDependencies(dependencies, natives) {
  const byName = new Map(natives.map((native) => [native.name, native]));
  return dependencies.flatMap(({ name, dependency }) => {
    const native = byName.get(dependency);
    return native ? [{ name, native }] : [];
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

/** The release set, the workspaces and the sibling references `checkNpm` would otherwise read from the repository. */
export function repositoryScan(packages = publishedPackageVersions(), workspaces = workspaceEntries()) {
  return { packages, workspaces, references: siblingReferences(packages, workspaces) };
}

/** `missing` (the crate is unclaimed) or `present`. Versions are the publisher's job. */
export async function auditCrates(names, registry = CRATES_REGISTRY, timeoutMs = REQUEST_TIMEOUT_MS) {
  const audit = [];
  for (const name of names) {
    const body = await fetchRegistry(`${registry}/${encodeURIComponent(name)}`, { timeoutMs });
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
  natives = nativeReleaseVersions(workspaces),
  later = skipsNative() ? [] : natives,
  platforms = skipsNative() ? [] : platformPackageVersions(),
  holdNative = true,
  onHeld = (_held) => {},
  references
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

  const pending = audit.filter(({ state }) => state !== 'published').map(({ name, version }) => ({ name, version }));
  const dependencies = workspaceDependencies(pending, workspaces);
  const nativeDeps = holdNative ? nativeDependencies(dependencies, natives) : [];
  const needsNativeVersion = new Set(nativeDeps.map(({ native }) => native.name));
  const packageExists = npmPackages(registry, timeoutMs, false);
  const nativeMissing = [];
  let lookupFailure;
  for (const { name } of [...later, ...platforms]) {
    try {
      const lookup = needsNativeVersion.has(name) ? fetchPackage : packageExists;
      if ((await lookup(name)) === null) nativeMissing.push(name);
    } catch (error) {
      lookupFailure = { name, error };
      break;
    }
  }
  if (nativeMissing.length > 0) {
    failed = true;
    for (const name of nativeMissing) out.error(`${name} is not on npm.`);
    out.error(
      "release.yml's native-npm job publishes these after this one, with OIDC Trusted Publishing, which cannot create a package. Publish each by hand first: RELEASING.md, \"Native packages\"."
    );
  }
  if (lookupFailure) {
    out.error(`${lookupFailure.name} could not be looked up on npm: ${lookupFailure.error.message}`);
    return 1;
  }

  const heldNames = new Set();
  for (const { name, native } of nativeDeps) {
    const body = await fetchPackage(native.name);
    if (!(native.version in (body?.versions ?? {}))) heldNames.add(name);
  }
  const held = pending.filter(({ name }) => heldNames.has(name));
  for (const { name, version } of held) out.log(`${name}@${version} is held for its native dependencies.`);
  onHeld(held);

  const immediate = packages.filter(({ name }) => !heldNames.has(name));
  const releaseVersions = new Map(immediate.map(({ name, version }) => [name, version]));
  const nativeVersions = new Map(natives.map(({ name, version }) => [name, version]));
  const available = async (name, range, exact = false) => {
    const matches = (version) => exact ? version === range : satisfies(version, range);
    const version = releaseVersions.get(name);
    if (version && matches(version)) return true;
    const body = await fetchPackage(name);
    return Object.keys(body?.versions ?? {}).some(matches);
  };
  for (const { name, version, dependency, range } of dependencies) {
    if (heldNames.has(name)) {
      if (heldNames.has(dependency)) {
        out.error(`${name}@${version} depends on ${dependency}@${range}, which is also held`);
        failed = true;
        continue;
      }
      const nativeVersion = nativeVersions.get(dependency);
      if (nativeVersion !== undefined) {
        if (!satisfies(nativeVersion, range)) {
          out.error(`${name}@${version} depends on ${dependency}@${range}, but the native manifest version ${nativeVersion} does not satisfy that range`);
          failed = true;
        }
        continue;
      }
    }
    if (await available(dependency, range)) continue;
    out.error(`${name}@${version} depends on ${dependency}@${range}, which is neither in this release nor on npm`);
    failed = true;
  }
  for (const { dependent, name, version, file, line } of references ?? siblingReferences(packages, workspaces)) {
    if (await available(name, version, true)) continue;
    out.error(`${dependent} builds references to ${name}@${version} (${file}:${line}), which is neither in this release nor on npm`);
    failed = true;
  }
  return failed ? 1 : 0;
}

function npmVersionUrl(registry, name, version) {
  return `${registry.replace(/\/+$/, '')}/${encodeURIComponent(name)}/${encodeURIComponent(version)}`;
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

function cdnPublishTargets(packages, workspaces, baseUrl) {
  return releaseEntries(packages, workspaces).flatMap(({ name, version, directory }) => {
    const tag = `${name}@${version}`;
    const asset = FONT_PACKAGES.has(name)
      ? JSON.parse(readFileSync(join(directory, 'assets', 'manifest.json'), 'utf8'))?.faces?.[0]?.file
      : undefined;
    if (FONT_PACKAGES.has(name) && (typeof asset !== 'string' || !asset.endsWith('.zst'))) throw new Error(`${tag} has no compressed assets in its manifest`);
    return [
      { label: tag, service: 'the CDN', method: 'HEAD', url: `${baseUrl}/${tag}/package.json` },
      ...(asset ? [
        { label: `${tag}/assets/${asset}`, service: 'the CDN', method: 'HEAD', url: `${baseUrl}/${tag}/assets/${asset}` }
      ] : [])
    ];
  });
}

export async function missingHeldNatives(packages, {
  registry = NPM_REGISTRY,
  timeoutMs = REQUEST_TIMEOUT_MS,
  maxWaitMs = VERIFY_WAIT_MS,
  workspaces = workspaceEntries(),
  natives = nativeReleaseVersions(workspaces),
  wait = false
} = {}) {
  const dependencies = workspaceDependencies(packages, workspaces);
  const required = [...new Set(nativeDependencies(dependencies, natives).map(({ native }) => native))];
  const deadline = Date.now() + maxWaitMs;
  const missing = await Promise.all(required.map(async ({ name, version }) => {
    const url = npmVersionUrl(registry, name, version);
    const found = wait
      ? await waitForPublished(url, 'GET', timeoutMs, deadline)
      : (await fetchRegistry(url, { timeoutMs, readBody: false })) !== null;
    return found ? null : { name, version };
  }));
  return missing.filter(Boolean);
}

export async function verifyPublished({
  registry = NPM_REGISTRY,
  cdnBase = CDN_BASE,
  out = console,
  timeoutMs = REQUEST_TIMEOUT_MS,
  maxWaitMs = VERIFY_WAIT_MS,
  packages = publishedPackageVersions(),
  workspaces = workspaceEntries(),
  held = [],
  annotate = false,
  summary
} = {}) {
  const heldVersions = new Set(held.map(({ name, version }) => `${name}@${version}`));
  for (const tag of heldVersions) out.log(`${tag} is held for its native dependencies.`);
  const immediate = packages.filter(({ name, version }) => !heldVersions.has(`${name}@${version}`));
  const cdnPackages = new Map();
  for (const { name, version } of [...immediate, ...siblingReferences(immediate, workspaces)]) {
    const tag = `${name}@${version}`;
    if (!heldVersions.has(tag)) cdnPackages.set(tag, { name, version });
  }
  const targets = [
    ...immediate.map(({ name, version }) => ({
      label: `${name}@${version}`, service: 'npm', method: 'GET',
      url: npmVersionUrl(registry, name, version)
    })),
    ...cdnPublishTargets([...cdnPackages.values()], workspaces, cdnBase.replace(/\/+$/, ''))
  ];
  const deadline = Date.now() + Math.min(maxWaitMs, VERIFY_WAIT_MS);
  const results = await Promise.all(targets.map(async (target) => ({
    ...target,
    found: await waitForPublished(target.url, target.method, timeoutMs, deadline)
  })));
  for (const { label, service, found } of results) {
    const message = `${label} is ${found ? '' : 'not '}on ${service} after publishing.`;
    if (found) out.log(message);
    else out.error(message);
  }
  const missing = results.filter(({ found }) => !found);
  if (missing.length === 0) return 0;
  if (annotate) {
    for (const { label, service } of missing) {
      out.error(`::error title=Release incomplete::${label} is not on ${service} after publishing`);
    }
  }
  if (summary) {
    appendFileSync(summary, [
      '### Release incomplete',
      '',
      'These versions were not found after publishing. If the publish step succeeded, the registry or CDN is still propagating: check them again later. If it failed, publish what is missing (RELEASING.md); native packages and the deploy are dispatched with publish-native.yml and deploy.yml.',
      '',
      ...missing.map(({ label, service }) => `- \`${label}\` on ${service}`),
      ''
    ].join('\n'));
  }
  return 1;
}

const BUMP_ORDER = { none: 0, patch: 1, minor: 2, major: 3 };

function changesetConfig() {
  return JSON.parse(readFileSync(join(ROOT, '.changeset/config.json'), 'utf8'));
}

/** Major bumps in a changesets release plan that no changeset asked for, each with the peer range that forced it. */
export function unaskedMajors(plan, workspaces = workspaceEntries(), fixed = changesetConfig().fixed ?? []) {
  const asked = new Map();
  for (const changeset of plan.changesets) {
    for (const { name, type } of changeset.releases) {
      for (const member of fixed.find((group) => group.includes(name)) ?? [name]) {
        if (BUMP_ORDER[type] > BUMP_ORDER[asked.get(member) ?? 'none']) asked.set(member, type);
      }
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

export async function checkPlan({ out = console, plan, workspaces = workspaceEntries(), fixed } = {}) {
  const resolved = plan ?? await createRequire(require.resolve('@changesets/cli'))('@changesets/get-release-plan').default(ROOT);
  const majors = unaskedMajors(resolved, workspaces, fixed);
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
  let options;
  if (mode === '--npm' && process.env.GITHUB_OUTPUT) {
    options = { onHeld: (held) => appendFileSync(process.env.GITHUB_OUTPUT, `held=${JSON.stringify(held)}\n`) };
  }
  if (mode === '--verify-published') {
    options = {
      annotate: process.env.GITHUB_ACTIONS === 'true',
      summary: process.env.GITHUB_STEP_SUMMARY || undefined,
      held: JSON.parse(process.env.HELD_NPM_PACKAGES || '[]')
    };
  }
  process.exit(await modes[mode](options));
}
