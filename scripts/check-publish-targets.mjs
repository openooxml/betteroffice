// Fail a release before it publishes half a set.
//
// Trusted Publishing attaches only to a name that already exists, so a
// brand-new package or crate cannot be created by the workflow. `changeset
// publish` attempts every unpublished package in parallel, tags the ones that
// worked, and only then throws — leaving live dependents pointing at a name
// that 404s. This runs first and names what has to be bootstrapped by hand.
import { fileURLToPath } from 'node:url';
import { publishedPackageVersions } from './published-packages.mjs';
import { RUST_CRATES } from './rust-crates.mjs';

const NPM_REGISTRY = process.env.NPM_REGISTRY_URL ?? 'https://registry.npmjs.org';
const CRATES_REGISTRY = process.env.CRATES_REGISTRY_URL ?? 'https://crates.io/api/v1/crates';
// crates.io rejects a request without one.
const USER_AGENT = 'betteroffice-release (https://github.com/openooxml/betteroffice)';

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchRegistry(url, headers = {}) {
  let lastError;
  for (let attempt = 0; attempt < 5; attempt++) {
    let response;
    try {
      response = await fetch(url, {
        headers: { 'User-Agent': USER_AGENT, ...headers },
        cache: 'no-store'
      });
    } catch (error) {
      lastError = error;
      await sleep(2 ** attempt * 500);
      continue;
    }
    if (response.ok || response.status === 404) return response;
    if (response.status !== 429 && response.status < 500) {
      throw new Error(`${url} returned ${response.status}`);
    }
    lastError = new Error(`${url} returned ${response.status}`);
    await sleep(2 ** attempt * 500);
  }
  throw lastError;
}

/** `missing` (the name is unclaimed), `new` (this version is), or `published`. */
export async function auditNpmPackages(packages, registry = NPM_REGISTRY) {
  const audit = [];
  for (const { name, version } of packages) {
    const response = await fetchRegistry(`${registry}/${encodeURIComponent(name)}`, {
      Accept: 'application/vnd.npm.install-v1+json'
    });
    if (response.status === 404) {
      audit.push({ name, version, state: 'missing' });
      continue;
    }
    const versions = (await response.json()).versions ?? {};
    audit.push({ name, version, state: version in versions ? 'published' : 'new' });
  }
  return audit;
}

/** `missing` (the crate is unclaimed) or `present`. Versions are the publisher's job. */
export async function auditCrates(names, registry = CRATES_REGISTRY) {
  const audit = [];
  for (const name of names) {
    const response = await fetchRegistry(`${registry}/${encodeURIComponent(name)}`);
    audit.push({ name, state: response.status === 404 ? 'missing' : 'present' });
  }
  return audit;
}

async function checkNpm() {
  const audit = await auditNpmPackages(publishedPackageVersions());
  for (const { name, version, state } of audit) {
    if (state === 'published') console.log(`${name}@${version} is on npm.`);
    if (state === 'new') console.log(`${name}@${version} is a new version of a package on npm.`);
  }

  const missing = audit.filter((entry) => entry.state === 'missing');
  if (missing.length === 0) return 0;
  for (const { name } of missing) console.error(`${name} is not on npm.`);
  console.error(
    'release.yml publishes with OIDC Trusted Publishing, which cannot create a package. `changeset publish` would publish the rest of the set before it fails on these.'
  );
  console.error('Publish each by hand first: RELEASING.md, "Initial npm release of a new package".');
  return 1;
}

async function checkCrates() {
  const audit = await auditCrates(RUST_CRATES.map((crate) => crate.name));
  // Mirrors release.yml's `Detect crates.io bootstrap token`: a token can create a crate.
  const bootstrap = Boolean(process.env.CRATES_IO_BOOTSTRAP_TOKEN);
  for (const { name, state } of audit) {
    if (state === 'present') console.log(`${name} is on crates.io.`);
    if (state === 'missing' && bootstrap) {
      console.log(`${name} is not on crates.io; CRATES_IO_BOOTSTRAP_TOKEN can create it.`);
    }
  }

  const missing = bootstrap ? [] : audit.filter((entry) => entry.state === 'missing');
  if (missing.length === 0) return 0;
  for (const { name } of missing) console.error(`${name} is not on crates.io.`);
  console.error(
    'The OIDC token from crates-io-auth-action cannot create a crate, and the crates publish stops at the first failure.'
  );
  console.error(
    'Set CRATES_IO_BOOTSTRAP_TOKEN, or publish each by hand first: RELEASING.md, "Initial crates.io release".'
  );
  return 1;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [mode, ...rest] = process.argv.slice(2);
  if (rest.length > 0 || (mode !== '--npm' && mode !== '--crates')) {
    console.error('check-publish-targets.mjs: expected exactly one of --npm or --crates');
    process.exit(2);
  }
  process.exit(await (mode === '--npm' ? checkNpm() : checkCrates()));
}
