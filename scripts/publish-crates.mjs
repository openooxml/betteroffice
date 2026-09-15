import { appendFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  RUST_PUBLISH_CRATES,
  cargoMetadata,
  run,
  rustReleaseVersion,
  validateRustTrain
} from './rust-crates.mjs';

const USER_AGENT = 'betteroffice-release (https://github.com/openooxml/betteroffice)';
const EXPECTED_OWNER = process.env.CRATES_IO_OWNER ?? 'eliahilse';
const WAIT_TIMEOUT_MS = 5 * 60 * 1000;
const WAIT_INTERVAL_MS = 10 * 1000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchRegistry(url) {
  let lastError;
  for (let attempt = 0; attempt < 6; attempt++) {
    let response;
    try {
      response = await fetch(url, {
        headers: { 'User-Agent': USER_AGENT },
        cache: 'no-store'
      });
    } catch (error) {
      lastError = error;
      await sleep(2 ** attempt * 1000);
      continue;
    }
    if (response.ok || response.status === 404) return response;
    if (response.status !== 429 && response.status < 500) {
      throw new Error(`${url} returned ${response.status}`);
    }
    lastError = new Error(`${url} returned ${response.status}`);
    await sleep(2 ** attempt * 1000);
  }
  throw lastError;
}

/** Pick the publish token: OIDC for crates that exist, bootstrap for new ones. */
export function selectPublishToken({ name, exists, oidcToken, bootstrapToken }) {
  if (exists) {
    if (!oidcToken) throw new Error(`${name} is on crates.io but CARGO_REGISTRY_TOKEN is missing`);
    return { token: oidcToken, source: 'oidc' };
  }
  if (!bootstrapToken) {
    throw new Error(
      `${name} is not on crates.io and CRATES_IO_BOOTSTRAP_TOKEN is missing: OIDC cannot create a crate`
    );
  }
  return { token: bootstrapToken, source: 'bootstrap' };
}

/** Reminder that bootstrap-created crates still need a Trusted Publisher. */
export function formatBootstrapSummary(names) {
  return [
    `Created with CRATES_IO_BOOTSTRAP_TOKEN: ${names.join(', ')}.`,
    'Add a Trusted Publisher to each on crates.io (owner openooxml, repository betteroffice, workflow release.yml) before its next release.',
    'Remove CRATES_IO_BOOTSTRAP_TOKEN once no crate is missing.'
  ].join('\n');
}

function reportBootstrapCrates(names) {
  if (names.length === 0) return;
  const summary = formatBootstrapSummary(names);
  console.log(summary);
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${summary}\n`);
}

async function crateExists(name) {
  const response = await fetchRegistry(
    `https://crates.io/api/v1/crates/${encodeURIComponent(name)}`
  );
  return response.status !== 404;
}

async function crateVersion(name, version) {
  const response = await fetchRegistry(
    `https://crates.io/api/v1/crates/${encodeURIComponent(name)}/${encodeURIComponent(version)}`
  );
  if (response.status === 404) return null;
  return (await response.json()).version;
}

async function assertCrateOwnership(name) {
  const response = await fetchRegistry(
    `https://crates.io/api/v1/crates/${encodeURIComponent(name)}/owners`
  );
  if (response.status === 404) throw new Error(`${name} has no crates.io owners`);
  const owners = await response.json();
  if (!owners.users?.some((owner) => owner.login === EXPECTED_OWNER)) {
    throw new Error(`${name} is not owned by ${EXPECTED_OWNER}`);
  }
}

function sparseIndexPath(name) {
  const normalized = name.toLowerCase();
  if (normalized.length === 1) return `1/${normalized}`;
  if (normalized.length === 2) return `2/${normalized}`;
  if (normalized.length === 3) return `3/${normalized[0]}/${normalized}`;
  return `${normalized.slice(0, 2)}/${normalized.slice(2, 4)}/${normalized}`;
}

async function indexHasVersion(name, version) {
  const response = await fetchRegistry(`https://index.crates.io/${sparseIndexPath(name)}`);
  if (response.status === 404) return false;
  const entries = (await response.text())
    .trim()
    .split('\n')
    .filter(Boolean)
    .map((line) => JSON.parse(line));
  return entries.some((entry) => entry.vers === version && !entry.yanked);
}

async function waitFor(description, predicate) {
  const deadline = Date.now() + WAIT_TIMEOUT_MS;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    console.log(`Waiting for ${description}...`);
    await sleep(WAIT_INTERVAL_MS);
  }
  throw new Error(`Timed out waiting for ${description}`);
}

async function waitForRegistry(name, version) {
  await waitFor(`${name}@${version} on crates.io`, async () => {
    const found = await crateVersion(name, version);
    if (found?.yanked) throw new Error(`${name}@${version} is yanked`);
    return found !== null;
  });
  await assertCrateOwnership(name);
  await waitFor(`${name}@${version} in the sparse index`, () => indexHasVersion(name, version));
}

function publishDryRun() {
  for (const crate of RUST_PUBLISH_CRATES) {
    run('cargo', [
      'package',
      '--no-verify',
      '--exclude-lockfile',
      '--allow-dirty',
      '--locked',
      '-p',
      crate.name
    ]);
  }
}

async function publish() {
  const version = rustReleaseVersion();
  const packages = validateRustTrain(cargoMetadata(), version);

  if (process.argv.includes('--dry-run')) {
    publishDryRun();
    return;
  }
  if (version === '0.0.0') {
    console.log('Rust release train is unreleased; skipping crates.io publication.');
    return;
  }

  const createdWithBootstrap = [];
  for (const crate of RUST_PUBLISH_CRATES) {
    const existing = await crateVersion(crate.name, version);
    if (existing) {
      if (existing.yanked) throw new Error(`${crate.name}@${version} is yanked`);
      console.log(`${crate.name}@${version} is already published.`);
      await waitForRegistry(crate.name, version);
      continue;
    }

    const internalDependencies = packages
      .get(crate.name)
      .dependencies.filter((dependency) =>
        RUST_PUBLISH_CRATES.some((crate) => crate.name === dependency.name)
      );
    for (const dependency of internalDependencies) {
      await waitForRegistry(dependency.name, version);
    }

    const { token, source } = selectPublishToken({
      name: crate.name,
      exists: await crateExists(crate.name),
      oidcToken: process.env.CARGO_REGISTRY_TOKEN,
      bootstrapToken: process.env.CRATES_IO_BOOTSTRAP_TOKEN
    });
    console.log(`${crate.name}: publishing with ${source}.`);
    const result = run(
      'cargo',
      ['publish', '--locked', '--registry', 'crates-io', '-p', crate.name],
      { allowFailure: true, env: { CARGO_REGISTRY_TOKEN: token } }
    );
    if (result.status !== 0 && !(await crateVersion(crate.name, version))) {
      throw new Error(`Failed to publish ${crate.name}@${version}`);
    }
    await waitForRegistry(crate.name, version);
    if (source === 'bootstrap') createdWithBootstrap.push(crate.name);
  }
  reportBootstrapCrates(createdWithBootstrap);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await publish();
}
