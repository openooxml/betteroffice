import {
  RUST_CRATES,
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
  for (const crate of RUST_CRATES) {
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

  for (const crate of RUST_CRATES) {
    const existing = await crateVersion(crate.name, version);
    if (existing) {
      if (existing.yanked) throw new Error(`${crate.name}@${version} is yanked`);
      console.log(`${crate.name}@${version} is already published.`);
      await waitForRegistry(crate.name, version);
      continue;
    }

    const internalDependencies = packages
      .get(crate.name)
      .dependencies.filter((dependency) => packages.has(dependency.name));
    for (const dependency of internalDependencies) {
      await waitForRegistry(dependency.name, version);
    }

    if (!process.env.CARGO_REGISTRY_TOKEN) {
      throw new Error('CARGO_REGISTRY_TOKEN is required to publish Rust crates');
    }

    const result = run(
      'cargo',
      ['publish', '--locked', '--registry', 'crates-io', '-p', crate.name],
      { allowFailure: true }
    );
    if (result.status !== 0 && !(await crateVersion(crate.name, version))) {
      throw new Error(`Failed to publish ${crate.name}@${version}`);
    }
    await waitForRegistry(crate.name, version);
  }
}

await publish();
