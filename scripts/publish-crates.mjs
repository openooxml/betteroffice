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

/** Initial token choice per crate. Fallback after an OIDC failure is orchestrated separately. */
export function selectPublishToken({ name, exists, oidcToken, bootstrapToken }) {
  if (exists) {
    if (oidcToken) return { token: oidcToken, source: 'oidc' };
    if (bootstrapToken) return { token: bootstrapToken, source: 'bootstrap-fallback' };
    throw new Error(
      `${name} is on crates.io but CARGO_REGISTRY_TOKEN and CRATES_IO_BOOTSTRAP_TOKEN are both missing`
    );
  }
  if (!bootstrapToken) {
    throw new Error(
      `${name} is not on crates.io and CRATES_IO_BOOTSTRAP_TOKEN is missing: OIDC cannot create a crate`
    );
  }
  return { token: bootstrapToken, source: 'bootstrap' };
}

/** Child env for `cargo publish`: only the selected token, under both aliases. */
export function cargoPublishEnv(selectedToken) {
  return {
    CARGO_REGISTRY_TOKEN: selectedToken,
    CARGO_REGISTRIES_CRATES_IO_TOKEN: selectedToken,
    CRATES_IO_BOOTSTRAP_TOKEN: undefined
  };
}

/** Child env for cargo invocations that need no credentials. */
export function cargoNoAuthEnv() {
  return {
    CARGO_REGISTRY_TOKEN: undefined,
    CARGO_REGISTRIES_CRATES_IO_TOKEN: undefined,
    CRATES_IO_BOOTSTRAP_TOKEN: undefined
  };
}

function appendStepSummary(text) {
  console.log(text);
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${text}\n`);
}

/** Reminder covering both bootstrap uses. Existing fallback is not a new crate. */
export function formatBootstrapSummary(created, fallback = []) {
  let createdNames;
  let fallbackNames;
  if (created && typeof created === 'object' && !Array.isArray(created)) {
    createdNames = created.created ?? [];
    fallbackNames = created.fallback ?? [];
  } else {
    createdNames = created ?? [];
    fallbackNames = fallback ?? [];
  }
  const lines = [];
  if (createdNames.length > 0) {
    lines.push(`Created with CRATES_IO_BOOTSTRAP_TOKEN: ${createdNames.join(', ')}.`);
  }
  if (fallbackNames.length > 0) {
    lines.push(
      `Published with CRATES_IO_BOOTSTRAP_TOKEN fallback (existing crates, OIDC unavailable or failed): ${fallbackNames.join(', ')}.`
    );
  }
  lines.push(
    'Add a Trusted Publisher to each on crates.io (owner openooxml, repository betteroffice, workflow release.yml) before its next release.'
  );
  lines.push(
    'Remove CRATES_IO_BOOTSTRAP_TOKEN once OIDC works for all crates, not just once all names exist.'
  );
  return lines.join('\n');
}

export function recordBootstrapUse(name, kind) {
  const text =
    kind === 'bootstrap'
      ? `Created ${name} with CRATES_IO_BOOTSTRAP_TOKEN.`
      : `Published ${name} with CRATES_IO_BOOTSTRAP_TOKEN fallback (existing crate).`;
  appendStepSummary(text);
}

function reportAuthSummary(created, fallback) {
  if (created.length === 0 && fallback.length === 0) return;
  appendStepSummary(formatBootstrapSummary(created, fallback));
}

/**
 * Publish one crate, trying OIDC first for existing crates and falling back
 * to the bootstrap token. New crates go straight to the bootstrap token.
 * `checkVersion` must throw on registry lookup errors, never resolve missing.
 */
export async function attemptPublishWithFallback({
  name,
  version,
  exists,
  oidcToken,
  bootstrapToken,
  checkVersion,
  runPublish
}) {
  if (!exists) {
    if (!bootstrapToken) {
      throw new Error(
        `${name} is not on crates.io and CRATES_IO_BOOTSTRAP_TOKEN is missing: OIDC cannot create a crate`
      );
    }
    const result = await runPublish(bootstrapToken, 'bootstrap');
    if (result.status !== 0) {
      const found = await checkVersion();
      if (!found) throw new Error(`Failed to publish ${name}@${version} with the bootstrap token`);
      if (found.yanked) throw new Error(`${name}@${version} is yanked`);
      return { source: 'bootstrap' };
    }
    return { source: 'bootstrap' };
  }

  if (oidcToken) {
    const oidcResult = await runPublish(oidcToken, 'oidc');
    if (oidcResult.status === 0) return { source: 'oidc' };
    const found = await checkVersion();
    if (found) {
      if (found.yanked) throw new Error(`${name}@${version} is yanked`);
      return { source: 'oidc' };
    }
    if (!bootstrapToken) {
      throw new Error(
        `Failed to publish ${name}@${version} with OIDC and CRATES_IO_BOOTSTRAP_TOKEN is missing: cannot fall back`
      );
    }
    const fallbackResult = await runPublish(bootstrapToken, 'bootstrap-fallback');
    if (fallbackResult.status !== 0) {
      const retried = await checkVersion();
      if (!retried) {
        throw new Error(
          `Failed to publish ${name}@${version} with OIDC and with the bootstrap token fallback`
        );
      }
      if (retried.yanked) throw new Error(`${name}@${version} is yanked`);
      return { source: 'bootstrap-fallback' };
    }
    return { source: 'bootstrap-fallback' };
  }

  if (!bootstrapToken) {
    throw new Error(
      `${name} is on crates.io but CARGO_REGISTRY_TOKEN and CRATES_IO_BOOTSTRAP_TOKEN are both missing`
    );
  }
  const direct = await runPublish(bootstrapToken, 'bootstrap-fallback');
  if (direct.status !== 0) {
    const found = await checkVersion();
    if (!found) {
      throw new Error(
        `Failed to publish ${name}@${version} with the bootstrap token fallback (OIDC unavailable)`
      );
    }
    if (found.yanked) throw new Error(`${name}@${version} is yanked`);
    return { source: 'bootstrap-fallback' };
  }
  return { source: 'bootstrap-fallback' };
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

function publishDryRun(env) {
  for (const crate of RUST_PUBLISH_CRATES) {
    run(
      'cargo',
      [
        'package',
        '--no-verify',
        '--exclude-lockfile',
        '--allow-dirty',
        '--locked',
        '-p',
        crate.name
      ],
      { env }
    );
  }
}

async function publish() {
  const oidcToken = process.env.CARGO_REGISTRY_TOKEN;
  const bootstrapToken = process.env.CRATES_IO_BOOTSTRAP_TOKEN;
  const noAuthEnv = cargoNoAuthEnv();
  const version = rustReleaseVersion();
  const packages = validateRustTrain(cargoMetadata({ env: noAuthEnv }), version);

  if (process.argv.includes('--dry-run')) {
    publishDryRun(noAuthEnv);
    return;
  }
  if (version === '0.0.0') {
    console.log('Rust release train is unreleased; skipping crates.io publication.');
    return;
  }

  const createdWithBootstrap = [];
  const fallbackWithBootstrap = [];
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

    const exists = await crateExists(crate.name);
    const { source } = await attemptPublishWithFallback({
      name: crate.name,
      version,
      exists,
      oidcToken,
      bootstrapToken,
      checkVersion: () => crateVersion(crate.name, version),
      runPublish: (token) =>
        run('cargo', ['publish', '--locked', '--registry', 'crates-io', '-p', crate.name], {
          allowFailure: true,
          env: cargoPublishEnv(token)
        })
    });
    console.log(`${crate.name}: publishing with ${source}.`);
    if (source === 'bootstrap') {
      createdWithBootstrap.push(crate.name);
      recordBootstrapUse(crate.name, 'bootstrap');
    } else if (source === 'bootstrap-fallback') {
      fallbackWithBootstrap.push(crate.name);
      recordBootstrapUse(crate.name, 'bootstrap-fallback');
    }
    await waitForRegistry(crate.name, version);
  }
  reportAuthSummary(createdWithBootstrap, fallbackWithBootstrap);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await publish();
}
