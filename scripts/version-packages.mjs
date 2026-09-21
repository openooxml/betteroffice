import { readFileSync, writeFileSync } from 'node:fs';
import { PYTHON_BINDINGS } from './python-bindings.mjs';
import { NODE_BINDINGS, synchronizeNodeLoader, validateNodeVersions } from './node-bindings.mjs';
import {
  RUST_CRATES,
  STANDALONE_WORKSPACES,
  WORKSPACE_MANIFEST,
  cargoMetadata,
  run,
  rustReleaseVersion,
  validateRustTrain
} from './rust-crates.mjs';

function releaseManifest(binding) {
  return `${binding}/package.json`;
}

function cargoManifest(binding) {
  return `${binding}/Cargo.toml`;
}

function bindingReleaseVersion(binding) {
  return JSON.parse(readFileSync(releaseManifest(binding), 'utf8')).version;
}

function packageVersion(binding, source) {
  const section = source.match(/\[package\]\n([\s\S]*?)(?=\n\[|$)/);
  const version = section?.[1].match(/^version = "([^"]+)"$/m)?.[1];
  if (!version) throw new Error(`${binding} package.version is missing`);
  return version;
}

function synchronizeBindingVersion(binding, source, from, to) {
  if (packageVersion(binding, source) !== from) {
    throw new Error(`${binding} is not at ${from}`);
  }
  return source.replace(/(\[package\]\n[\s\S]*?^version = ")[^"]+("$)/m, `$1${to}$2`);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function workspaceVersion(source) {
  const section = source.match(/\[workspace\.package\]\n([\s\S]*?)(?=\n\[|$)/);
  const version = section?.[1].match(/^version = "([^"]+)"$/m)?.[1];
  if (!version) throw new Error('workspace.package.version is missing');
  return version;
}

function synchronizeCargoVersion(source, from, to) {
  if (workspaceVersion(source) !== from) {
    throw new Error(`Cargo release train does not match ${from}`);
  }

  let updated = source.replace(
    /(\[workspace\.package\]\n[\s\S]*?^version = ")[^"]+("$)/m,
    `$1${to}$2`
  );

  for (const crate of RUST_CRATES) {
    const key = escapeRegExp(crate.dependency);
    const pattern = new RegExp(`^(${key} = \\{[^\\n]*version = ")[^"]+("[^\\n]*\\})$`, 'm');
    if (!pattern.test(updated)) {
      throw new Error(`workspace dependency ${crate.dependency} has no version`);
    }
    updated = updated.replace(pattern, `$1${to}$2`);
  }

  return updated;
}

function validate(version, locked) {
  const metadata = cargoMetadata({ locked });
  validateRustTrain(metadata, version);
}

// Nothing else rewrites these lockfiles, which pin every bumped crate by version.
function synchronizeStandaloneLocks() {
  for (const workspace of STANDALONE_WORKSPACES) {
    cargoMetadata({ locked: false, manifestPath: cargoManifest(workspace) });
    // Re-assert under `--locked`, the way CI reads the lock it just wrote.
    cargoMetadata({ manifestPath: cargoManifest(workspace) });
  }
}

function synchronizeBunLock() {
  run('bun', ['install', '--lockfile-only']);
}

const checkOnly = process.argv.includes('--check');
validateNodeVersions();
const before = rustReleaseVersion();
const cargoBefore = readFileSync(WORKSPACE_MANIFEST, 'utf8');
const bindings = [...PYTHON_BINDINGS, ...NODE_BINDINGS];
const bindingBefore = new Map(
  bindings.map((binding) => [binding, bindingReleaseVersion(binding)])
);

for (const binding of bindings) {
  const marker = bindingBefore.get(binding);
  const locked = packageVersion(binding, readFileSync(cargoManifest(binding), 'utf8'));
  if (locked !== marker) {
    throw new Error(`${binding} changeset marker is ${marker}, but Cargo is ${locked}`);
  }
}
if (workspaceVersion(cargoBefore) !== before) {
  throw new Error(
    `Rust changeset marker is ${before}, but Cargo is ${workspaceVersion(cargoBefore)}`
  );
}

if (checkOnly) {
  const simulated = synchronizeCargoVersion(cargoBefore, before, '999.999.999');
  if (workspaceVersion(simulated) !== '999.999.999') {
    throw new Error('Cargo release train version synchronization failed');
  }
  for (const binding of bindings) {
    const source = readFileSync(cargoManifest(binding), 'utf8');
    const simulatedBinding = synchronizeBindingVersion(
      binding,
      source,
      bindingBefore.get(binding),
      '999.999.999'
    );
    if (packageVersion(binding, simulatedBinding) !== '999.999.999') {
      throw new Error(`${binding} version synchronization failed`);
    }
  }
  validate(before, true);
  for (const workspace of STANDALONE_WORKSPACES) {
    cargoMetadata({ manifestPath: cargoManifest(workspace) });
  }
  console.log(`Rust release train is synchronized at ${before}.`);
  for (const binding of bindings) {
    console.log(`${binding} is synchronized at ${bindingBefore.get(binding)}.`);
  }
  for (const workspace of STANDALONE_WORKSPACES) {
    console.log(`${workspace}/Cargo.lock is current.`);
  }
  process.exit(0);
}

run('bun', ['run', 'changeset', 'version']);
const after = rustReleaseVersion();

if (after !== before) {
  writeFileSync(
    WORKSPACE_MANIFEST,
    synchronizeCargoVersion(readFileSync(WORKSPACE_MANIFEST, 'utf8'), before, after)
  );
  validate(after, false);
}

const bindingAfter = new Map(
  bindings.map((binding) => [binding, bindingReleaseVersion(binding)])
);
for (const binding of bindings) {
  const from = bindingBefore.get(binding);
  const to = bindingAfter.get(binding);
  if (to === from) continue;
  writeFileSync(
    cargoManifest(binding),
    synchronizeBindingVersion(binding, readFileSync(cargoManifest(binding), 'utf8'), from, to)
  );
}
for (const binding of NODE_BINDINGS) {
  run('bunx', ['napi', 'version', '--cwd', binding]);
  const loader = `${binding}/index.js`;
  writeFileSync(loader, synchronizeNodeLoader(
    readFileSync(loader, 'utf8'), bindingBefore.get(binding), bindingAfter.get(binding)
  ));
}

validateNodeVersions();
validate(after, true);
synchronizeStandaloneLocks();
synchronizeBunLock();
for (const binding of bindings) {
  const from = bindingBefore.get(binding);
  const to = bindingAfter.get(binding);
  console.log(
    to === from ? `${binding} remains at ${to}.` : `Synchronized ${binding} ${from} -> ${to}.`
  );
}
for (const workspace of STANDALONE_WORKSPACES) {
  console.log(`Synchronized ${workspace}/Cargo.lock.`);
}
console.log('Synchronized bun.lock.');
console.log(
  after === before
    ? `Rust release train remains at ${after}.`
    : `Synchronized Rust release train ${before} -> ${after}.`
);
