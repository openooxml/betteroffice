import { readFileSync, writeFileSync } from 'node:fs';
import {
  RUST_CRATES,
  WORKSPACE_MANIFEST,
  cargoMetadata,
  run,
  rustReleaseVersion,
  validateRustTrain
} from './rust-crates.mjs';

const PYTHON_BINDINGS = ['bindings/python-xlsx'];

function releaseManifest(binding) {
  return `${binding}/package.json`;
}

function cargoManifest(binding) {
  return `${binding}/Cargo.toml`;
}

function pythonReleaseVersion(binding) {
  return JSON.parse(readFileSync(releaseManifest(binding), 'utf8')).version;
}

function packageVersion(binding, source) {
  const section = source.match(/\[package\]\n([\s\S]*?)(?=\n\[|$)/);
  const version = section?.[1].match(/^version = "([^"]+)"$/m)?.[1];
  if (!version) throw new Error(`${binding} package.version is missing`);
  return version;
}

// pyproject reads its version from Cargo.toml, so this is the only file to rewrite.
function synchronizePythonVersion(binding, source, from, to) {
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

const checkOnly = process.argv.includes('--check');
const before = rustReleaseVersion();
const cargoBefore = readFileSync(WORKSPACE_MANIFEST, 'utf8');
const pythonBefore = new Map(
  PYTHON_BINDINGS.map((binding) => [binding, pythonReleaseVersion(binding)])
);

for (const binding of PYTHON_BINDINGS) {
  const marker = pythonBefore.get(binding);
  const locked = packageVersion(binding, readFileSync(cargoManifest(binding), 'utf8'));
  if (locked !== marker) {
    throw new Error(`${binding} changeset marker is ${marker}, but Cargo is ${locked}`);
  }
}
if (workspaceVersion(cargoBefore) !== before) {
  throw new Error(`Rust changeset marker is ${before}, but Cargo is ${workspaceVersion(cargoBefore)}`);
}

if (checkOnly) {
  const simulated = synchronizeCargoVersion(cargoBefore, before, '999.999.999');
  if (workspaceVersion(simulated) !== '999.999.999') {
    throw new Error('Cargo release train version synchronization failed');
  }
  for (const binding of PYTHON_BINDINGS) {
    const source = readFileSync(cargoManifest(binding), 'utf8');
    const simulatedPython = synchronizePythonVersion(
      binding,
      source,
      pythonBefore.get(binding),
      '999.999.999'
    );
    if (packageVersion(binding, simulatedPython) !== '999.999.999') {
      throw new Error(`${binding} version synchronization failed`);
    }
  }
  validate(before, true);
  console.log(`Rust release train is synchronized at ${before}.`);
  for (const binding of PYTHON_BINDINGS) {
    console.log(`${binding} is synchronized at ${pythonBefore.get(binding)}.`);
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

const pythonAfter = new Map(
  PYTHON_BINDINGS.map((binding) => [binding, pythonReleaseVersion(binding)])
);
for (const binding of PYTHON_BINDINGS) {
  const from = pythonBefore.get(binding);
  const to = pythonAfter.get(binding);
  if (to === from) continue;
  writeFileSync(
    cargoManifest(binding),
    synchronizePythonVersion(binding, readFileSync(cargoManifest(binding), 'utf8'), from, to)
  );
}

validate(after, true);
for (const binding of PYTHON_BINDINGS) {
  const from = pythonBefore.get(binding);
  const to = pythonAfter.get(binding);
  console.log(
    to === from ? `${binding} remains at ${to}.` : `Synchronized ${binding} ${from} -> ${to}.`
  );
}
console.log(
  after === before
    ? `Rust release train remains at ${after}.`
    : `Synchronized Rust release train ${before} -> ${after}.`
);
