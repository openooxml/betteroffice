import { readFileSync, writeFileSync } from 'node:fs';
import {
  RUST_CRATES,
  WORKSPACE_MANIFEST,
  cargoMetadata,
  run,
  rustReleaseVersion,
  validateRustTrain
} from './rust-crates.mjs';

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
if (workspaceVersion(cargoBefore) !== before) {
  throw new Error(`Rust changeset marker is ${before}, but Cargo is ${workspaceVersion(cargoBefore)}`);
}

if (checkOnly) {
  const simulated = synchronizeCargoVersion(cargoBefore, before, '999.999.999');
  if (workspaceVersion(simulated) !== '999.999.999') {
    throw new Error('Cargo release train version synchronization failed');
  }
  validate(before, true);
  console.log(`Rust release train is synchronized at ${before}.`);
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

validate(after, true);
console.log(
  after === before
    ? `Rust release train remains at ${after}.`
    : `Synchronized Rust release train ${before} -> ${after}.`
);
