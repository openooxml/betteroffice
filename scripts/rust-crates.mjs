import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

export const RUST_RELEASE_MANIFEST = 'crates/package.json';
export const WORKSPACE_MANIFEST = 'Cargo.toml';

export const RUST_CRATES = [
  { name: 'betteroffice-opc', dependency: 'ooxml-opc' },
  { name: 'betteroffice-xlsx-model', dependency: 'xlsx-model' },
  { name: 'betteroffice-xlsx-parse', dependency: 'xlsx-parse' },
  { name: 'betteroffice-xlsx-calc', dependency: 'xlsx-calc' },
  { name: 'betteroffice-xlsx-render', dependency: 'xlsx-render' },
  { name: 'betteroffice-xlsx-ops', dependency: 'xlsx-ops' },
  { name: 'betteroffice-xlsx-raster', dependency: 'xlsx-raster' },
  { name: 'betteroffice-xlsx', dependency: 'betteroffice-xlsx' }
];

export function rustReleaseVersion() {
  return JSON.parse(readFileSync(RUST_RELEASE_MANIFEST, 'utf8')).version;
}

export function run(command, args, { capture = false, allowFailure = false } = {}) {
  const result = spawnSync(command, args, {
    encoding: capture ? 'utf8' : undefined,
    stdio: capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
    env: process.env
  });
  if (result.error) throw result.error;
  if (result.status !== 0 && !allowFailure) {
    if (capture && result.stderr) process.stderr.write(result.stderr);
    throw new Error(`${command} ${args.join(' ')} exited with ${result.status}`);
  }
  return result;
}

export function cargoMetadata({ locked = true } = {}) {
  const args = ['metadata', '--format-version', '1'];
  if (locked) args.push('--locked');
  const result = run('cargo', args, { capture: true });
  return JSON.parse(result.stdout);
}

export function validateRustTrain(metadata, version) {
  const packages = new Map(metadata.packages.map((pkg) => [pkg.name, pkg]));
  const positions = new Map(RUST_CRATES.map((crate, index) => [crate.name, index]));
  const rustPackages = new Map();

  for (const [index, crate] of RUST_CRATES.entries()) {
    const pkg = packages.get(crate.name);
    if (!pkg) throw new Error(`Cargo package ${crate.name} is missing`);
    if (pkg.version !== version) {
      throw new Error(`${crate.name} is ${pkg.version}; expected ${version}`);
    }
    if (JSON.stringify(pkg.publish) !== '["crates-io"]') {
      throw new Error(`${crate.name} must publish only to crates-io`);
    }
    rustPackages.set(crate.name, pkg);
    for (const dependency of pkg.dependencies) {
      const dependencyIndex = positions.get(dependency.name);
      if (dependencyIndex !== undefined && dependencyIndex >= index) {
        throw new Error(`${crate.name} must follow ${dependency.name} in publish order`);
      }
    }
  }

  return rustPackages;
}
