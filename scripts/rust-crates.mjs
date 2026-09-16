import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

export const RUST_RELEASE_MANIFEST = 'crates/package.json';
export const WORKSPACE_MANIFEST = 'Cargo.toml';
export const STANDALONE_WORKSPACES = ['bindings', 'fuzz', 'apps/native-viewer'];

export const RUST_CRATES = [
  { name: 'betteroffice-opc', dependency: 'ooxml-opc' },
  { name: 'betteroffice-ooxml-text', dependency: 'ooxml-text' },
  { name: 'betteroffice-drawingml', dependency: 'ooxml-drawingml' },
  { name: 'betteroffice-xlsx-model', dependency: 'xlsx-model' },
  { name: 'betteroffice-xlsx-parse', dependency: 'xlsx-parse' },
  { name: 'betteroffice-xlsx-calc', dependency: 'xlsx-calc' },
  { name: 'betteroffice-xlsx-render', dependency: 'xlsx-render' },
  { name: 'betteroffice-xlsx-ops', dependency: 'xlsx-ops' },
  { name: 'betteroffice-xlsx-raster', dependency: 'xlsx-raster' },
  { name: 'betteroffice-xlsx', dependency: 'betteroffice-xlsx' },
  { name: 'betteroffice-docx-parse', dependency: 'docx-parse' },
  { name: 'betteroffice-docx-layout', dependency: 'docx-layout' },
  { name: 'betteroffice-docx-raster', dependency: 'docx-raster' },
  { name: 'betteroffice-docx-edit', dependency: 'docx-edit' },
  { name: 'betteroffice-docx', dependency: 'betteroffice-docx' },
  { name: 'betteroffice-pptx-parse', dependency: 'pptx-parse' },
  { name: 'betteroffice-vsdx-parse', dependency: 'vsdx-parse', publish: false },
  { name: 'betteroffice-vsdx-formula', dependency: 'vsdx-formula', publish: false },
  { name: 'betteroffice-vsdx-resolve', dependency: 'vsdx-resolve', publish: false },
  { name: 'betteroffice-vsdx-eval', dependency: 'vsdx-eval', publish: false },
  { name: 'betteroffice-vsdx-render', dependency: 'vsdx-render', publish: false },
  { name: 'betteroffice-vsdx-edit', dependency: 'vsdx-edit', publish: false },
  { name: 'betteroffice-vsdx', dependency: 'betteroffice-vsdx', publish: false },
  { name: 'betteroffice-pptx-edit', dependency: 'pptx-edit' },
  { name: 'betteroffice-pptx-render', dependency: 'pptx-render' },
  { name: 'betteroffice-pptx-raster', dependency: 'pptx-raster' },
  { name: 'betteroffice-pptx', dependency: 'betteroffice-pptx' }
];

export const RUST_PUBLISH_CRATES = RUST_CRATES.filter((crate) => crate.publish !== false);

export function rustReleaseVersion() {
  return JSON.parse(readFileSync(RUST_RELEASE_MANIFEST, 'utf8')).version;
}

export function run(command, args, { capture = false, allowFailure = false, env } = {}) {
  let childEnv = process.env;
  if (env) {
    childEnv = { ...process.env };
    for (const [key, value] of Object.entries(env)) {
      if (value === undefined) delete childEnv[key];
      else childEnv[key] = value;
    }
  }
  const result = spawnSync(command, args, {
    encoding: capture ? 'utf8' : undefined,
    stdio: capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
    // cargo metadata output exceeds the 1 MiB default.
    maxBuffer: 64 * 1024 * 1024,
    env: childEnv
  });
  if (result.error) throw result.error;
  if (result.status !== 0 && !allowFailure) {
    if (capture && result.stderr) process.stderr.write(result.stderr);
    throw new Error(`${command} ${args.join(' ')} exited with ${result.status}`);
  }
  return result;
}

export function cargoMetadata({ locked = true, manifestPath, env } = {}) {
  const args = ['metadata', '--format-version', '1'];
  if (manifestPath) args.push('--manifest-path', manifestPath);
  if (locked) args.push('--locked');
  const result = run('cargo', args, { capture: true, ...(env ? { env } : {}) });
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
    const registries = crate.publish === false ? [] : ['crates-io'];
    if (JSON.stringify(pkg.publish) !== JSON.stringify(registries)) {
      throw new Error(`${crate.name} must declare publish = ${JSON.stringify(registries)}`);
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
