import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export { PYPI_DISTRIBUTIONS } from './python-bindings.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));

function readManifest(directory) {
  try {
    return JSON.parse(readFileSync(join(ROOT, directory, 'package.json'), 'utf8'));
  } catch {
    return undefined;
  }
}

/** Every workspace manifest, from the same globs Bun and Changesets expand. */
function workspaceManifests() {
  const manifests = [];
  for (const pattern of readManifest('.').workspaces) {
    if (pattern.includes('*') && !pattern.endsWith('/*')) {
      throw new Error(`published-packages.mjs cannot expand the workspace glob ${pattern}`);
    }
    const parent = pattern.endsWith('/*') ? pattern.slice(0, -2) : undefined;
    const directories = parent
      ? readdirSync(join(ROOT, parent)).map((entry) => join(parent, entry))
      : [pattern];
    for (const directory of directories) {
      const manifest = readManifest(directory);
      if (manifest) manifests.push(manifest);
    }
  }
  return manifests;
}

/** Every `@betteroffice` package name in the repository, published or not. */
export function workspacePackages() {
  return workspaceManifests()
    .map((manifest) => manifest.name)
    .filter((name) => name?.startsWith('@betteroffice/'))
    .sort();
}

/**
 * The npm packages a release uploads, with the version it would upload: every
 * non-private workspace, which is exactly the set `changeset publish` takes.
 */
export function publishedPackageVersions() {
  return workspaceManifests()
    .filter((manifest) => !manifest.private && manifest.name && manifest.version)
    .map((manifest) => ({ name: manifest.name, version: manifest.version }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** The npm packages a release uploads. */
export function publishedPackages() {
  return publishedPackageVersions().map((entry) => entry.name);
}

/** The crates a release uploads to crates.io. */
export function publishedCrates() {
  const dir = join(ROOT, 'crates');
  const names = [];
  for (const entry of readdirSync(dir)) {
    let manifest;
    try {
      manifest = readFileSync(join(dir, entry, 'Cargo.toml'), 'utf8');
    } catch {
      continue;
    }
    if (!manifest.includes('publish = ["crates-io"]')) continue;
    const name = manifest.match(/^name = "([^"]+)"/m)?.[1];
    if (name) names.push(name);
  }
  return names.sort();
}
