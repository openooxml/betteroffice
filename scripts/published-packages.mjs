import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export { PYPI_DISTRIBUTIONS } from './python-bindings.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const WORKSPACES = ['apps', 'bindings', 'packages'];

function workspaceManifests() {
  const manifests = [];
  for (const workspace of WORKSPACES) {
    for (const entry of readdirSync(join(ROOT, workspace))) {
      try {
        manifests.push(JSON.parse(readFileSync(join(ROOT, workspace, entry, 'package.json'), 'utf8')));
      } catch {
        continue;
      }
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

/** The npm packages a release uploads. */
export function publishedPackages() {
  return workspaceManifests()
    .filter((manifest) => !manifest.private && manifest.name?.startsWith('@betteroffice/'))
    .map((manifest) => manifest.name)
    .sort();
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
