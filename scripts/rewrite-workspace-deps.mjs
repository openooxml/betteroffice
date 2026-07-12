// Pin `workspace:` dependency ranges to concrete versions before publishing.
//
// Neither `changeset version` nor `changeset publish` rewrites the workspace
// protocol in this Bun workspace, so an unresolved `workspace:*` would ship in
// the tarball and break `npm install` for consumers. Run this on the publish
// path only (ephemeral CI checkout) — source keeps `workspace:*`.
//
// Replacement follows the pnpm/bun convention: `workspace:*` → exact version,
// `workspace:^`/`workspace:~` → caret/tilde range.
import { readdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const PKG_DIR = 'packages';

const versions = new Map();
const manifests = [];
for (const name of readdirSync(PKG_DIR)) {
  const path = join(PKG_DIR, name, 'package.json');
  if (!existsSync(path)) continue;
  const json = JSON.parse(readFileSync(path, 'utf8'));
  if (json.name && json.version) versions.set(json.name, json.version);
  manifests.push({ path, json });
}

function resolve(dep, range) {
  const version = versions.get(dep);
  if (!version) return range;
  const suffix = range.slice('workspace:'.length);
  if (suffix === '*' || suffix === '') return version;
  if (suffix === '^' || suffix === '~') return suffix + version;
  return suffix; // workspace:1.2.3 → 1.2.3
}

let changed = 0;
for (const { path, json } of manifests) {
  for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
    const deps = json[field];
    if (!deps) continue;
    for (const [dep, range] of Object.entries(deps)) {
      if (typeof range === 'string' && range.startsWith('workspace:')) {
        deps[dep] = resolve(dep, range);
        changed++;
      }
    }
  }
  writeFileSync(path, JSON.stringify(json, null, 2) + '\n');
}

console.log(`Pinned ${changed} workspace dependency range(s) for publish.`);
