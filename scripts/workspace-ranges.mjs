import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';

const require = createRequire(import.meta.url);
const WORKSPACE_PREFIX = 'workspace:';
const DEPENDENCY_FIELDS = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'];
const BARE_RANGES = new Set(['*', '^', '~']);

function changesetsEntry() {
  try {
    return require.resolve('@changesets/cli/package.json');
  } catch (error) {
    if (error.code !== 'ERR_PACKAGE_PATH_NOT_EXPORTED') throw error;
    return require.resolve('@changesets/cli');
  }
}

let semver;

function getSemver() {
  return (semver ??= createRequire(changesetsEntry())('semver'));
}

function explicitWorkspaceRange(value, semver) {
  if (typeof value !== 'string' || !value.startsWith(WORKSPACE_PREFIX)) return;
  const range = value.slice(WORKSPACE_PREFIX.length);
  if (!range.trim() || BARE_RANGES.has(range) || !semver.validRange(range)) return;
  return range;
}

function workspacePackages(root) {
  const { workspaces } = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
  if (!Array.isArray(workspaces)) throw new Error('Expected a workspaces array in package.json.');

  const manifests = workspaces.flatMap((pattern) => {
    if (typeof pattern !== 'string' || !/^[\w-]+(?:\/\*)?$/.test(pattern)) {
      throw new Error(`Unsupported workspace pattern: ${JSON.stringify(pattern)}.`);
    }
    if (!pattern.endsWith('/*')) return [`${pattern}/package.json`];

    const directory = pattern.slice(0, -2);
    return readdirSync(join(root, directory), { withFileTypes: true })
      .filter((entry) => entry.isDirectory() && !entry.name.startsWith('.'))
      .map((entry) => `${directory}/${entry.name}/package.json`)
      .sort();
  });

  return manifests
    .filter((manifest) => existsSync(join(root, manifest)))
    .map((manifest) => {
      const source = readFileSync(join(root, manifest), 'utf8');
      return { manifest, source, json: JSON.parse(source) };
    });
}

export function explicitWorkspaceRanges(root = process.cwd()) {
  const semver = getSemver();
  const recorded = [];
  for (const { manifest, json } of workspacePackages(root)) {
    for (const section of DEPENDENCY_FIELDS) {
      for (const [name, value] of Object.entries(json[section] ?? {})) {
        const range = explicitWorkspaceRange(value, semver);
        if (range === undefined) continue;
        recorded.push({ manifest, section, name, range });
      }
    }
  }
  return recorded;
}

export function restoreWorkspaceRanges(recorded, root = process.cwd()) {
  if (recorded.length === 0) return [];

  const semver = getSemver();
  const packages = workspacePackages(root);
  const manifests = new Map(packages.map((pkg) => [pkg.manifest, pkg]));
  const versions = new Map(packages.map(({ json }) => [json.name, json.version]));
  const updates = new Set();
  const restored = [];

  for (const entry of recorded) {
    const { manifest, section, name, range } = entry;
    const version = versions.get(name);
    if (!versions.has(name) || !semver.satisfies(version, range)) {
      throw new Error(
        `${manifest} ${section} ${name} declares workspace:${range}, which does not include the new version ${version ?? '(missing)'}.`
      );
    }

    const pkg = manifests.get(manifest);
    const current = pkg?.json[section]?.[name];
    const original = WORKSPACE_PREFIX + range;
    if (current === original) continue;
    if (current === undefined) {
      throw new Error(`Could not find ${manifest} ${section} ${name} to restore.`);
    }
    if (explicitWorkspaceRange(current, semver) === undefined) {
      const value = typeof current === 'string' ? current : JSON.stringify(current);
      throw new Error(`${manifest} ${section} ${name} changed to ${value}, not a workspace range`);
    }
    pkg.json[section][name] = original;
    updates.add(pkg);
    restored.push(entry);
  }

  for (const { manifest, source, json } of updates) {
    const indent = source.match(/^[ \t]+/m)?.[0] ?? '  ';
    writeFileSync(join(root, manifest), JSON.stringify(json, null, indent) + (source.endsWith('\n') ? '\n' : ''));
  }
  return restored;
}
