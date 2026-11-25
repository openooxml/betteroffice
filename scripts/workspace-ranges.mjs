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

const semver = createRequire(changesetsEntry())('semver');

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
      .filter((manifest) => existsSync(join(root, manifest)))
      .sort();
  });

  return manifests.map((manifest) => {
    const source = readFileSync(join(root, manifest), 'utf8');
    return { manifest, source, json: JSON.parse(source) };
  });
}

export function explicitWorkspaceRanges(root = process.cwd()) {
  const recorded = [];
  for (const { manifest, json } of workspacePackages(root)) {
    for (const section of DEPENDENCY_FIELDS) {
      for (const [name, value] of Object.entries(json[section] ?? {})) {
        if (typeof value !== 'string' || !value.startsWith(WORKSPACE_PREFIX)) continue;
        const range = value.slice(WORKSPACE_PREFIX.length);
        if (!range.trim() || BARE_RANGES.has(range) || !semver.validRange(range)) continue;
        recorded.push({ manifest, section, name, range });
      }
    }
  }
  return recorded;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function restoreRange(source, { manifest, section, name, range }, current) {
  const sectionPattern = new RegExp(
    `^  "${section}"[ \\t]*:[ \\t]*\\{[ \\t]*\\r?\\n[\\s\\S]*?^  \\}`,
    'm'
  );
  const key = escapeRegExp(JSON.stringify(name));
  const value = escapeRegExp(JSON.stringify(current));
  const dependencyPattern = new RegExp(
    `(^    ${key}[ \\t]*:[ \\t]*)${value}(?=[ \\t]*(?:,[ \\t]*)?\\r?$)`,
    'm'
  );
  const original = JSON.stringify(WORKSPACE_PREFIX + range);
  const updated = source.replace(sectionPattern, (block) =>
    block.replace(dependencyPattern, (_match, prefix) => prefix + original)
  );
  if (updated === source) throw new Error(`Could not restore ${manifest} ${section} ${name}.`);
  return updated;
}

export function restoreWorkspaceRanges(recorded, root = process.cwd()) {
  if (recorded.length === 0) return [];

  const packages = workspacePackages(root);
  const manifests = new Map(packages.map((pkg) => [pkg.manifest, pkg]));
  const versions = new Map(packages.map(({ json }) => [json.name, json.version]));
  const updates = new Map();
  const restored = [];

  for (const entry of recorded) {
    const { manifest, section, name, range } = entry;
    const pkg = manifests.get(manifest);
    const current = pkg?.json[section]?.[name];
    if (current === WORKSPACE_PREFIX + range) continue;

    const version = versions.get(name);
    if (!semver.satisfies(version, range)) {
      throw new Error(
        `${manifest} ${section} ${name} declares workspace:${range}, which does not include the new version ${version ?? '(missing)'}.`
      );
    }
    if (typeof current !== 'string') {
      throw new Error(`Could not find ${manifest} ${section} ${name} to restore.`);
    }
    updates.set(manifest, restoreRange(updates.get(manifest) ?? pkg.source, entry, current));
    restored.push(entry);
  }

  for (const [manifest, source] of updates) {
    writeFileSync(join(root, manifest), source);
  }
  return restored;
}
