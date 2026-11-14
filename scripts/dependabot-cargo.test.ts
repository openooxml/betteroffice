import { expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const config = Bun.YAML.parse(readFileSync(`${root}.github/dependabot.yml`, 'utf8')) as any;
const cargo = config.updates.filter((entry: any) => entry['package-ecosystem'] === 'cargo');

function lockfileDirectories(): string[] {
  const listed = Bun.spawnSync(['git', 'ls-files', '-z', '*Cargo.lock'], { cwd: root });
  expect(listed.success).toBe(true);
  return listed.stdout
    .toString()
    .split('\0')
    .filter((file) => file && !file.split('/').includes('vendor'))
    .map((file) => `/${file.replace(/\/?Cargo\.lock$/, '')}`)
    .sort();
}

test('one Dependabot cargo entry covers every Cargo lockfile, so a bump moves them together', () => {
  expect(cargo).toHaveLength(1);
  expect([...cargo[0].directories].sort()).toEqual(lockfileDirectories());
});

test('cargo minor, patch and major bumps each arrive as one pull request per dependency across those lockfiles', () => {
  const groups = cargo[0].groups;
  expect(groups['cargo-minor-patch']).toEqual({ 'update-types': ['minor', 'patch'], 'group-by': 'dependency-name' });
  expect(groups['cargo-major']).toEqual({ 'update-types': ['major'], 'group-by': 'dependency-name' });
});

test('every dependency edge in each cargo lockfile resolves to exactly one locked package', () => {
  const problems: string[] = [];
  for (const directory of [...cargo[0].directories].sort()) {
    const lockfile = `${directory}/Cargo.lock`.replace(/^\/+/, '');
    const packages = (Bun.TOML.parse(readFileSync(`${root}${lockfile}`, 'utf8')) as any).package as any[];
    for (const pkg of packages) {
      for (const dependency of (pkg.dependencies ?? []) as string[]) {
        const [name, version] = dependency.split(' ');
        const matches = packages.filter((p) => p.name === name && (version === undefined || p.version === version));
        if (version === undefined && matches.length !== 1) {
          problems.push(`${lockfile}: ${pkg.name} depends on "${dependency}" but ${matches.length} versions are locked`);
        } else if (version !== undefined && matches.length === 0) {
          problems.push(`${lockfile}: ${pkg.name} depends on "${dependency}" which is not locked`);
        }
      }
    }
  }
  expect(problems).toEqual([]);
});

function workspaceMembers(directory: string): string[] | undefined {
  const manifest = `${root}${`${directory}/Cargo.toml`.replace(/^\/+/, '')}`;
  if (!existsSync(manifest)) return undefined;
  return (Bun.TOML.parse(readFileSync(manifest, 'utf8')) as any).workspace?.members;
}

function globToRegExp(pattern: string): RegExp {
  const escaped = pattern
    .split('*')
    .map((part) => part.replace(/[.+?^${}()|[\]\\]/g, '\\$&'))
    .join('[^/]*');
  return new RegExp(`^${escaped}$`);
}

test('workspace member patterns can be listed by Dependabot', () => {
  const problems: string[] = [];
  for (const directory of [...cargo[0].directories].sort()) {
    for (const pattern of workspaceMembers(directory) ?? []) {
      const last = pattern.split('/').pop() ?? '';
      if (last.includes('*') && !last.startsWith('*')) {
        problems.push(`${directory}: member pattern "${pattern}" has text before a "*" in its last segment`);
      }
    }
  }
  expect(problems).toEqual([]);
});

test('every crate under bindings is a workspace member', () => {
  const members = workspaceMembers('/bindings') ?? [];
  const crates = readdirSync(`${root}bindings`, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .filter((name) => {
      const manifest = `${root}bindings/${name}/Cargo.toml`;
      return existsSync(manifest) && (Bun.TOML.parse(readFileSync(manifest, 'utf8')) as any).package !== undefined;
    })
    .sort();
  expect(crates.length).toBeGreaterThan(0);
  const unlisted = crates.filter((name) => !members.some((pattern) => globToRegExp(pattern).test(name)));
  expect(unlisted).toEqual([]);
});
