import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
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

test('major cargo bumps arrive as one pull request per dependency across those lockfiles', () => {
  const groups = cargo[0].groups;
  expect(groups['cargo-minor-patch']['update-types']).toEqual(['minor', 'patch']);
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
