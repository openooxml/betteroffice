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
