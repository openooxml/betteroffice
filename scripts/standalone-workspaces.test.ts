import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { STANDALONE_WORKSPACES } from './rust-crates.mjs';

const repository = fileURLToPath(new URL('..', import.meta.url));
const crates = resolve(repository, 'crates');

function capture(command: string, args: string[]): string {
  const result = spawnSync(command, args, { cwd: repository, encoding: 'utf8' });
  if (result.status !== 0) throw new Error(`${command} ${args.join(' ')}: ${result.stderr}`);
  return result.stdout;
}

function committedLockfiles(): string[] {
  return capture('git', ['ls-files', '--', '*Cargo.lock'])
    .split('\n')
    .filter((path) => path && !path.includes('/vendor/'));
}

function workspaceMembers(workspace: string) {
  const manifest = `${workspace}/Cargo.toml`;
  const json = capture('cargo', ['metadata', '--format-version', '1', '--no-deps', '--manifest-path', manifest]);
  return JSON.parse(json).packages as {
    name: string;
    dependencies: { name: string; req: string; path?: string }[];
  }[];
}

function pinsWorkspaceCrate(dependency: { req: string; path?: string }): boolean {
  if (!dependency.path || dependency.req === '*') return false;
  return !relative(crates, dependency.path).startsWith('..');
}

describe('standalone Cargo workspaces', () => {
  test('every committed lockfile belongs to the root or a listed workspace', () => {
    const expected = ['Cargo.lock', ...STANDALONE_WORKSPACES.map((workspace) => `${workspace}/Cargo.lock`)];
    expect(committedLockfiles().sort()).toEqual(expected.sort());
  });

  for (const workspace of STANDALONE_WORKSPACES) {
    test(`${workspace} depends on workspace crates by path alone`, () => {
      const pinned = workspaceMembers(workspace).flatMap((member) =>
        member.dependencies
          .filter(pinsWorkspaceCrate)
          .map((dependency) => `${member.name} -> ${dependency.name} ${dependency.req}`)
      );
      expect(pinned).toEqual([]);
    });
  }
});
