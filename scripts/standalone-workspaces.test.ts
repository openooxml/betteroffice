import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { STANDALONE_WORKSPACES } from './rust-crates.mjs';

const repository = fileURLToPath(new URL('..', import.meta.url));
const skipped = new Set(['target', 'node_modules', 'vendor']);

function manifests(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    if (entry.isDirectory()) return skipped.has(entry.name) ? [] : manifests(join(dir, entry.name));
    return entry.name === 'Cargo.toml' ? [join(dir, entry.name)] : [];
  });
}

describe('standalone Cargo workspaces', () => {
  for (const workspace of STANDALONE_WORKSPACES) {
    test(`${workspace} commits the lockfile the release script refreshes`, () => {
      expect(existsSync(join(repository, workspace, 'Cargo.lock'))).toBe(true);
    });

    test(`${workspace} depends on workspace crates by path alone`, () => {
      for (const manifest of manifests(join(repository, workspace))) {
        const pinned = readFileSync(manifest, 'utf8')
          .split('\n')
          .filter((line) => /path = "[^"]*crates\//.test(line) && /\bversion = /.test(line));
        expect({ manifest, pinned }).toEqual({ manifest, pinned: [] });
      }
    });
  }
});
