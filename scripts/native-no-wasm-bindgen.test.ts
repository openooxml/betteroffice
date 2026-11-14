import { describe, expect, test } from 'bun:test';
import { resolve } from 'node:path';
import { NATIVE_PACKAGES } from './native-bindings.mjs';

const ROOT = resolve(import.meta.dir, '..');
const HOST = /^host: (\S+)$/m.exec(Bun.spawnSync(['rustc', '-vV']).stdout.toString())?.[1] ?? '';

function packageName(path: string): string {
  return `betteroffice-${path.replace(/^bindings\//, '')}`;
}

describe('native bindings compile no wasm-bindgen', () => {
  test('the host target is known', () => expect(HOST).not.toBe(''));

  for (const path of NATIVE_PACKAGES) {
    test(path, () => {
      const tree = Bun.spawnSync([
        'cargo', 'tree', '--locked', '--manifest-path', 'bindings/Cargo.toml', '-p', packageName(path),
        '--target', HOST, '-e', 'normal', '--prefix', 'none', '-f', '{p}',
      ], { cwd: ROOT });
      expect(tree.exitCode, tree.stderr.toString()).toBe(0);
      const found = new Set(tree.stdout.toString().split('\n').filter((line) => /^(wasm-bindgen|js-sys|web-sys)\b/.test(line)));
      expect([...found]).toEqual([]);
    }, 120_000);
  }
});
