import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
export const backstop = 'betteroffice-base-latin';

export function cargoTree(manifest, ...args) {
  return spawnSync('cargo', ['tree', '--locked', '--manifest-path', manifest, '--target', 'all', ...args], {
    cwd: root,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
}

export function assertNoBackstop(manifest, crate, ...args) {
  const tree = cargoTree(manifest, '-p', crate, '-e', 'features,normal,build', '--prefix', 'none', ...args);
  assert.equal(tree.status, 0, `cargo tree failed for ${crate}:\n${tree.stderr}`);
  const lines = tree.stdout.split('\n');
  assert.ok(lines.some((line) => line.startsWith('betteroffice-ooxml-text ')), `${crate} no longer resolves betteroffice-ooxml-text`);
  const offending = lines.filter((line) => line.startsWith(`${backstop} `) || line.includes('feature "base-latin"'));
  assert.deepEqual([...new Set(offending)], [], `${crate} links the bundled Latin faces, which only the Rust facades may ship`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const control = cargoTree('Cargo.toml', '-p', 'betteroffice-docx-raster', '-e', 'features,normal', '-i', backstop);
  assert.equal(control.status, 0, `the guard's control no longer finds ${backstop} under docx-raster:\n${control.stderr}`);
  assert.match(control.stdout, /betteroffice-ooxml-text feature "base-latin"/, `the guard's control no longer finds the base-latin feature:\n${control.stdout}`);

  for (const binding of ['betteroffice-docx-native', 'betteroffice-pptx-native']) {
    assertNoBackstop('bindings/Cargo.toml', binding);
    console.log(`${binding}: no ${backstop}`);
  }
}
