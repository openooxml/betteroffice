import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const backstop = 'betteroffice-base-latin';

function cargoTree(manifest, ...args) {
  return spawnSync('cargo', ['tree', '--locked', '--manifest-path', manifest, '--target', 'all', ...args], {
    cwd: root,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
}

const control = cargoTree('Cargo.toml', '-p', 'betteroffice-docx-raster', '-e', 'features,normal', '-i', backstop);
assert.equal(control.status, 0, `the guard's control no longer finds ${backstop} under docx-raster:\n${control.stderr}`);
assert.match(control.stdout, /betteroffice-ooxml-text feature "base-latin"/, `the guard's control no longer finds the base-latin feature:\n${control.stdout}`);

for (const binding of ['betteroffice-docx-native', 'betteroffice-pptx-native']) {
  const tree = cargoTree('bindings/Cargo.toml', '-p', binding, '-e', 'features,normal,build', '--prefix', 'none');
  assert.equal(tree.status, 0, `cargo tree failed for ${binding}:\n${tree.stderr}`);
  const lines = tree.stdout.split('\n');
  assert.ok(lines.some((line) => line.startsWith('betteroffice-ooxml-text ')), `${binding} no longer resolves betteroffice-ooxml-text`);
  const offending = lines.filter((line) => line.startsWith(`${backstop} `) || line.includes('feature "base-latin"'));
  assert.deepEqual([...new Set(offending)], [], `${binding} links the bundled Latin faces, which only the Rust facades may ship`);
  console.log(`${binding}: no ${backstop}`);
}
