import { spawnSync } from 'node:child_process';
import { copyFile, mkdir, stat } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { requireWasmOpt } from './wasm.ts';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
requireWasmOpt();

const build = spawnSync(
  'cargo',
  [
    'build',
    '-p',
    'fontpack',
    '--target',
    'wasm32-unknown-unknown',
    '--profile',
    'fonts-decoder',
    '--no-default-features',
    '--features',
    'wasm',
    '--locked',
    '--target-dir',
    resolve(root, 'target/fonts-decoder'),
  ],
  { cwd: root, stdio: 'inherit' },
);
if (build.error) throw build.error;
if (build.status !== 0) process.exit(build.status ?? 1);

const artifact = resolve(
  root,
  'target/fonts-decoder/wasm32-unknown-unknown/fonts-decoder/fontpack.wasm',
);
const optimized = resolve(root, 'target/fonts-decoder/fontpack.wasm');
const features = [
  '--enable-bulk-memory',
  '--enable-bulk-memory-opt',
  '--enable-nontrapping-float-to-int',
  '--enable-sign-ext',
  '--enable-mutable-globals',
  '--enable-reference-types',
  '--enable-multivalue',
];
const optimize = spawnSync('wasm-opt', ['-Oz', ...features, artifact, '-o', optimized], {
  cwd: root,
  stdio: 'inherit',
});
if (optimize.error) throw optimize.error;
if (optimize.status !== 0) process.exit(optimize.status ?? 1);

const destination = resolve(root, 'packages/fonts/dist/decoder/fontpack.wasm');
await mkdir(dirname(destination), { recursive: true });
await copyFile(optimized, destination);
const sourceDestination = resolve(root, 'packages/fonts/src/decoder/fontpack.wasm');
await mkdir(dirname(sourceDestination), { recursive: true });
await copyFile(optimized, sourceDestination);
console.log(`[wasm] fontpack: ${(await stat(destination)).size} bytes`);
