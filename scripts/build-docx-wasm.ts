// Builds the four docx wasm cores (container / layout / edit / parse) and
// vendors the wasm-pack output into packages/docx/src/wasm/generated/.
// The glue .js/.d.ts are committed; the *_bg.wasm binaries are gitignored and
// rebuilt on demand (predev/prebuild hooks), mirroring scripts/build-xlsx-wasm.ts
// so the repo never carries multi-MB binaries in history.
import { copyFile, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

interface WasmModuleBuild {
  crate: string;
  name: string;
  dir: string;
  cargoArgs: string[];
}

const WASM_PACK_VERSION = '0.15.0';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const version = spawnSync('wasm-pack', ['--version'], { encoding: 'utf8' });
const versionErrorCode =
  version.error && 'code' in version.error && typeof version.error.code === 'string'
    ? version.error.code
    : undefined;
if (versionErrorCode === 'ENOENT') {
  throw new Error(
    `wasm-pack ${WASM_PACK_VERSION} is required; install it with cargo install wasm-pack --version ${WASM_PACK_VERSION} --locked`
  );
}
if (version.status !== 0) process.exit(version.status ?? 1);
if (version.stdout.trim() !== `wasm-pack ${WASM_PACK_VERSION}`) {
  throw new Error(`expected wasm-pack ${WASM_PACK_VERSION}, got ${version.stdout.trim()}`);
}

// name = wasm-bindgen artifact base name (crate name with dashes underscored);
// dir = subdirectory under packages/docx/src/wasm/generated/;
// cargoArgs = extra flags after `--` (docx-edit keeps its wasm-bindgen boundary
// behind the `wasm` cargo feature so native builds never pull that stack).
const MODULES: WasmModuleBuild[] = [
  { crate: 'ooxml-opc', name: 'ooxml_opc', dir: 'opc', cargoArgs: [] },
  { crate: 'docx-layout', name: 'docx_layout', dir: 'layout', cargoArgs: [] },
  // --locked must ride with the cargo pass-through here: wasm-pack forwards its
  // own trailing args verbatim once a `--` section exists, and cargo rejects a
  // stray `--` marker.
  { crate: 'docx-edit', name: 'docx_edit', dir: 'edit', cargoArgs: ['--locked', '--features', 'wasm'] },
  { crate: 'docx-parse', name: 'docx_parse', dir: 'parse', cargoArgs: [] },
];

for (const { crate, name, dir, cargoArgs } of MODULES) {
  const crateDir = resolve(root, 'crates', crate);
  const output = resolve(root, 'target/wasm-pack/docx', dir);
  const generated = resolve(root, 'packages/docx/src/wasm/generated', dir);

  await rm(output, { recursive: true, force: true });
  const build = spawnSync(
    'wasm-pack',
    [
      'build',
      crateDir,
      '--release',
      '--target',
      'web',
      '--out-dir',
      output,
      ...(cargoArgs.length ? ['--', ...cargoArgs] : ['--locked']),
    ],
    { stdio: 'inherit' }
  );
  if (build.status !== 0) process.exit(build.status ?? 1);

  await mkdir(generated, { recursive: true });
  const gluePath = resolve(output, `${name}.js`);
  const glue = await readFile(gluePath, 'utf8');
  const fallback = `module_or_path = new URL('${name}_bg.wasm', import.meta.url);`;
  if (!glue.includes(fallback)) throw new Error(`wasm-pack glue fallback changed (${name})`);
  await writeFile(
    resolve(generated, `${name}.js`),
    glue.replace(fallback, `throw new Error('${crate} wasm requires an explicit module or URL');`)
  );

  for (const file of [`${name}.d.ts`, `${name}_bg.wasm`, `${name}_bg.wasm.d.ts`]) {
    await copyFile(resolve(output, file), resolve(generated, file));
  }
}
