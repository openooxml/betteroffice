// Drives the wasm-pack builds vendored into packages/*/src/wasm/generated. The
// predev/prebuild hooks, the package build scripts and CI each invoke these back
// to back, and wasm-pack rewrites its output unconditionally, so every module is
// fingerprinted: its wasm-opt pass alone costs minutes under the workspace's
// `lto = true` release profile.
import { spawnSync } from 'node:child_process';
import { createHash, type Hash } from 'node:crypto';
import { copyFile, mkdir, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export interface WasmModule {
  /** Directory under crates/. */
  crate: string;
  /** wasm-bindgen artifact base name: the crate name with dashes underscored. */
  name: string;
  /** Destination for the vendored output, relative to the repo root. */
  generated: string;
  /** Flags passed through to cargo. */
  cargoArgs?: string[];
}

// `source` is what wasm-pack left in target/, `vendored` what was copied into the
// package. Both are recorded so a target/ restored from a CI cache can be
// re-vendored — the expensive part is wasm-opt, not the copy.
interface Stamp {
  key: string;
  source: Record<string, string>;
  vendored: Record<string, string>;
}

const WASM_PACK_VERSION = '0.15.0';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function requireWasmPack(): void {
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
}

async function hashTree(hash: Hash, dir: string, prefix: string): Promise<void> {
  const entries = await readdir(dir, { withFileTypes: true });
  entries.sort((left, right) => (left.name < right.name ? -1 : 1));
  for (const entry of entries) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      await hashTree(hash, path, `${prefix}${entry.name}/`);
      continue;
    }
    hash.update(`${prefix}${entry.name}\0`);
    hash.update(await readFile(path));
  }
}

// Everything cargo lets you redirect codegen with from outside the manifest.
// CARGO_ENCODED_RUSTFLAGS takes precedence over RUSTFLAGS, and any
// CARGO_PROFILE_* override rewrites the release profile the wasm is built under.
const CODEGEN_ENV = [
  'RUSTFLAGS',
  'CARGO_ENCODED_RUSTFLAGS',
  'CARGO_BUILD_RUSTFLAGS',
  'CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUSTFLAGS',
];

/** Digest of every input the emitted wasm depends on: the crates, the workspace
 * manifest (release profiles), the toolchain, the cargo environment and these
 * scripts. Shared by every module, so it doubles as the CI cache key. */
export async function sourcesFingerprint(): Promise<string> {
  const hash = createHash('sha256');
  // `stable` moves, so the toolchain has to be part of the identity.
  const rustc = spawnSync('rustc', ['-vV'], { encoding: 'utf8' });
  if (rustc.status !== 0) throw new Error('rustc -vV failed');
  hash.update(rustc.stdout);
  hash.update(`${WASM_PACK_VERSION}\0`);
  const overrides = Object.keys(process.env)
    .filter((name) => name.startsWith('CARGO_PROFILE_'))
    .sort();
  for (const name of [...CODEGEN_ENV, ...overrides]) {
    hash.update(`${name}=${process.env[name] ?? ''}\0`);
  }
  for (const file of ['Cargo.toml', 'Cargo.lock', '.cargo/config.toml']) {
    hash.update(`${file}\0`);
    hash.update(await readFile(resolve(root, file)).catch(() => Buffer.alloc(0)));
  }
  await hashTree(hash, resolve(root, 'crates'), 'crates/');
  const scripts = resolve(root, 'scripts');
  const names = (await readdir(scripts)).filter((name) => name.includes('wasm')).sort();
  for (const name of names) {
    hash.update(`scripts/${name}\0`);
    hash.update(await readFile(join(scripts, name)));
  }
  return hash.digest('hex');
}

async function digest(path: string): Promise<string | null> {
  const bytes = await readFile(path).catch(() => null);
  return bytes ? createHash('sha256').update(bytes).digest('hex') : null;
}

async function hashAll(dir: string, files: string[]): Promise<Record<string, string>> {
  const hashes: Record<string, string> = {};
  for (const file of files) hashes[file] = (await digest(resolve(dir, file))) as string;
  return hashes;
}

async function intact(dir: string, expected: Record<string, string> | undefined): Promise<boolean> {
  const entries = Object.entries(expected ?? {});
  if (entries.length === 0) return false;
  for (const [file, hash] of entries) {
    if ((await digest(resolve(dir, file))) !== hash) return false;
  }
  return true;
}

/** Builds every module whose inputs or vendored output changed. */
export async function buildWasmModules(modules: WasmModule[]): Promise<void> {
  requireWasmPack();
  const sources = await sourcesFingerprint();

  for (const { crate, name, generated, cargoArgs = [] } of modules) {
    const dest = resolve(root, generated);
    const output = resolve(root, 'target/wasm-pack', crate);
    const stamp = resolve(root, 'target/wasm-pack', `${crate}.json`);
    const files = [`${name}.js`, `${name}.d.ts`, `${name}_bg.wasm`, `${name}_bg.wasm.d.ts`];
    const key = createHash('sha256')
      .update(sources)
      .update(`\0${crate}\0${name}\0${cargoArgs.join(' ')}`)
      .digest('hex');

    const recorded: Stamp | null = await readFile(stamp, 'utf8')
      .then((text) => JSON.parse(text))
      .catch(() => null);
    const current = recorded?.key === key;
    if (current && (await intact(dest, recorded?.vendored))) {
      console.log(`[wasm] ${crate}: up to date`);
      continue;
    }

    if (current && (await intact(output, recorded?.source))) {
      console.log(`[wasm] ${crate}: vendoring the cached build`);
    } else {
      await rm(output, { recursive: true, force: true });
      // --locked must ride with the cargo pass-through: wasm-pack forwards its
      // own trailing args verbatim once a `--` section exists, and cargo rejects
      // a stray `--` marker.
      const build = spawnSync(
        'wasm-pack',
        [
          'build',
          resolve(root, 'crates', crate),
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
    }

    await mkdir(dest, { recursive: true });
    const glue = await readFile(resolve(output, `${name}.js`), 'utf8');
    const fallback = `module_or_path = new URL('${name}_bg.wasm', import.meta.url);`;
    if (!glue.includes(fallback)) throw new Error(`wasm-pack glue fallback changed (${name})`);
    await writeFile(
      resolve(dest, `${name}.js`),
      glue.replace(fallback, `throw new Error('${crate} requires an explicit module or URL');`)
    );
    for (const file of files.filter((file) => file !== `${name}.js`)) {
      await copyFile(resolve(output, file), resolve(dest, file));
    }

    await mkdir(dirname(stamp), { recursive: true });
    await writeFile(
      stamp,
      JSON.stringify({
        key,
        source: await hashAll(output, files),
        vendored: await hashAll(dest, files),
      } satisfies Stamp)
    );
  }
}
