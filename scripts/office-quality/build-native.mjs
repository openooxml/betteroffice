import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { promisify } from 'node:util';

const execute = promisify(execFile);
const harness = resolve(dirname(fileURLToPath(import.meta.url)), 'native/main.rs');

export function nativeFeatures(layoutSource) {
  const signature = layoutSource.match(/pub fn register_substitute_measure_font\s*\(([^)]*)\)/s)?.[1];
  if (!signature) return [];
  return [/\bbold\s*:/.test(signature) ? 'substitute-styles' : 'substitute-metrics'];
}

export async function buildNative(source, output) {
  source = resolve(source);
  output = resolve(output);
  const build = resolve(output, 'build');
  await mkdir(resolve(build, 'src'), { recursive: true });
  await copyFile(harness, resolve(build, 'src/main.rs'));
  const harnessHash = createHash('sha256').update(await readFile(resolve(build, 'src/main.rs'))).digest('hex');
  await copyFile(resolve(source, 'Cargo.lock'), resolve(build, 'Cargo.lock'));
  const names = ['docx-edit', 'docx-layout', 'docx-parse', 'docx-raster', 'ooxml-text'];
  const dependencies = await Promise.all(names.map(async (name) => {
    const path = resolve(source, 'crates', name);
    const manifest = await readFile(resolve(path, 'Cargo.toml'), 'utf8');
    const features = /^tiff\s*=/m.test(manifest) ? ', features = ["tiff"]' : '';
    return `${name} = { package = ${JSON.stringify(`betteroffice-${name}`)}, path = ${JSON.stringify(path)}${features} }`;
  }));
  const features = nativeFeatures(await readFile(resolve(source, 'crates/docx-layout/src/lib.rs'), 'utf8'));
  await writeFile(resolve(build, 'Cargo.toml'), `[package]
name = "docx-native-bench"
version = "0.0.0"
edition = "2024"
publish = false

[workspace]

[features]
substitute-metrics = []
substitute-styles = ["substitute-metrics"]

[dependencies]
${dependencies.join('\n')}
serde_json = "1"
sha2 = "0.10"

[profile.release]
opt-level = 3
lto = "thin"
`);
  const target = resolve(process.env.CARGO_TARGET_DIR ?? resolve(output, 'target'));
  const args = ['build', '--release', '--manifest-path', resolve(build, 'Cargo.toml')];
  if (features.length) args.push('--features', features.join(','));
  const { spawn } = await import('node:child_process');
  await new Promise((accept, reject) => {
    const child = spawn('cargo', args, { stdio: 'inherit', env: { ...process.env, CARGO_TARGET_DIR: target } });
    child.on('error', reject);
    child.on('exit', (code) => code === 0 ? accept() : reject(new Error(`Native build failed: ${code}`)));
  });
  const binary = resolve(output, 'docx-native-bench');
  await copyFile(resolve(target, 'release/docx-native-bench'), binary);
  const sha = (await execute('git', ['-C', source, 'rev-parse', 'HEAD'])).stdout.trim();
  const metadata = {
    source_sha: sha,
    binary_sha256: createHash('sha256').update(await readFile(binary)).digest('hex'),
    harness_sha256: harnessHash,
    rustc: (await execute('rustc', ['--version'])).stdout.trim(),
    features, profile: 'release opt-level=3 lto=thin',
  };
  await writeFile(resolve(output, 'build.json'), JSON.stringify(metadata, null, 2) + '\n');
  return metadata;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [source, output] = process.argv.slice(2);
  if (!source || !output) throw new Error('usage: build-native.mjs source-checkout output-directory');
  console.log(JSON.stringify(await buildNative(source, output)));
}
