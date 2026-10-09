import { execFile, spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { promisify } from 'node:util';

const execute = promisify(execFile);
const [format, sourceArg, outputArg] = process.argv.slice(2);
if (!['xlsx', 'pptx'].includes(format) || !sourceArg || !outputArg)
  throw new Error('usage: build-format.mjs xlsx|pptx source-checkout output-directory');
const source = resolve(sourceArg);
const output = resolve(outputArg);
const build = resolve(output, 'build');
await mkdir(resolve(build, 'src'), { recursive: true });
await copyFile(new URL(`./native/${format}.rs`, import.meta.url), resolve(build, 'src/main.rs'));
await copyFile(resolve(source, 'Cargo.lock'), resolve(build, 'Cargo.lock'));
await writeFile(resolve(build, 'Cargo.toml'), `[package]
name = "${format}-native-bench"
version = "0.0.0"
edition = "2024"
publish = false

[workspace]

[dependencies]
betteroffice-${format} = { path = ${JSON.stringify(resolve(source, `crates/betteroffice-${format}`))}, default-features = false${format === 'pptx' ? ', features = ["raster"]' : ''} }
serde_json = "1"
sha2 = "0.10"

[profile.release]
opt-level = 3
lto = "thin"
`);
const target = resolve(process.env.CARGO_TARGET_DIR ?? resolve(output, 'target'));
await new Promise((accept, reject) => {
  const child = spawn('cargo', ['build', '--release', '--manifest-path', resolve(build, 'Cargo.toml')], {
    stdio: 'inherit', env: { ...process.env, CARGO_TARGET_DIR: target },
  });
  child.on('error', reject);
  child.on('exit', code => code === 0 ? accept() : reject(new Error(`${format} build failed: ${code}`)));
});
const binary = resolve(output, `${format}-native-bench`);
await copyFile(resolve(target, `release/${format}-native-bench`), binary);
const sha = data => createHash('sha256').update(data).digest('hex');
const metadata = {
  source_sha: (await execute('git', ['-C', source, 'rev-parse', 'HEAD'])).stdout.trim(),
  binary_sha256: sha(await readFile(binary)),
  harness_sha256: sha(await readFile(resolve(build, 'src/main.rs'))),
  rustc: (await execute('rustc', ['--version'])).stdout.trim(),
  profile: 'release opt-level=3 lto=thin',
};
await writeFile(resolve(output, 'build.json'), JSON.stringify(metadata, null, 2) + '\n');
console.log(JSON.stringify(metadata));
