// Packs each given package folder dry, and each platform package folder under its `npm/`, and
// fails for every tarball without the project's LICENSE at its root. No binary is needed:
//   node scripts/check-packed-license.mjs <package directory>...
import { spawnSync } from 'node:child_process';
import { existsSync, readdirSync, realpathSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const LICENSE = 'LICENSE';

/** The folders that become tarballs for `directory`: itself, then its platform packages. */
export function tarballFolders(directory) {
  const platforms = join(directory, 'npm');
  if (!existsSync(platforms)) return [directory];
  const children = readdirSync(platforms, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(join(platforms, entry.name, 'package.json')))
    .map((entry) => join(platforms, entry.name))
    .sort();
  return [directory, ...children];
}

/** The file paths in what `npm pack --json` printed for one package. */
export function packedPaths(stdout) {
  const report = JSON.parse(stdout);
  const files = (Array.isArray(report) ? report[0] : report)?.files;
  if (!Array.isArray(files)) throw new Error('npm pack printed no file list');
  return files.map((file) => file.path);
}

function pack(folder) {
  const result = spawnSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
    cwd: folder,
    encoding: 'utf8',
    timeout: 120_000,
  });
  if (result.status !== 0) {
    throw new Error(`npm pack failed in ${folder}: ${(result.stderr ?? '').trim().split('\n').pop() ?? result.error?.message}`);
  }
  return packedPaths(result.stdout);
}

function main() {
  const directories = process.argv.slice(2);
  if (!directories.length) throw new Error('usage: node scripts/check-packed-license.mjs <package directory>...');
  const failures = [];
  let checked = 0;
  for (const directory of directories.map((entry) => resolve(entry))) {
    for (const folder of tarballFolders(directory)) {
      checked += 1;
      const paths = pack(folder);
      if (!paths.includes(LICENSE)) {
        failures.push(`${relative(process.cwd(), folder) || '.'}: the tarball has no ${LICENSE} (${paths.length} files packed)`);
      }
    }
  }
  if (failures.length) {
    console.error(`${failures.length} of ${checked} tarballs lack the license:\n`);
    for (const failure of failures) console.error(`- ${failure}`);
    process.exit(1);
  }
  console.log(`${checked} tarballs carry ${LICENSE}.`);
}

const invoked = process.argv[1] ? realpathSync(process.argv[1]) : null;
if (invoked === realpathSync(fileURLToPath(import.meta.url))) main();
