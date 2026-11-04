// Resolves every export of every library package with the webpack that Next.js bundles, through a
// static import, import() and require.resolveWeak (what next/dynamic emits). Run after
// `bun run build:packages`; pass package directories to check other packages instead.
import { readFileSync, realpathSync } from 'node:fs';
import { mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { libraryPackageManifests } from './published-packages.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));

const ENTRIES = {
  static: (specifiers) =>
    specifiers.map((specifier, index) => `import * as entry${index} from ${JSON.stringify(specifier)};`).join('\n') +
    `\nexport default [${specifiers.map((_, index) => `entry${index}`).join(', ')}];\n`,
  lazy: (specifiers) =>
    `export default () => [${specifiers.map((specifier) => `import(${JSON.stringify(specifier)})`).join(', ')}];\n`,
  weak: (specifiers) =>
    `module.exports = () => [${specifiers.map((specifier) => `require.resolveWeak(${JSON.stringify(specifier)})`).join(', ')}];\n`,
};

/** The import specifier of every export subpath in `manifest`, wildcard patterns left out. */
export function exportSpecifiers(manifest) {
  const map = manifest.exports;
  if (!map) return [];
  const conditional =
    typeof map === 'string' || Array.isArray(map) || Object.keys(map).some((key) => !key.startsWith('.'));
  return (conditional ? ['.'] : Object.keys(map))
    .filter((subpath) => !subpath.includes('*'))
    .map((subpath) => manifest.name + subpath.slice(1));
}

/** The library packages, or the packages in `directories` when any are given. */
export function packageManifests(directories = []) {
  if (!directories.length) return libraryPackageManifests();
  return directories.map((directory) => {
    const absolute = resolve(ROOT, directory);
    return { directory: absolute, manifest: JSON.parse(readFileSync(join(absolute, 'package.json'), 'utf8')) };
  });
}

/** Links each package into `directory`/node_modules, as a workspace install does. */
export async function linkPackages(directory, packages) {
  for (const { directory: source, manifest } of packages) {
    const target = join(directory, 'node_modules', ...manifest.name.split('/'));
    await mkdir(dirname(target), { recursive: true });
    await symlink(source, target, 'dir');
  }
}

/** The webpack that Next.js bundles, as `next build --webpack` runs it. */
export function nextWebpack() {
  const compiled = createRequire(join(ROOT, 'apps/demo/package.json'))('next/dist/compiled/webpack/webpack');
  if (typeof compiled.init === 'function') compiled.init();
  if (typeof compiled.webpack !== 'function') throw new Error('Next.js ships no bundled webpack');
  return compiled.webpack;
}

async function resolveAll(specifiers, packages) {
  const directory = await mkdtemp(join(tmpdir(), 'webpack-exports-'));
  try {
    await linkPackages(directory, packages);
    await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true }));
    for (const [name, source] of Object.entries(ENTRIES)) {
      await writeFile(join(directory, `${name}.js`), source(specifiers));
    }
    const wanted = new Set(specifiers);
    const resolved = new Map();
    const record = {
      apply(compiler) {
        compiler.hooks.normalModuleFactory.tap('check-webpack-exports', (factory) => {
          factory.hooks.afterResolve.tap('check-webpack-exports', (data) => {
            if (!wanted.has(data.request)) return undefined;
            resolved.set(`${basename(data.contextInfo.issuer, '.js')} ${data.request}`, data.createData.resource);
            return false;
          });
        });
      },
    };
    const webpack = nextWebpack();
    const stats = await new Promise((done, fail) =>
      webpack(
        {
          mode: 'production',
          target: 'web',
          devtool: false,
          context: directory,
          entry: Object.fromEntries(Object.keys(ENTRIES).map((name) => [name, `./${name}.js`])),
          output: { path: join(directory, 'dist') },
          optimization: { minimize: false },
          plugins: [record],
        },
        (error, result) => (error ? fail(error) : done(result))
      )
    );
    const failures = stats.compilation.errors.map((error) => error.message.split('\n').slice(0, 3).join(' '));
    for (const entry of Object.keys(ENTRIES)) {
      for (const specifier of specifiers) {
        if (!resolved.has(`${entry} ${specifier}`)) failures.push(`${specifier}: not resolved from the ${entry} entry`);
      }
    }
    return { failures, resolved: resolved.size };
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

async function main() {
  const packages = packageManifests(process.argv.slice(2));
  const specifiers = packages.flatMap(({ manifest }) => exportSpecifiers(manifest));
  if (!specifiers.length) throw new Error('no package exports to resolve');
  const { failures, resolved } = await resolveAll(specifiers, packages);
  if (failures.length) {
    console.error(`${failures.length} package exports failed to resolve with webpack:\n`);
    for (const failure of failures) console.error(`- ${failure}`);
    process.exit(1);
  }
  console.log(
    `${specifiers.length} exports of ${packages.length} packages resolve with webpack ` +
      `(${resolved} static, import() and require.resolveWeak references).`
  );
}

const invoked = process.argv[1] ? realpathSync(process.argv[1]) : null;
if (invoked === realpathSync(fileURLToPath(import.meta.url))) await main();
