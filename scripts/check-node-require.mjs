// Loads every JavaScript export of the given packages in plain Node the three ways a host can:
// require() from a CommonJS file, a static import from an ES module, and import() from a
// CommonJS file, each in its own process. Run after the packages are built:
//   node scripts/check-node-require.mjs <package directory>...
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, realpathSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { importPlan } from './check-node-import.mjs';

const require = createRequire(import.meta.url);
const { satisfies, subset, validRange } = createRequire(require.resolve('@changesets/cli'))('semver');

/** The Node versions where require() of an ES module works without a flag. */
export const REQUIRE_ESM_NODE = '^20.19.0 || >=22.12.0';
const WITHOUT_REQUIRE_ESM = ['16.20.2', '18.20.8', '19.9.0', '20.18.3', '21.7.3', '22.11.0'];

function rootTarget(manifest) {
  let entry = manifest.exports;
  if (entry && typeof entry === 'object' && !Array.isArray(entry) && '.' in entry) entry = entry['.'];
  while (entry && typeof entry === 'object' && !Array.isArray(entry)) entry = entry.require ?? entry.node ?? entry.default;
  return typeof entry === 'string' ? entry : manifest.main ?? 'index.js';
}

/** Whether a CommonJS `require()` of the package root loads an ES module. */
export function rootIsModule(directory, manifest) {
  const file = resolve(directory, rootTarget(manifest));
  if (extname(file) === '.mjs') return true;
  if (extname(file) === '.cjs') return false;
  for (let folder = dirname(file); folder.startsWith(resolve(directory)); folder = dirname(folder)) {
    const scope = join(folder, 'package.json');
    if (existsSync(scope)) return JSON.parse(readFileSync(scope, 'utf8')).type === 'module';
    if (folder === resolve(directory)) break;
  }
  return false;
}

/** How `manifest.engines.node` admits Node versions where a `require()` of its ES module root fails. */
export function requireFloorProblems(manifest, rootModule = true) {
  const range = manifest.engines?.node;
  if (!rootModule || typeof range !== 'string') return [];
  if (!validRange(range)) return [`${manifest.name}: engines.node "${range}" is not a valid range`];
  if (subset(range, REQUIRE_ESM_NODE)) return [];
  const example = WITHOUT_REQUIRE_ESM.find((version) => satisfies(version, range));
  return [`${manifest.name}: engines.node "${range}" admits ${example ? `Node ${example}` : 'versions'} outside ${REQUIRE_ESM_NODE}, where require() of an ES module fails with ERR_REQUIRE_ESM`];
}

const LOADED = 'betteroffice-entry-loaded';
const done = `process.stdout.write(${JSON.stringify(LOADED)});`;

/** How a host loads an entry: the name of the form, and the node arguments that do it. */
export const FORMS = {
  'require() from CommonJS': (specifier) => ['--input-type=commonjs', '--eval', `require(${JSON.stringify(specifier)}); ${done}`],
  'a static import': (specifier) => ['--input-type=module', '--eval', `import ${JSON.stringify(specifier)}; ${done}`],
  'import() from CommonJS': (specifier) => ['--input-type=commonjs', '--eval', `import(${JSON.stringify(specifier)}).then(() => { ${done} });`],
};

/** The Node error code in `stderr`, e.g. `ERR_PACKAGE_PATH_NOT_EXPORTED`, if it names one. */
export function errorCode(stderr) {
  return stderr.match(/code: '([A-Z0-9_]+)'/)?.[1] ?? stderr.match(/\b(ERR_[A-Z0-9_]+)\b/)?.[1];
}

/**
 * The exports of `manifest` whose `import` target a CommonJS `require()` cannot reach: no
 * `default` condition names the same file. Node resolves `require()` through `require`,
 * `node`, `module-sync` and `default`, never through `import`.
 */
export function importOnlyExports(manifest) {
  const exportsField = manifest.exports;
  if (!exportsField || typeof exportsField !== 'object' || Array.isArray(exportsField)) return [];
  const entries = Object.keys(exportsField).some((key) => !key.startsWith('.')) ? [['.', exportsField]] : Object.entries(exportsField);
  return entries
    .filter(([, value]) => value && typeof value === 'object' && typeof value.import === 'string' && value.default !== value.import)
    .map(([subpath]) => manifest.name + subpath.slice(1));
}

function load(directory, form, specifier) {
  const result = spawnSync(process.execPath, ['--unhandled-rejections=strict', ...FORMS[form](specifier)], {
    cwd: directory,
    encoding: 'utf8',
    timeout: 120_000,
    env: { ...process.env, NODE_OPTIONS: '' },
  });
  if (result.status === 0 && result.stdout.endsWith(LOADED)) return null;
  const stderr = result.stderr ?? '';
  const reason = errorCode(stderr) ?? result.error?.message ?? `exit ${result.status ?? result.signal}`;
  const line = stderr.split('\n').find((text) => /Error/.test(text))?.trim() ?? '';
  return `${specifier}: ${form} failed with ${reason}${line ? `\n    ${line.slice(0, 300)}` : ''}`;
}

function main() {
  if (typeof Bun !== 'undefined' || !process.versions.node) {
    throw new Error('check-node-require.mjs must run under node');
  }
  const directories = process.argv.slice(2);
  if (!directories.length) throw new Error('usage: node scripts/check-node-require.mjs <package directory>...');
  const failures = [];
  let checked = 0;
  for (const directory of directories.map((entry) => resolve(entry))) {
    const manifest = JSON.parse(readFileSync(join(directory, 'package.json'), 'utf8'));
    const plan = importPlan(manifest);
    failures.push(...plan.failures, ...requireFloorProblems(manifest, rootIsModule(directory, manifest)));
    const entries = plan.entries.filter((entry) => !entry.json);
    if (!entries.length) failures.push(`${manifest.name}: no package entries found to load`);
    for (const { specifier } of entries) {
      checked += 1;
      for (const form of Object.keys(FORMS)) {
        const failure = load(directory, form, specifier);
        if (failure) failures.push(failure);
      }
    }
  }
  if (failures.length) {
    console.error(`${failures.length} loads failed on Node ${process.version}:\n`);
    for (const failure of failures) console.error(`- ${failure}`);
    process.exit(1);
  }
  console.log(`${checked} package entries load by require(), a static import and import() on Node ${process.version}.`);
}

const invoked = process.argv[1] ? realpathSync(process.argv[1]) : null;
if (invoked === realpathSync(fileURLToPath(import.meta.url))) main();
