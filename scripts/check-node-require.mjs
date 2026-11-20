// Loads every JavaScript export of the given packages in plain Node the three ways a host can:
// require() from a CommonJS file, a static import from an ES module, and import() from a
// CommonJS file, each in its own process. Run after the packages are built:
//   node scripts/check-node-require.mjs <package directory>...
import { spawnSync } from 'node:child_process';
import { readFileSync, realpathSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { importPlan } from './check-node-import.mjs';

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
    failures.push(...plan.failures);
    for (const { specifier, json } of plan.entries) {
      if (json) continue;
      checked += 1;
      for (const form of Object.keys(FORMS)) {
        const failure = load(directory, form, specifier);
        if (failure) failures.push(failure);
      }
    }
  }
  if (checked === 0) failures.push('no package entries found to load');
  if (failures.length) {
    console.error(`${failures.length} loads failed on Node ${process.version}:\n`);
    for (const failure of failures) console.error(`- ${failure}`);
    process.exit(1);
  }
  console.log(`${checked} package entries load by require(), a static import and import() on Node ${process.version}.`);
}

const invoked = process.argv[1] ? realpathSync(process.argv[1]) : null;
if (invoked === realpathSync(fileURLToPath(import.meta.url))) main();
