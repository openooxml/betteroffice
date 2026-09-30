// Imports every export of every published package in plain Node, without a DOM,
// each in its own process. Run after `bun run build:packages`.
import { spawnSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';
import { publishedPackageManifests } from './published-packages.mjs';

const BROWSER_GLOBALS = [
  'window',
  'document',
  'self',
  'navigator',
  'Worker',
  'OffscreenCanvas',
  'HTMLCanvasElement',
  'Image',
  'ImageBitmap',
  'createImageBitmap',
  'FontFace',
  'requestAnimationFrame',
  'localStorage',
  'sessionStorage',
  'matchMedia',
  'getComputedStyle',
  'customElements',
  'CSS',
];

/** Exports that are not JavaScript modules, by specifier. */
const NOT_MODULES = new Set(['@betteroffice/docx-react/styles.css']);

const PROBE = `
for (const name of ${JSON.stringify(BROWSER_GLOBALS)}) {
  if (!Reflect.deleteProperty(globalThis, name) || name in globalThis) {
    throw new Error('cannot remove the global ' + name);
  }
}
const [specifier, json] = process.argv.slice(1);
await import(specifier, json === 'json' ? { with: { type: 'json' } } : undefined);
`;

const CONDITIONS = new Set(['node', 'import', 'module-sync', 'default']);

function importTarget(value) {
  if (typeof value === 'string') return value;
  if (!value || Array.isArray(value)) return null;
  for (const [condition, nested] of Object.entries(value)) {
    if (!CONDITIONS.has(condition)) continue;
    const target = importTarget(nested);
    if (target) return target;
  }
  return null;
}

function subpaths(exportsField) {
  const conditional =
    typeof exportsField === 'string' ||
    Object.keys(exportsField).some((key) => !key.startsWith('.'));
  return conditional ? [['.', exportsField]] : Object.entries(exportsField);
}

/**
 * The entries of `manifest` to import, as `{ specifier, json }`, and the entries the check cannot
 * import, as messages. Only the exports in {@link NOT_MODULES} are left out.
 */
export function importPlan(manifest) {
  const entries = [];
  const failures = [];
  if (!manifest.exports) return { entries, failures: [`${manifest.name}: no exports map`] };
  for (const [subpath, value] of subpaths(manifest.exports)) {
    const specifier = manifest.name + subpath.slice(1);
    if (NOT_MODULES.has(specifier)) continue;
    const target = subpath.includes('*') ? null : importTarget(value);
    if (!target || target.endsWith('.css') || /\.d\.[cm]?ts$/.test(target)) {
      failures.push(`${specifier}: no JavaScript target under the node, import or default condition`);
      continue;
    }
    entries.push({ specifier, json: target.endsWith('.json') });
  }
  return { entries, failures };
}

function importInNode(directory, { specifier, json }) {
  const result = spawnSync(
    process.execPath,
    [
      '--unhandled-rejections=strict',
      '--input-type=module',
      '--eval',
      PROBE,
      specifier,
      json ? 'json' : 'js',
    ],
    { cwd: directory, encoding: 'utf8', timeout: 60_000, env: { ...process.env, NODE_OPTIONS: '' } }
  );
  if (result.status === 0) return null;
  const reason = result.error?.message ?? `exit ${result.status ?? result.signal}`;
  const output = `${result.stderr ?? ''}${result.stdout ?? ''}`.trim().split('\n').slice(-8);
  return `${specifier}: ${reason}\n    ${output.join('\n    ')}`;
}

function main() {
  if (typeof Bun !== 'undefined' || !process.versions.node) {
    throw new Error('check-node-import.mjs must run under node');
  }
  const failures = [];
  let checked = 0;
  for (const { directory, manifest } of publishedPackageManifests()) {
    const plan = importPlan(manifest);
    failures.push(...plan.failures);
    for (const entry of plan.entries) {
      checked += 1;
      const failure = importInNode(directory, entry);
      if (failure) failures.push(failure);
    }
  }
  if (failures.length) {
    console.error(`${failures.length} package entries failed to import in plain Node:\n`);
    for (const failure of failures) console.error(`- ${failure}`);
    process.exit(1);
  }
  console.log(`${checked} package entries import in plain Node without a DOM.`);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) main();
