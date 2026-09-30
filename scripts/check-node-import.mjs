// Imports every export of every published package in plain Node, without a DOM,
// each in its own process. Run after `bun run build:packages`.
import { spawnSync } from 'node:child_process';
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

const PROBE = `
for (const name of ${JSON.stringify(BROWSER_GLOBALS)}) {
  if (!Reflect.deleteProperty(globalThis, name) || name in globalThis) {
    throw new Error('cannot remove the global ' + name);
  }
}
const [specifier, json] = process.argv.slice(1);
await import(specifier, json === 'json' ? { with: { type: 'json' } } : undefined);
`;

const CONDITIONS = new Set(['node', 'import', 'default']);

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

if (typeof Bun !== 'undefined' || !process.versions.node) {
  throw new Error('check-node-import.mjs must run under node');
}

const failures = [];
let checked = 0;
for (const { directory, manifest } of publishedPackageManifests()) {
  if (!manifest.exports) {
    failures.push(`${manifest.name}: no exports map`);
    continue;
  }
  for (const [subpath, value] of subpaths(manifest.exports)) {
    const specifier = manifest.name + subpath.slice(1);
    if (subpath.includes('*')) {
      failures.push(`${specifier}: wildcard exports are not enumerated`);
      continue;
    }
    const target = importTarget(value);
    if (!target || target.endsWith('.css') || /\.d\.[cm]?ts$/.test(target)) continue;
    const result = spawnSync(
      process.execPath,
      [
        '--unhandled-rejections=strict',
        '--input-type=module',
        '--eval',
        PROBE,
        specifier,
        target.endsWith('.json') ? 'json' : 'js',
      ],
      {
        cwd: directory,
        encoding: 'utf8',
        timeout: 60_000,
        env: { ...process.env, NODE_OPTIONS: '' },
      }
    );
    checked += 1;
    if (result.status !== 0) {
      const reason = result.error?.message ?? `exit ${result.status ?? result.signal}`;
      const output = `${result.stderr ?? ''}${result.stdout ?? ''}`.trim().split('\n').slice(-8);
      failures.push(`${specifier}: ${reason}\n    ${output.join('\n    ')}`);
    }
  }
}

if (failures.length) {
  console.error(`${failures.length} package entries failed to import in plain Node:\n`);
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
console.log(`${checked} package entries import in plain Node without a DOM.`);
