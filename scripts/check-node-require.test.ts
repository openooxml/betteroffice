import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FORMS, errorCode, importOnlyExports, requireFloorProblems, rootIsModule } from './check-node-require.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const manifest = (directory: string) => JSON.parse(readFileSync(`${root}${directory}/package.json`, 'utf8'));

test('an export a require() cannot reach is one with import and no default for the same file', () => {
  expect(
    importOnlyExports({
      name: '@betteroffice/fixture',
      exports: {
        '.': { types: './dist/index.d.ts', browser: './dist/browser.js', import: './dist/index.js' },
        './worker': { types: './dist/worker.d.ts', import: './dist/worker.js', default: './dist/worker.js' },
        './other': { import: './dist/other.js', default: './dist/another.js' },
        './styles.css': './dist/styles.css',
        './types-only': { types: './dist/types.d.ts' },
      },
    })
  ).toEqual(['@betteroffice/fixture', '@betteroffice/fixture/other']);
  expect(importOnlyExports({ name: 'a', exports: { import: './a.js' } })).toEqual(['a']);
  expect(importOnlyExports({ name: 'b', exports: { import: './b.js', default: './b.js' } })).toEqual([]);
  expect(importOnlyExports({ name: 'c', exports: './c.js' })).toEqual([]);
  expect(importOnlyExports({ name: 'd' })).toEqual([]);
});

test('the three forms are a require, a static import and an import() of the same specifier', () => {
  const sources = Object.values(FORMS).map((form) => form('@betteroffice/fixture/worker'));
  expect(sources.map(([type]) => type)).toEqual(['--input-type=commonjs', '--input-type=module', '--input-type=commonjs']);
  expect(sources[0]![2]).toStartWith('require("@betteroffice/fixture/worker");');
  expect(sources[1]![2]).toStartWith('import "@betteroffice/fixture/worker";');
  expect(sources[2]![2]).toStartWith('import("@betteroffice/fixture/worker").then(');
});

test('the error code is read from what node prints', () => {
  expect(errorCode("  code: 'ERR_PACKAGE_PATH_NOT_EXPORTED'\n}")).toBe('ERR_PACKAGE_PATH_NOT_EXPORTED');
  expect(errorCode('Error [ERR_REQUIRE_ESM]: require() of ES Module')).toBe('ERR_REQUIRE_ESM');
  expect(errorCode('SyntaxError: Unexpected token')).toBeUndefined();
});

test('every export of the native bindings is reachable by require(), as in the wasm packages', () => {
  for (const format of ['docx', 'xlsx', 'pptx']) {
    expect(importOnlyExports(manifest(`packages/${format}`))).toEqual([]);
    expect(importOnlyExports(manifest(`bindings/${format}-native`))).toEqual([]);
  }
});

test('CI loads the wasm packages and the native bindings in all three forms', () => {
  const workflow = readFileSync(`${root}.github/workflows/ci.yml`, 'utf8');
  expect(workflow).toContain('run: node scripts/check-node-require.mjs packages/docx packages/xlsx packages/pptx\n');
  expect(workflow).toContain('node scripts/check-node-require.mjs $(node scripts/native-bindings.mjs --paths)\n');
});

test('an engines range that admits a Node without require() of ES modules is a problem', () => {
  const problems = (node?: string, rootModule = true) =>
    requireFloorProblems({ name: 'p', ...(node === undefined ? {} : { engines: { node } }) }, rootModule);
  expect(problems('^20.16.0 || >=22.3.0')).toEqual([
    'p: engines.node "^20.16.0 || >=22.3.0" admits Node 20.18.3 outside ^20.19.0 || >=22.12.0, where require() of an ES module fails with ERR_REQUIRE_ESM',
  ]);
  expect(problems('>=18')[0]).toContain('Node 18.20.8');
  expect(problems('^19.0.0 || >=22.12.0')[0]).toContain('Node 19.9.0');
  expect(problems('>=20.17.0 <20.18.0 || >=22.12.0')[0]).toContain('admits versions outside');
  expect(problems('not a range')).toEqual(['p: engines.node "not a range" is not a valid range']);
  expect(problems('^20.19.0 || >=22.12.0')).toEqual([]);
  expect(problems('>=22.12')).toEqual([]);
  expect(problems()).toEqual([]);
  expect(problems('^20.16.0 || >=22.3.0', false)).toEqual([]);
});

test('the root a require() loads is an ES module by extension or by the nearest package.json', () => {
  const directory = mkdtempSync(join(tmpdir(), 'require-floor-'));
  try {
    mkdirSync(join(directory, 'dist'));
    expect(rootIsModule(directory, { exports: { '.': { import: './dist/core.mjs', default: './dist/core.mjs' } } })).toBe(true);
    expect(rootIsModule(directory, { exports: { require: './dist/core.cjs', import: './dist/core.mjs' } })).toBe(false);
    expect(rootIsModule(directory, { type: 'commonjs', exports: { default: './dist/index.js' } })).toBe(false);
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ type: 'commonjs' }));
    writeFileSync(join(directory, 'dist', 'package.json'), JSON.stringify({ type: 'module' }));
    expect(rootIsModule(directory, { type: 'commonjs', exports: { default: './dist/index.js' } })).toBe(true);
    expect(rootIsModule(directory, { main: 'index.js' })).toBe(false);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test('the native roots declare the Node versions their require() needs', () => {
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const native = manifest(`bindings/${format}-native`);
    expect(native.engines.node).toBe('^20.19.0 || >=22.12.0');
    expect(requireFloorProblems(native)).toEqual([]);
  }
});
