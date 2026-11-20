import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { FORMS, errorCode, importOnlyExports } from './check-node-require.mjs';

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
