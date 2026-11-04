import { expect, test } from 'bun:test';
import { exportSpecifiers, importTarget, packageManifests } from './check-webpack-exports.mjs';

test('every export subpath becomes a specifier, wildcards left out', () => {
  expect(
    exportSpecifiers({
      name: '@betteroffice/fixture',
      exports: {
        '.': { types: './dist/index.d.ts', import: './dist/index.js', default: './dist/index.js' },
        './styles.css': './dist/styles.css',
        './*': './dist/*.js',
        './package.json': './package.json',
      },
    })
  ).toEqual(['@betteroffice/fixture', '@betteroffice/fixture/styles.css', '@betteroffice/fixture/package.json']);
});

test('a root-only exports field is the package itself', () => {
  expect(exportSpecifiers({ name: 'a', exports: './index.js' })).toEqual(['a']);
  expect(exportSpecifiers({ name: 'b', exports: { import: './b.js', default: './b.js' } })).toEqual(['b']);
  expect(exportSpecifiers({ name: 'c' })).toEqual([]);
});

test('the import target of an export entry', () => {
  const manifest = {
    name: '@betteroffice/fixture-native',
    exports: {
      '.': { types: './dist/index.d.ts', browser: './dist/browser.js', import: './dist/index.js' },
      './worker': { types: './dist/worker.d.ts', browser: './dist/browser.js', import: './dist/worker.js' },
      './styles.css': './dist/styles.css',
    },
  };
  expect(importTarget(manifest, '@betteroffice/fixture-native')).toBe('./dist/index.js');
  expect(importTarget(manifest, '@betteroffice/fixture-native/worker')).toBe('./dist/worker.js');
  expect(importTarget(manifest, '@betteroffice/fixture-native/styles.css')).toBeUndefined();
  expect(importTarget({ name: 'b', exports: { import: './b.js', default: './b.js' } }, 'b')).toBe('./b.js');
});

test('the checked libraries are every package under packages/ with an exports map', () => {
  const names = packageManifests().map(({ manifest }) => manifest.name);
  for (const format of ['docx', 'xlsx', 'pptx', 'vsdx']) {
    for (const suffix of ['', '-react', '-i18n']) expect(names).toContain(`@betteroffice/${format}${suffix}`);
  }
  expect(names).toContain('@betteroffice/agents');
  expect(names).toContain('@betteroffice/fonts');
  expect(names.every((name) => name.startsWith('@betteroffice/'))).toBe(true);
});
