import { expect, test } from 'bun:test';
import { importPlan } from './check-node-import.mjs';

test('every export is imported or refused, never skipped', () => {
  const plan = importPlan({
    name: '@betteroffice/fixture',
    exports: {
      '.': { types: './dist/index.d.ts', import: './dist/index.js' },
      './data': './dist/data.json',
      './types-only': { types: './dist/types.d.ts' },
      './browser-only': { browser: './dist/browser.js' },
      './fallbacks': { import: ['./dist/a.js', './dist/b.js'] },
      './styles.css': './dist/styles.css',
      './*': './dist/*.js',
      './package.json': './package.json',
    },
  });
  expect(plan.entries).toEqual([
    { specifier: '@betteroffice/fixture', json: false },
    { specifier: '@betteroffice/fixture/data', json: true },
    { specifier: '@betteroffice/fixture/package.json', json: true },
  ]);
  expect(plan.failures.map((failure) => failure.split(':')[0])).toEqual([
    '@betteroffice/fixture/types-only',
    '@betteroffice/fixture/browser-only',
    '@betteroffice/fixture/fallbacks',
    '@betteroffice/fixture/styles.css',
    '@betteroffice/fixture/*',
  ]);
});

test('a root export array is refused', () => {
  expect(importPlan({ name: '@betteroffice/fixture', exports: [] }).failures).toHaveLength(1);
});

test('the editor stylesheet is the one export left out', () => {
  const plan = importPlan({
    name: '@betteroffice/docx-react',
    exports: { '.': { import: './dist/index.mjs' }, './styles.css': './dist/styles.css' },
  });
  expect(plan).toEqual({
    entries: [{ specifier: '@betteroffice/docx-react', json: false }],
    failures: [],
  });
});
