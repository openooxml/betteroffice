import { expect, test } from 'bun:test';

import { validateNativeWorker, validatePublishedWorker } from './worker-entry-check.mjs';

const fonts = "import '@betteroffice/fonts/worker';";

test('a bundled worker may statically import the worker font entry', async () => {
  expect(await validatePublishedWorker("import { createFontProvider } from '@betteroffice/fonts/worker'; globalThis.fonts = createFontProvider;"))
    .toEqual({ problems: [], target: null });
});

test('literal dynamic imports fail even when nested or behind a comment', async () => {
  for (const source of ["import('./x.js');", "globalThis.load = () => import /*c*/ ('./x.js');"]) {
    expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  }
});

test('nonliteral node-only imports remain permitted', async () => {
  expect(await validatePublishedWorker(`${fonts} const name = globalThis.moduleName; globalThis.load = () => import(name);`))
    .toEqual({ problems: [], target: null });
});

test('a browser worker requires exactly one static font import', async () => {
  for (const source of ['', 'globalThis.ready = true;', `${fonts} import { createFontProvider } from '@betteroffice/fonts/worker'; globalThis.fonts = createFontProvider;`]) {
    expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  }
});

test('other static imports and requires fail', async () => {
  for (const source of ["import /*c*/ './x.js';", "import '@betteroffice/fonts';", "require('./x.js');"]) {
    expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  }
});

test('only the docx sibling entry may import its canonical bundle', async () => {
  const source = "import '../docxWorker.mjs';";
  expect(await validatePublishedWorker(source, '../docxWorker.mjs')).toEqual({ problems: [], target: '../docxWorker.mjs' });
  expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  expect((await validatePublishedWorker('', '../docxWorker.mjs')).problems.length).toBeGreaterThan(0);
  expect((await validatePublishedWorker(`${source}globalThis.extra = true;`, '../docxWorker.mjs')).problems.length).toBeGreaterThan(0);
  expect((await validatePublishedWorker(`${source}${fonts}`, '../docxWorker.mjs')).problems.length).toBeGreaterThan(0);
});

test('native workers allow bare static imports and bare dynamic font imports', async () => {
  expect(await validateNativeWorker("import { createRequire } from 'node:module'; globalThis.require = createRequire; globalThis.fonts = () => import('@betteroffice/fonts'); globalThis.progressive = () => import('@betteroffice/fonts/progressive');"))
    .toEqual({ problems: [], target: null });
});

test('the native DOCX sibling contains only the canonical worker import', async () => {
  const sibling = '../docxWorker.mjs';
  const source = `import '${sibling}';`;
  expect(await validateNativeWorker(source, sibling)).toEqual({ problems: [], target: sibling });
  expect((await validateNativeWorker(source)).problems.length).toBeGreaterThan(0);
  expect((await validateNativeWorker(`${source}${fonts}`, sibling)).problems.length).toBeGreaterThan(0);
});

test('native workers reject static and dynamic file imports and the browser font entry', async () => {
  for (const source of [
    "import './chunk.js';", "import '../chunk.js';", "import('./chunk.js');", "import /*c*/ ('../chunk.js');",
    "import '/chunk.js';", "import { setOptionalFontPacks } from '@betteroffice/fonts/worker'; globalThis.install = setOptionalFontPacks;",
  ]) {
    expect((await validateNativeWorker(source)).problems.length).toBeGreaterThan(0);
  }
});
