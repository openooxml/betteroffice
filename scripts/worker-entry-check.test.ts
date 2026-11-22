import { expect, test } from 'bun:test';

import { validateNativeWorker, validatePublishedWorker } from './worker-entry-check.mjs';

const fonts = "import '@betteroffice/fonts/worker';";

test('a bundled worker has no imports and may locate its engine wasm', async () => {
  expect(await validatePublishedWorker("self.engine = new URL('./generated/engine_bg.wasm', import.meta.url); self.reason = 'No import() needed';"))
    .toEqual({ problems: [], target: null });
});

test('literal and computed imports fail even when nested or behind a comment', async () => {
  for (const source of ["import('./x.js');", "globalThis.load = () => import /*c*/ ('./x.js');", 'self.load = (name) => import(name);']) {
    expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  }
});

test('other static imports and requires fail', async () => {
  for (const source of [fonts, "import /*c*/ './x.js';", "import '@betteroffice/fonts';", "require('./x.js');"]) {
    expect((await validatePublishedWorker(source)).problems.length).toBeGreaterThan(0);
  }
});

test('font assets and the decoder cannot be resolved relative to a worker', async () => {
  for (const path of ['../assets/face.ttf.zst', './decoder/fontpack.wasm', './generated/fontpack.wasm']) {
    const source = `self.asset = new URL(${JSON.stringify(path)}, import.meta.url);`;
    expect((await validatePublishedWorker(source)).problems).toContain(`resolves ${path} relative to the worker module`);
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
