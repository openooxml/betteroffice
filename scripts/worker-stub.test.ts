import { afterAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import type { buildWorkerStub as BuildWorkerStub } from './worker-stub';
import { validateNativeWorker, validatePublishedWorker } from './worker-entry-check.mjs';

const buildWorkerStub: typeof BuildWorkerStub | undefined = await import('./worker-stub').then(
  (module) => module.buildWorkerStub,
  () => undefined,
);

const packages = resolve(import.meta.dir, '../packages');
const stubs = [
  { name: 'docx', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/docxWorker.mjs' },
  { name: 'docx legacy', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/yrs/docxWorker.mjs', sibling: '../docxWorker.mjs' },
  { name: 'pptx', entry: 'pptx/src/session/worker.ts', outfile: 'pptx/pptxSessionWorker.mjs' },
  { name: 'xlsx', entry: 'xlsx/src/session/worker.ts', outfile: 'xlsx/xlsxSessionWorker.mjs' },
];
const directory = await mkdtemp(join(tmpdir(), 'worker-stub-'));

afterAll(() => rm(directory, { recursive: true, force: true }));

for (const { name, entry, outfile, ...options } of stubs) {
  test(`${name} worker bundles its start module without literal dynamic imports`, async () => {
    expect(buildWorkerStub).toBeFunction();
    const file = join(directory, outfile);
    await buildWorkerStub!({ entry: join(packages, entry), outfile: file, ...options }, { browserFonts: true });
    const contents = await readFile(file, 'utf8');
    expect((await validatePublishedWorker(contents, 'sibling' in options ? options.sibling : null)).problems).toEqual([]);
    if ('sibling' in options) {
      expect(contents.trim()).toBe('import "../docxWorker.mjs";');
    } else {
      const urls = [...contents.matchAll(/new URL\(["']([^"']+)["'],\s*import\.meta\.url\)/g)].map(([, asset]) => asset!);
      expect(urls.length).toBeGreaterThan(0);
      for (const asset of urls) {
        expect(asset).toStartWith('./generated/');
        expect(new URL(asset, `https://workers.example/${outfile}`).pathname).toStartWith(`/${outfile.split('/')[0]}/generated/`);
      }
      expect(contents).not.toMatch(/data:application\/wasm/);
      expect(contents).toContain('font-pack-request');
      expect(contents).toContain('stopImmediatePropagation');
    }
  });
}

test('only browser worker bundles redirect dynamic font imports and install the pack channel', async () => {
  const entry = join(directory, 'font-imports.ts');
  await writeFile(entry, `globalThis.fonts = () => import('@betteroffice/fonts');
globalThis.progressive = () => import('@betteroffice/fonts/progressive');`);
  const browser = join(directory, 'font-imports.browser.mjs');
  await buildWorkerStub!({ entry, outfile: browser }, { browserFonts: true });
  const source = await readFile(browser, 'utf8');
  expect((await validatePublishedWorker(source)).problems).toEqual([]);
  expect(source).toContain('font-pack-request');
  const native = join(directory, 'font-imports.native.mjs');
  await buildWorkerStub!({ entry, outfile: native }, { external: ['@betteroffice/fonts', '@betteroffice/fonts/*'] });
  const plain = await readFile(native, 'utf8');
  expect((await validateNativeWorker(plain)).problems).toEqual([]);
  expect(plain).toContain('import("@betteroffice/fonts")');
  expect(plain).toContain('import("@betteroffice/fonts/progressive")');
  expect(plain).not.toContain('font-pack-request');
});

test('a docx build into another directory writes both worker stubs there and none into the package dist', async () => {
  const docx = join(packages, 'docx');
  const out = join(directory, 'docx-out');
  const published = ['dist/docxWorker.mjs', 'dist/yrs/docxWorker.mjs'].map((file) => join(docx, file));
  const snapshot = () => Promise.all(published.map((file) => readFile(file).then(
    (bytes) => createHash('sha256').update(bytes).digest('hex'),
    () => null,
  )));
  const before = await snapshot();
  const build = spawnSync(process.execPath, ['x', 'tsup', '--out-dir', out], {
    cwd: docx,
    encoding: 'utf8',
    env: { ...process.env, NODE_OPTIONS: '--max-old-space-size=8192' },
  });
  if (build.status !== 0) throw new Error(build.stdout + build.stderr);
  const emitted = ['docxWorker.mjs', 'yrs/docxWorker.mjs', 'yrs/residentEngineWorker.mjs'];
  const present = await Promise.all(emitted.map((file) => access(join(out, file)).then(() => file, () => `${file} missing`)));
  expect({ present, packageDist: await snapshot() }).toEqual({ present: emitted, packageDist: before });
}, 300_000);
