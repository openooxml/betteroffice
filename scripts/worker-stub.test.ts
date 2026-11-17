import { afterAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { access, mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import type { buildWorkerStub as BuildWorkerStub } from './worker-stub';

const buildWorkerStub: typeof BuildWorkerStub | undefined = await import('./worker-stub').then(
  (module) => module.buildWorkerStub,
  () => undefined,
);

const packages = resolve(import.meta.dir, '../packages');
const staticImport = /\b(?:from|import)\s*["'`]([^"'`]+)["'`]/g;
const dynamicImport = /(?<![\w$.])import\s*\(/g;
const dynamicTarget = /(?<![\w$.])import\s*\(\s*["'`]([^"'`]+)["'`]\s*\)/g;
const stubs = [
  { name: 'docx', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/docxWorker.mjs', start: './yrs/residentEngineWorker.mjs' },
  { name: 'docx legacy', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/yrs/docxWorker.mjs', start: './residentEngineWorker.mjs' },
  { name: 'pptx', entry: 'pptx/src/session/worker.ts', outfile: 'pptx/pptxSessionWorker.mjs', start: './pptxWorkerStart.js' },
  { name: 'xlsx', entry: 'xlsx/src/session/worker.ts', outfile: 'xlsx/xlsxSessionWorker.mjs', start: './xlsxWorkerStart.js' },
];
const directory = await mkdtemp(join(tmpdir(), 'worker-stub-'));

afterAll(() => rm(directory, { recursive: true, force: true }));

for (const { name, entry, outfile, start } of stubs) {
  test(`${name} worker stub imports nothing statically and its start file once`, async () => {
    expect(buildWorkerStub).toBeFunction();
    const file = join(directory, outfile);
    await buildWorkerStub!({ entry: join(packages, entry), outfile: file, start });
    const contents = await readFile(file, 'utf8');
    expect([...contents.matchAll(staticImport)].map(([, specifier]) => specifier)).toEqual([]);
    expect(contents.match(dynamicImport)?.length ?? 0).toBe(1);
    expect([...contents.matchAll(dynamicTarget)].map(([, specifier]) => specifier)).toEqual([start]);
  });
}

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
