import { afterAll, expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { buildWorkerStub } from './worker-stub';

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
    const file = join(directory, outfile);
    await buildWorkerStub({ entry: join(packages, entry), outfile: file, start });
    const contents = await readFile(file, 'utf8');
    expect([...contents.matchAll(staticImport)].map(([, specifier]) => specifier)).toEqual([]);
    expect(contents.match(dynamicImport)?.length ?? 0).toBe(1);
    expect([...contents.matchAll(dynamicTarget)].map(([, specifier]) => specifier)).toEqual([start]);
  });
}
