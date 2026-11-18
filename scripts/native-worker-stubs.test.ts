import { afterAll, expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { validatePublishedWorker } from './worker-entry-check.mjs';
import type { buildWorkerStub as BuildWorkerStub, WorkerStub } from './worker-stub';

type Format = 'xlsx' | 'pptx' | 'docx';
type WorkerStubs = (format: Format, outDir: string) => WorkerStub[];

const stubModule = await import('./worker-stub').catch(() => undefined);
const buildWorkerStub: typeof BuildWorkerStub | undefined = stubModule?.buildWorkerStub;
const workerStubs: WorkerStubs | undefined = (stubModule as { workerStubs?: WorkerStubs } | undefined)?.workerStubs;

const root = resolve(import.meta.dir, '..');
const formats: Format[] = ['xlsx', 'pptx', 'docx'];
const startEntries: Record<Format, string[]> = {
  xlsx: ['xlsxWorkerStart'],
  pptx: ['pptxWorkerStart'],
  docx: ['yrs/residentEngineWorker', 'yrs/residentEngineWorker'],
};
const directory = await mkdtemp(join(tmpdir(), 'native-worker-stubs-'));

afterAll(() => rm(directory, { recursive: true, force: true }));

async function entryKeys(config: string): Promise<string[]> {
  const module = await import(join(root, config));
  const exported = module.default;
  const resolved = typeof exported === 'function' ? await exported({}) : exported;
  const keys = new Set<string>();
  for (const options of [resolved].flat()) {
    const entry = options.entry;
    for (const key of Array.isArray(entry) ? entry : Object.keys(entry ?? {})) keys.add(key);
  }
  return [...keys].sort();
}

for (const format of formats) {
  test(`${format}-native builds the same entries as ${format}`, async () => {
    const wasm = await entryKeys(`packages/${format}/tsup.config.ts`);
    const native = await entryKeys(`bindings/${format}-native/js/tsup.config.ts`);
    expect(native).toEqual(wasm);
  });

  test(`${format}-native worker stubs import nothing statically and start an entry the binding builds`, async () => {
    expect(buildWorkerStub).toBeFunction();
    expect(workerStubs).toBeFunction();
    const out = join(directory, `${format}-native`);
    const stubs = workerStubs!(format, out);
    const native = await entryKeys(`bindings/${format}-native/js/tsup.config.ts`);
    const started: string[] = [];
    const failures: string[] = [];
    for (const stub of stubs) {
      await buildWorkerStub!(stub);
      const { problems, target } = await validatePublishedWorker(await readFile(stub.outfile, 'utf8'));
      failures.push(...problems.map((problem: string) => `${relative(out, stub.outfile)} ${problem}`));
      if (target === null) continue;
      const entry = relative(out, resolve(dirname(stub.outfile), target)).replace(/\.m?js$/, '');
      started.push(entry);
      if (!native.includes(entry)) failures.push(`${relative(out, stub.outfile)} starts ${target}, which ${format}-native does not build`);
    }
    expect({ failures, started }).toEqual({ failures: [], started: startEntries[format] });
  });
}
