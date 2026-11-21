import { afterAll, expect, test } from 'bun:test';
import { access, mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { build } from 'esbuild';
import { nativeEnginePlugin } from './native-engine';
import { validateNativeWorker } from './worker-entry-check.mjs';
import type { WorkerStub } from './worker-stub';

type Format = 'xlsx' | 'pptx' | 'docx';
type WorkerStubs = (format: Format, outDir: string) => WorkerStub[];

const stubModule = await import('./worker-stub').catch(() => undefined);
const buildWorkerStubs = stubModule?.buildWorkerStubs;
const workerStubs: WorkerStubs | undefined = (stubModule as { workerStubs?: WorkerStubs } | undefined)?.workerStubs;

const root = resolve(import.meta.dir, '..');
const formats: Format[] = ['xlsx', 'pptx', 'docx'];
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

  test(`${format}-native module workers bundle native engines without wasm assets or chunk imports`, async () => {
    expect(buildWorkerStubs).toBeFunction();
    expect(workerStubs).toBeFunction();
    const out = join(directory, `${format}-native`);
    const stubs = workerStubs!(format, out);
    const config = (await import(join(root, `bindings/${format}-native/js/tsup.config.ts`))).default;
    const external: string[] = [...[config].flat()[0].external, '@betteroffice/fonts', '@betteroffice/fonts/*'];
    await buildWorkerStubs!(format, out, { plugins: [nativeEnginePlugin(format)], external });
    const failures: string[] = [];
    for (const stub of stubs) {
      const contents = await readFile(stub.outfile, 'utf8');
      const { problems } = await validateNativeWorker(contents);
      failures.push(...problems.map((problem: string) => `${relative(out, stub.outfile)} ${problem}`));
      expect(contents).not.toMatch(/new URL\([^)]*\.wasm|WebAssembly\.instantiate|generated\//);
      expect(contents).not.toContain('font-pack-request');
      const parsed = await build({
        stdin: { contents },
        bundle: true,
        format: 'esm',
        external: ['*'],
        write: false,
        metafile: true,
      });
      const imports = Object.values(parsed.metafile!.outputs).flatMap((output) => output.imports);
      expect(imports.some(({ path }) => path === 'node:module')).toBe(true);
      expect(imports.filter(({ path }) => path === '@betteroffice/fonts').map(({ kind }) => kind))
        .toEqual(format === 'xlsx' ? [] : expect.arrayContaining(['dynamic-import']));
      for (const { path } of imports) {
        if (external.some((pattern) =>
          pattern.endsWith('*') ? path.startsWith(pattern.slice(0, -1)) : path === pattern)) continue;
        failures.push(`${relative(out, stub.outfile)} imports ${path}`);
      }
    }
    expect(failures).toEqual([]);
    expect(await access(join(out, 'generated')).then(() => true, () => false)).toBe(false);
  });
}
