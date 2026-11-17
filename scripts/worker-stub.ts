import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

export type WorkerStub = { entry: string; outfile: string; start: string };

const packages = fileURLToPath(new URL('../packages/', import.meta.url));

const WORKER_STUBS = {
  xlsx: [{ entry: 'xlsx/src/session/worker.ts', outfile: 'xlsxSessionWorker.mjs', start: './xlsxWorkerStart.js' }],
  pptx: [{ entry: 'pptx/src/session/worker.ts', outfile: 'pptxSessionWorker.mjs', start: './pptxWorkerStart.js' }],
  docx: [
    { entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docxWorker.mjs', start: './yrs/residentEngineWorker.mjs' },
    { entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'yrs/docxWorker.mjs', start: './residentEngineWorker.mjs' },
  ],
} as const;

/** The import-free worker stubs the `format` package and its native binding both write into their `outDir`. */
export function workerStubs(format: keyof typeof WORKER_STUBS, outDir: string): WorkerStub[] {
  return WORKER_STUBS[format].map(({ entry, outfile, start }) => ({
    entry: join(packages, entry),
    outfile: join(outDir, outfile),
    start,
  }));
}

export async function buildWorkerStubs(format: keyof typeof WORKER_STUBS, outDir: string): Promise<void> {
  for (const stub of workerStubs(format, outDir)) await buildWorkerStub(stub);
}

export async function buildWorkerStub({ entry, outfile, start }: WorkerStub): Promise<void> {
  const importer = resolve(entry);
  await build({
    entryPoints: [entry],
    outfile,
    bundle: true,
    splitting: false,
    format: 'esm',
    platform: 'browser',
    minify: true,
    logLevel: 'warning',
    plugins: [{
      name: 'worker-start',
      setup(stub) {
        stub.onResolve({ filter: /.*/ }, (args) =>
          args.kind === 'dynamic-import' && resolve(args.importer) === importer ? { path: start, external: true } : undefined);
      },
    }],
  });
}
