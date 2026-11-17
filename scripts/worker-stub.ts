import { resolve } from 'node:path';
import { build } from 'esbuild';

export type WorkerStub = { entry: string; outfile: string; start: string };

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
