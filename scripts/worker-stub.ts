import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build, type BuildOptions, type Plugin } from 'esbuild';

export type WorkerStub = { entry: string; outfile: string; sibling?: string };
type WorkerBuildOptions = Pick<BuildOptions, 'plugins' | 'external'> & { browserFonts?: boolean };

const packages = fileURLToPath(new URL('../packages/', import.meta.url));
const fontPackInstaller = fileURLToPath(new URL('../shared/office-session/workerFontPacks.ts', import.meta.url));

function workerFontsPlugin(): Plugin {
  const namespace = 'worker-fonts';
  return {
    name: namespace,
    setup(build) {
      build.onResolve({ filter: /^@betteroffice\/fonts(?:\/progressive|\/worker)?$/ }, ({ path, kind, importer, namespace: from }) => {
        if (path === '@betteroffice/fonts/worker') {
          if (from === namespace || importer !== fontPackInstaller) return { path, external: true };
        } else if (kind !== 'dynamic-import') return;
        return { path: 'fonts', namespace };
      });
      build.onLoad({ filter: /.*/, namespace }, () => ({
        contents: "export * from '@betteroffice/fonts/worker';",
        resolveDir: dirname(fontPackInstaller),
        loader: 'js',
      }));
    },
  };
}

const WORKER_STUBS = {
  xlsx: [{ entry: 'xlsx/src/session/worker.ts', outfile: 'xlsxSessionWorker.mjs' }],
  pptx: [{ entry: 'pptx/src/session/worker.ts', outfile: 'pptxSessionWorker.mjs' }],
  docx: [
    { entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docxWorker.mjs' },
    { entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'yrs/docxWorker.mjs', sibling: '../docxWorker.mjs' },
  ],
} as const;

/** Browser worker entries shared by the format packages and native bindings. */
export function workerStubs(format: keyof typeof WORKER_STUBS, outDir: string): WorkerStub[] {
  return WORKER_STUBS[format].map((stub) => ({
    ...stub,
    entry: join(packages, stub.entry),
    outfile: join(outDir, stub.outfile),
  }));
}

export async function buildWorkerStubs(format: keyof typeof WORKER_STUBS, outDir: string, options: WorkerBuildOptions = {}): Promise<void> {
  for (const stub of workerStubs(format, outDir)) await buildWorkerStub(stub, options);
}

export async function buildWorkerStub({ entry, outfile, sibling }: WorkerStub, options: WorkerBuildOptions = {}): Promise<void> {
  if (sibling && options.browserFonts) {
    await mkdir(dirname(outfile), { recursive: true });
    await writeFile(outfile, `import ${JSON.stringify(sibling)};\n`);
    return;
  }
  const { browserFonts = false, plugins = [], external = [] } = options;
  const originalEntry = resolve(entry);
  await build({
    ...(browserFonts ? {
      stdin: {
        contents: `import ${JSON.stringify(fontPackInstaller)}; import ${JSON.stringify(originalEntry)};`,
        resolveDir: dirname(originalEntry),
        sourcefile: 'browser-worker-entry.ts',
        loader: 'ts' as const,
      },
    } : { entryPoints: [originalEntry] }),
    outfile,
    bundle: true,
    splitting: false,
    format: 'esm',
    platform: 'browser',
    minify: true,
    logLevel: 'warning',
    plugins: browserFonts ? [workerFontsPlugin(), ...plugins] : plugins,
    external,
  });
}
