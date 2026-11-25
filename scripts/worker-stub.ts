import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build, type BuildOptions, type Plugin } from 'esbuild';

export type WorkerStub = { entry: string; outfile: string; sibling?: string };
type WorkerBuildOptions = Pick<BuildOptions, 'plugins' | 'external'> & { browserFonts?: boolean };

const packages = fileURLToPath(new URL('../packages/', import.meta.url));
const fontPackInstaller = fileURLToPath(new URL('../shared/office-session/workerFontPacks.ts', import.meta.url));
const fontModules = fileURLToPath(new URL('../shared/office-session/fontModules', import.meta.url));
const fontSources = join(packages, 'fonts/src');

function workerFontsPlugin(): Plugin {
  return {
    name: 'worker-fonts',
    setup(build) {
      build.onResolve({ filter: /(?:^|\/)fontModules$/ }, ({ path, resolveDir }) => {
        if (resolve(resolveDir, path) === fontModules) return { path: join(fontSources, 'workerFontModules.ts') };
      });
      build.onResolve({ filter: /^@betteroffice\/fonts$/ }, () => ({ path: join(fontSources, 'worker.ts') }));
      build.onResolve({ filter: /^@betteroffice\/fonts\/remote$/ }, () => ({ path: join(fontSources, 'remote.ts') }));
      build.onResolve({ filter: /^\.\/(?:decodeAssets|assetHref)$/ }, ({ importer }) => {
        if (dirname(importer) === fontSources) return { path: join(fontSources, 'workerAssets.ts') };
      });
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
  if (sibling) {
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
