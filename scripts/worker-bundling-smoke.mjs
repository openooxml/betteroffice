import assert from 'node:assert/strict';
import { cp, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build as esbuild } from 'esbuild';
import { build as vite } from 'vite';

const root = fileURLToPath(new URL('../', import.meta.url));
const cache = join(root, 'node_modules/.cache');
await mkdir(cache, { recursive: true });
const directory = await mkdtemp(join(cache, 'worker-bundling-'));
const formats = { xlsx: 'xlsxSessionWorker', pptx: 'pptxSessionWorker', docx: 'docxWorker' };
const app = `
import { openWorkbookSession } from '@betteroffice/xlsx/worker';
import { openPresentationSession } from '@betteroffice/pptx/worker';
import { DocxSourceWorkerClient } from '@betteroffice/docx/worker';
globalThis.workerSessionSmoke = { openWorkbookSession, openPresentationSession, DocxSourceWorkerClient };
`;

async function outputFiles(directory) {
  const files = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await outputFiles(path));
    else files.push(path);
  }
  return files;
}

async function checkOutput(directory) {
  const files = await outputFiles(directory);
  let javascript = '';
  for (const worker of Object.values(formats)) {
    assert.ok(files.some(path => /\.(?:m?js)$/.test(path) && path.includes(worker)), `Missing emitted ${worker}`);
  }
  for (const path of files.filter(path => /\.(?:m?js|html)$/.test(path))) {
    const contents = await readFile(path, 'utf8');
    if (/\.(?:m?js)$/.test(path)) javascript += contents;
    assert.doesNotMatch(contents, /worker_threads|nodeSessionWorker|sessionWorker\.node/,
      `Node worker code in browser output: ${path}`);
  }
  assert.match(javascript, /new Worker\s*\(/, 'Browser worker constructors must survive bundling');
}

try {
  await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  const workers = {};
  for (const [format, worker] of Object.entries(formats)) {
    const source = join(root, 'packages', format);
    const target = join(directory, 'node_modules/@betteroffice', format);
    await mkdir(target, { recursive: true });
    await cp(join(source, 'package.json'), join(target, 'package.json'));
    await cp(join(source, 'dist'), join(target, 'dist'), { recursive: true });
    const manifest = JSON.parse(await readFile(join(target, 'package.json'), 'utf8'));
    assert.match(manifest.imports['#session-worker'].default, /sessionWorker\.browser\.(?:js|mjs)$/);
    assert.match(manifest.imports['#session-worker'].node, /sessionWorker\.node\.(?:js|mjs)$/);
    workers[worker] = join(target, 'dist', `${worker}.mjs`);
  }
  const entry = join(directory, 'app.mjs');
  await writeFile(entry, app);
  const esbuildDirectory = join(directory, 'esbuild');
  const bundled = await esbuild({
    absWorkingDir: root,
    entryPoints: { app: entry, ...workers },
    outdir: esbuildDirectory, outExtension: { '.js': '.mjs' },
    platform: 'browser', conditions: ['browser', 'default'],
    bundle: true, format: 'esm', splitting: true, metafile: true, minify: false,
  });
  for (const format of Object.keys(formats)) {
    const browser = join(directory, 'node_modules/@betteroffice', format, 'dist/sessionWorker.browser');
    assert.ok(Object.keys(bundled.metafile.inputs).some(path => resolve(root, path).startsWith(browser)),
      `esbuild did not select the ${format} browser variant`);
  }
  assert.ok(Object.keys(bundled.metafile.inputs).every(path => !/sessionWorker\.node|nodeSessionWorker/.test(path)));
  await checkOutput(esbuildDirectory);

  await writeFile(join(directory, 'index.html'), '<script type="module" src="./app.mjs"></script>');
  const viteDirectory = join(directory, 'vite');
  await vite({
    configFile: false, root: directory, logLevel: 'warn',
    resolve: { conditions: ['browser', 'import', 'default'] },
    worker: { format: 'es' },
    build: { outDir: viteDirectory, emptyOutDir: true, assetsInlineLimit: 0, minify: false },
  });
  await checkOutput(viteDirectory);
  console.log('esbuild and Vite browser worker bundles exclude Node worker code.');
} finally {
  await rm(directory, { recursive: true, force: true });
}
