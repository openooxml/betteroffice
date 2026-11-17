import assert from 'node:assert/strict';
import { access, cp, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build as esbuild } from 'esbuild';
import { build as vite } from 'vite';
import { nextWebpack } from './check-webpack-exports.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const cache = join(root, 'node_modules/.cache');
await mkdir(cache, { recursive: true });
const directory = await mkdtemp(join(cache, 'worker-bundling-'));
const formats = { xlsx: 'xlsxSessionWorker', pptx: 'pptxSessionWorker', docx: 'docxWorker' };
const publishedWorkers = [
  ['xlsx', 'dist/xlsxSessionWorker.mjs'],
  ['pptx', 'dist/pptxSessionWorker.mjs'],
  ['docx', 'dist/docxWorker.mjs'],
  ['docx', 'dist/yrs/docxWorker.mjs'],
];
const staticImport = /\b(?:from|import)\s*["'`]([^"'`]+)["'`]/g;
const dynamicImport = /(?<![\w$.])import\s*\(/g;
const dynamicTarget = /(?<![\w$.])import\s*\(\s*["'`]([^"'`]+)["'`]\s*\)/g;
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

async function checkOutput(directory, { namedWorkers = true } = {}) {
  const files = await outputFiles(directory);
  let javascript = '';
  for (const worker of namedWorkers ? Object.values(formats) : []) {
    assert.ok(files.some(path => /\.(?:m?js)$/.test(path) && path.includes(worker)), `Missing emitted ${worker}`);
  }
  for (const path of files.filter(path => /\.(?:m?js|html)$/.test(path))) {
    const contents = await readFile(path, 'utf8');
    if (/\.(?:m?js)$/.test(path)) javascript += contents;
    assert.doesNotMatch(contents, /worker_threads|nodeSessionWorker|sessionWorker\.node|SessionWorker\.node|docxWorker\.node/,
      `Node worker code in browser output: ${path}`);
    assert.doesNotMatch(contents, /["'`]#session-/, `Unresolved private import in browser output: ${path}`);
  }
  assert.match(javascript, /new Worker\s*\(/, 'Browser worker constructors must survive bundling');
}

async function checkWorkerEntriesLoadOnce(directory) {
  const scripts = (await outputFiles(directory)).filter(path => /\.(?:m?js)$/.test(path));
  const imports = new Map();
  for (const path of scripts) {
    const contents = await readFile(path, 'utf8');
    imports.set(path, [...contents.matchAll(staticImport)].map(([, specifier]) => basename(specifier)));
  }
  const offenders = [];
  for (const worker of Object.values(formats)) {
    const entries = scripts.filter(path => new RegExp(`^${worker}-[\\w-]+\\.m?js$`).test(basename(path)));
    assert.ok(entries.length > 0, `Missing emitted ${worker} entry chunk`);
    for (const entry of entries) {
      const importers = scripts.filter(path => path !== entry && imports.get(path).includes(basename(entry)));
      if (importers.length > 0) {
        offenders.push(`${relative(directory, entry)} <- ${importers.map(path => relative(directory, path)).join(', ')}`);
      }
    }
  }
  assert.deepEqual(offenders, [],
    `Worker chunks statically import their worker entry, so a WebKit worker evaluates it twice:\n${offenders.join('\n')}`);
}

async function checkPublishedWorkerStubs() {
  const failures = [];
  for (const [format, file] of publishedWorkers) {
    const dist = join(root, 'packages', format, 'dist');
    const path = join(root, 'packages', format, file);
    const contents = await readFile(path, 'utf8').catch(() => null);
    if (contents === null) {
      failures.push(`${format} ${file} is missing`);
      continue;
    }
    const imported = [...contents.matchAll(staticImport)].map(([, specifier]) => specifier);
    if (imported.length > 0) failures.push(`${format} ${file} statically imports ${imported.join(', ')}; it must only import its code dynamically`);
    const dynamic = contents.match(dynamicImport)?.length ?? 0;
    if (dynamic !== 1) failures.push(`${format} ${file} has ${dynamic} dynamic imports; it must load its code with exactly one`);
    for (const [, target] of contents.matchAll(dynamicTarget)) {
      const resolved = resolve(dirname(path), target);
      const inside = !relative(dist, resolved).startsWith('..');
      if (!inside || !await access(resolved).then(() => true, () => false)) {
        failures.push(`${format} ${file} dynamically imports ${target}, which is not a file in ${format}/dist`);
      }
    }
  }
  assert.deepEqual(failures, [], `Published worker entries must import nothing statically and one file of their dist dynamically:\n${failures.join('\n')}`);
}

try {
  await checkPublishedWorkerStubs();
  await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  const workers = {};
  for (const [format, worker] of Object.entries(formats)) {
    const source = join(root, 'packages', format);
    const target = join(directory, 'node_modules/@betteroffice', format);
    await mkdir(target, { recursive: true });
    await cp(join(source, 'package.json'), join(target, 'package.json'));
    await cp(join(source, 'dist'), join(target, 'dist'), { recursive: true });
    const manifest = JSON.parse(await readFile(join(target, 'package.json'), 'utf8'));
    assert.deepEqual(Object.keys(manifest.imports), ['#session-worker']);
    assert.match(manifest.imports['#session-worker'].default, /sessionWorker\.missing\.(?:js|mjs)$/);
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
    const missing = join(directory, 'node_modules/@betteroffice', format, 'dist/sessionWorker.missing');
    assert.ok(Object.keys(bundled.metafile.inputs).some(path => resolve(root, path).startsWith(missing)),
      `esbuild did not select the ${format} browser variant`);
  }
  assert.ok(Object.keys(bundled.metafile.inputs).every(path => !/sessionWorker\.node|SessionWorker\.node|docxWorker\.node|nodeSession(?:Worker|Scope)/.test(path)));
  for (const [worker, path] of Object.entries(workers)) {
    const bare = [];
    const alone = await esbuild({
      absWorkingDir: root, entryPoints: [path], write: false,
      platform: 'browser', conditions: ['browser', 'default'],
      bundle: true, format: 'esm', metafile: true, logLevel: 'silent',
      plugins: [{ name: 'bare-imports', setup(build) {
        build.onResolve({ filter: /^[^./]/ }, ({ path: specifier, kind }) => {
          bare.push(`${specifier} (${kind})`);
          return { path: specifier, external: true };
        });
      } }],
    });
    const inputs = Object.keys(alone.metafile.inputs);
    assert.ok(inputs.every(input => !/sessionWorker\.|nodeSession(?:Worker|Scope)/.test(input)), `${worker} bundles a module that constructs its own worker`);
    assert.deepEqual(bare.filter(entry => !/^@betteroffice\/fonts(?:\/progressive)? \(dynamic-import\)$/.test(entry)), [],
      `${worker} imports specifiers an unbundled browser worker cannot resolve`);
  }
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
  await checkWorkerEntriesLoadOnce(viteDirectory);

  const webpack = nextWebpack();
  const bundle = (context, app, output) => new Promise((done, fail) => webpack({
    mode: 'production', target: 'web', devtool: false, context,
    entry: { app },
    output: { path: output, publicPath: 'auto', filename: '[name].js', chunkFilename: '[name].[contenthash].js' },
    resolve: { conditionNames: ['browser', 'import', 'module', 'default'] },
    optimization: { minimize: false },
  }, (error, result) => {
    if (error) return fail(error);
    if (result.hasErrors()) return fail(new Error(result.toString({ all: false, errors: true })));
    done(result.compilation);
  }));
  const workerEntry = (compilation, file) => {
    const { chunkGraph } = compilation;
    const initial = new Set(compilation.entrypoints.get('app').chunks);
    for (const module of compilation.modules) {
      const inner = (module.modules ? [...module.modules] : [module]).find(candidate => candidate.resource?.endsWith(file));
      if (!inner || !inner.type.startsWith('javascript')) continue;
      const chunks = [...chunkGraph.getModuleChunks(module)];
      return chunks.some(chunk => !initial.has(chunk) && [...chunkGraph.getChunkEntryModulesIterable(chunk)].includes(module)) &&
        !chunks.some(chunk => initial.has(chunk));
    }
    return false;
  };
  const fixture = join(directory, 'webpack-fixture');
  await mkdir(fixture, { recursive: true });
  await writeFile(join(fixture, 'fixtureWorker.mjs'), 'self.postMessage(1);\n');
  await writeFile(join(fixture, 'worker.mjs'), "new Worker(new URL('./fixtureWorker.mjs', import.meta.url), { type: 'module' });\n");
  await writeFile(join(fixture, 'plain.mjs'), "import './fixtureWorker.mjs';\nglobalThis.started = typeof Worker;\n");
  assert.ok(workerEntry(await bundle(fixture, join(fixture, 'worker.mjs'), join(fixture, 'worker')), '/fixtureWorker.mjs'),
    'the webpack worker check misses a real worker entry');
  assert.ok(!workerEntry(await bundle(fixture, join(fixture, 'plain.mjs'), join(fixture, 'plain')), '/fixtureWorker.mjs'),
    'the webpack worker check accepts a plain import');
  const webpackDirectory = join(directory, 'webpack');
  const compilation = await bundle(directory, entry, webpackDirectory);
  for (const worker of Object.values(formats)) {
    assert.ok(workerEntry(compilation, `/dist/${worker}.mjs`), `webpack did not emit ${worker} as its own worker entry chunk`);
  }
  await checkOutput(webpackDirectory, { namedWorkers: false });
  console.log('esbuild, Vite and webpack browser worker bundles exclude Node worker code and private imports.');
} finally {
  await rm(directory, { recursive: true, force: true });
}
