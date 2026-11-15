import { afterAll, expect, test } from 'bun:test';
import { cpSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { engineKind } from './engine';

const root = resolve(import.meta.dir, '../../../..');
const loader = JSON.stringify(resolve(import.meta.dir, 'loader.ts'));
const fixture = JSON.stringify(resolve(root, 'apps/demo/public/betteroffice-demo.pptx'));
const native = engineKind === 'native';
const asset = JSON.stringify(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm'));

const RECORDING_FETCH = `
import { readFileSync } from 'node:fs';
const wasmBytes = new Uint8Array(readFileSync(${asset}));
const requested = [];
const serveWasm = (input) => {
  requested.push(String(input));
  return Promise.resolve(new Response(wasmBytes, { headers: { 'Content-Type': 'application/wasm' } }));
};
const rejectAll = (input) => {
  requested.push(String(input));
  return Promise.reject(new TypeError('fetch failed'));
};
const hideNodeBuiltins = () => Object.defineProperty(process, 'getBuiltinModule', {
  value: undefined, configurable: true, writable: true,
});
const assetNames = () => JSON.stringify(requested.map((href) => href.slice(0, 5) + href.slice(href.lastIndexOf('/') + 1)));
`;

const REJECT_FILE_FETCH = `
const realFetch = globalThis.fetch;
globalThis.fetch = (input, init) => {
  const href = typeof Request === 'function' && input instanceof Request ? input.url : String(input);
  return href.startsWith('file:') ? Promise.reject(new TypeError('fetch failed')) : realFetch(input, init);
};
`;

const OBSERVED_FETCH = `
const realFetch = globalThis.fetch;
const requested = [];
globalThis.fetch = (input, init) => {
  requested.push(typeof Request === 'function' && input instanceof Request ? input.url : String(input));
  return realFetch(input, init);
};
`;

const OPAQUE_MODULE_BASE = `
Bun.plugin({
  name: 'opaque-module-base',
  setup(build) {
    build.onLoad({ filter: /[\\\\/]wasm[\\\\/]asset\\.ts$/ }, async ({ path }) => ({
      contents: (await Bun.file(path).text()).replaceAll('import.meta.url', "'blob:https://app.example/4f1c'"),
      loader: 'ts',
    }));
  },
});
`;

async function run(command: string[], script: string) {
  const child = Bun.spawn([...command, script], {
    cwd: root, env: process.env, stdout: 'pipe', stderr: 'pipe',
  });
  const [stdout, stderr, exit] = await Promise.all([
    new Response(child.stdout).text(), new Response(child.stderr).text(), child.exited,
  ]);
  return { exit, stdout: stdout.trim(), stderr: stderr.trim() };
}

const bun = (script: string) => run([process.execPath, '--eval'], script);
const node = (script: string) => run(['node', '--input-type=module', '--eval'], script);

let deployment: string | undefined;
let entry: Promise<string> | undefined;

function bundledEntry(): Promise<string> {
  entry ??= (async () => {
    deployment = mkdtempSync(join(tmpdir(), 'pptx-node-entry-'));
    const build = await Bun.build({
      entrypoints: [resolve(import.meta.dir, '../index.ts')],
      target: 'node',
      format: 'esm',
      outdir: deployment,
      external: ['#session-worker', '@betteroffice/*'],
    });
    if (!build.success) throw new Error(build.logs.map(String).join('\n'));
    writeFileSync(join(deployment, 'package.json'), '{ "type": "module" }\n');
    cpSync(resolve(import.meta.dir, 'generated'), join(deployment, 'generated'), { recursive: true });
    return JSON.stringify(pathToFileURL(join(deployment, 'index.js')).href);
  })();
  return entry;
}

afterAll(() => {
  if (deployment) rmSync(deployment, { recursive: true, force: true });
});

test.skipIf(native)('initWasm() without an input resolves on Bun when fetch rejects file: URLs', async () => {
  const result = await bun(`${REJECT_FILE_FETCH}
    const { initWasm, wasmVersion } = await import(${loader});
    await initWasm();
    process.stdout.write(typeof wasmVersion());
  `);
  expect(result).toMatchObject({ exit: 0, stdout: 'string' });
}, 60_000);

test.skipIf(native)('initWasm() without an input on Bun reads the packaged asset without fetching it', async () => {
  const result = await bun(`${OBSERVED_FETCH}
    const { initWasm, wasmVersion } = await import(${loader});
    await initWasm();
    process.stdout.write(JSON.stringify({ requested, version: typeof wasmVersion() }));
  `);
  expect(result).toMatchObject({ exit: 0, stdout: '{"requested":[],"version":"string"}' });
}, 60_000);

test.skipIf(native)('initWasm() without an input resolves on Node from the bundled entry', async () => {
  const result = await node(`
    const { initWasm, wasmVersion } = await import(${await bundledEntry()});
    await initWasm();
    process.stdout.write(typeof wasmVersion());
  `);
  expect(result).toMatchObject({ exit: 0, stdout: 'string' });
}, 120_000);

test.skipIf(native)('exportPptxMarkdown initializes the engine on Node without a prior initWasm call', async () => {
  const result = await node(`
    import { readFileSync } from 'node:fs';
    const { exportPptxMarkdown } = await import(${await bundledEntry()});
    const content = await exportPptxMarkdown(new Uint8Array(readFileSync(${fixture})));
    process.stdout.write(typeof content.markdown);
  `);
  expect(result).toMatchObject({ exit: 0, stdout: 'string' });
}, 120_000);

test.skipIf(native)('an explicit URL input reaches fetch unchanged', async () => {
  const result = await bun(`${RECORDING_FETCH}
    globalThis.fetch = serveWasm;
    const { initWasm } = await import(${loader});
    await initWasm(new URL('https://cdn.example/pptx_wasm_bg.wasm'));
    process.stdout.write(JSON.stringify(requested));
  `);
  expect(result).toMatchObject({ exit: 0, stdout: '["https://cdn.example/pptx_wasm_bg.wasm"]' });
}, 60_000);

test.skipIf(native)('explicit bytes reach the engine untouched without a fetch', async () => {
  const result = await bun(`${RECORDING_FETCH}
    globalThis.fetch = rejectAll;
    const instantiate = WebAssembly.instantiate;
    const sources = [];
    WebAssembly.instantiate = (source, imports) => {
      sources.push(source);
      return instantiate(source, imports);
    };
    const { initWasm } = await import(${loader});
    await initWasm(wasmBytes);
    process.stdout.write(JSON.stringify({ requested, same: sources[0] === wasmBytes }));
  `);
  expect(result).toMatchObject({ exit: 0, stdout: '{"requested":[],"same":true}' });
}, 60_000);

test.skipIf(native)('explicit bytes initialize when the packaged asset URL cannot be resolved', async () => {
  const result = await bun(`${RECORDING_FETCH}${OPAQUE_MODULE_BASE}
    globalThis.fetch = rejectAll;
    const { initWasm, wasmVersion } = await import(${loader});
    await initWasm(wasmBytes);
    process.stdout.write(JSON.stringify({ requested, version: typeof wasmVersion() }));
  `);
  expect(result).toMatchObject({ exit: 0, stdout: '{"requested":[],"version":"string"}' });
}, 60_000);

test.skipIf(native)('without the Node builtin loader the packaged asset is fetched', async () => {
  const result = await bun(`${RECORDING_FETCH}
    hideNodeBuiltins();
    globalThis.fetch = serveWasm;
    const { initWasm } = await import(${loader});
    await initWasm();
    process.stdout.write(assetNames());
  `);
  expect(result).toMatchObject({ exit: 0, stdout: '["file:pptx_wasm_bg.wasm"]' });
}, 60_000);

test.skipIf(native)('a packaged asset that cannot load names the asset and the bytes input', async () => {
  const result = await bun(`${RECORDING_FETCH}
    hideNodeBuiltins();
    globalThis.fetch = rejectAll;
    const { initWasm } = await import(${loader});
    const error = await initWasm().then(() => undefined, (failure) => failure);
    process.stdout.write(String(error?.message));
  `);
  expect(result.exit).toBe(0);
  expect(result.stdout).toMatch(/^Could not load the packaged WASM asset file:\S*\/pptx_wasm_bg\.wasm \(fetch failed\); pass its bytes or a compiled WebAssembly\.Module to initWasm\(input\)$/);
}, 60_000);
