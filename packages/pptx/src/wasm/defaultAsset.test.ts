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

const REJECT_FILE_FETCH = `
const realFetch = globalThis.fetch;
globalThis.fetch = (input, init) => {
  const href = typeof Request === 'function' && input instanceof Request ? input.url : String(input);
  return href.startsWith('file:') ? Promise.reject(new TypeError('fetch failed')) : realFetch(input, init);
};
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
