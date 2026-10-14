import { expect, test } from 'bun:test';
import { build } from 'esbuild';
import { resolve } from 'node:path';

const root = resolve(import.meta.dir, '../../..');

test('the default editor loads peerless modules only through a dynamic chunk', async () => {
  const result = await build({
    stdin: { contents: 'export { XlsxEditor } from "./packages/xlsx-react/src/index.ts";',
      resolveDir: root, sourcefile: 'default-editor.ts' },
    bundle: true, write: false, minify: true, format: 'esm', platform: 'browser',
    splitting: true, outdir: resolve(root, '.bundle-test'), metafile: true,
    external: ['react', 'react-dom', 'react/jsx-runtime', '@betteroffice/xlsx', '@betteroffice/xlsx/*', '@betteroffice/xlsx-i18n'],
    logLevel: 'silent',
  });
  const outputs = result.metafile!.outputs;
  const entry = Object.keys(outputs).find((file) => outputs[file].entryPoint === 'default-editor.ts')!;
  const visited = new Set<string>();
  const visit = (file: string) => {
    if (visited.has(file)) return;
    visited.add(file);
    for (const imported of outputs[file].imports) {
      if (!imported.external && imported.kind !== 'dynamic-import') visit(imported.path);
    }
  };
  visit(entry);
  const eager = [...visited].flatMap((file) => Object.keys(outputs[file].inputs));
  for (const module of ['XlsxPeerlessWorkerEditor', 'PeerlessWorkerSurface', 'PeerlessWorkerInputs',
    'PeerlessWorkerPaintSource', 'WorkerDraftController', 'WorkerPresentationStore', 'createAsyncEditorApi',
    'createPeerlessPluginDocument', 'createPeerlessPluginClients', 'usePeerlessSessionWorkbook',
    'peerlessWorkerInputCoordinator', 'usePeerlessWorkerXlsxCommands']) {
    expect(eager.some((file) => file.includes(module))).toBe(false);
  }
  expect(Object.values(outputs).some((output) => output.imports.some((imported) => imported.kind === 'dynamic-import'))).toBe(true);
});
