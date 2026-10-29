import { expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';
import { build } from 'esbuild';
import ts from 'typescript';

const root = fileURLToPath(new URL('../../../', import.meta.url));
// 74,976 B measured; the native engine selector (packages/xlsx/src/wasm/engine.ts) adds ~0.9 KB to the static path (75,880 B).
const STATIC_GZIP_BUDGET = 76_000;
const deferredModules = ['XlsxEditorContent.tsx', 'LegacyXlsxEditor.tsx', 'XlsxAsyncInThreadSession.tsx',
  'viewer/XlsxSessionViewer.tsx', 'viewer/useSessionWorkbook.ts', 'createInThreadAsyncEditorApi.ts',
  'commands/useXlsxCommands.ts', 'commands/inputCoordinator.ts', 'clipboard.ts', 'presence/Presence.tsx',
  'worker/LegacyXlsxWorkerEditor.tsx', 'worker/useEditableSessionWorkbook.ts', 'worker/createWorkerEditorApi.ts',
  'worker/WorkerPaintSource.ts', 'commands/useWorkerXlsxCommands.ts', 'viewer/sessionGeometry.ts',
  'worker/XlsxPeerlessWorkerEditor.tsx', 'worker/usePeerlessSessionWorkbook.ts', 'worker/PeerlessWorkerPaintSource.ts',
  'worker/PeerlessWorkerSurface.tsx', 'worker/PeerlessWorkerInputs.tsx', 'worker/WorkerDraftController.ts',
  'worker/WorkerPresentationStore.ts', 'worker/createAsyncEditorApi.ts', 'commands/peerlessWorkerInputCoordinator.ts',
  'commands/usePeerlessWorkerXlsxCommands.ts', 'plugins/createPeerlessPluginDocument.ts', 'plugins/createPeerlessPluginClients.ts']
  .map((path) => `packages/xlsx-react/src/${path}`);

test('default editor static chunks exclude the legacy and in-thread editors', async () => {
  const result = await build({
    absWorkingDir: root,
    stdin: { contents: 'export { XlsxEditor } from "./packages/xlsx-react/src/index.ts";', resolveDir: root },
    bundle: true, minify: true, format: 'esm', splitting: true,
    outdir: resolve(root, 'node_modules/.cache/xlsx-imports'), write: false, metafile: true,
    external: ['react', 'react-dom'],
  });
  const outputs = result.metafile.outputs;
  const entry = Object.keys(outputs).find((path) => outputs[path].entryPoint === '<stdin>')!;
  const seen = new Set<string>();
  const visit = (path: string): void => {
    if (seen.has(path)) return;
    seen.add(path);
    for (const dependency of outputs[path].imports) {
      if (!dependency.external && dependency.kind !== 'dynamic-import') visit(dependency.path);
    }
  };
  visit(entry);
  const inputs = (paths: string[]) => new Set(paths.flatMap((path) => Object.entries(outputs[path].inputs)
    .filter(([, input]) => input.bytesInOutput > 0).map(([path]) => path)));
  const staticInputs = inputs([...seen]);
  const deferredInputs = inputs(Object.keys(outputs).filter((path) => !seen.has(path)));
  expect(deferredModules.map((path) => [path, staticInputs.has(path), deferredInputs.has(path)]))
    .toEqual(deferredModules.map((path) => [path, false, true]));
  expect(staticInputs.has('packages/xlsx-react/src/worker/XlsxWorkerEditor.tsx')).toBe(true);
  const files = new Map(result.outputFiles.map((file) => [file.path, file.contents]));
  const bytes = [...seen].reduce((sum, path) => sum + gzipSync(files.get(resolve(root, path))!, { level: 9 }).length, 0);
  expect(bytes).toBeLessThan(STATIC_GZIP_BUDGET);
});

function modulePath(from: string, specifier: string): string {
  const base = resolve(dirname(from), specifier);
  const stem = base.replace(/\.(?:js|mjs)$/, '');
  for (const path of [base, `${stem}.ts`, `${stem}.tsx`, `${base}/index.ts`, `${base}/index.tsx`]) {
    if (/\.(?:ts|tsx|js|mjs)$/.test(path) && existsSync(path)) return path;
  }
  throw new Error(`Unresolved source: ${relative(root, from)} -> ${specifier}`);
}

test('deferred editors have no runtime imports back into the editor shell', () => {
  const seen = new Set<string>();
  const visit = (path: string): void => {
    expect(relative(root, path)).not.toMatch(/^packages\/xlsx-react\/src\/(?:XlsxEditor|worker\/XlsxWorkerEditor)\.tsx$/);
    if (seen.has(path)) return;
    seen.add(path);
    const source = ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true);
    const walk = (node: ts.Node): void => {
      if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier) &&
        !node.importClause?.isTypeOnly && node.moduleSpecifier.text.startsWith('.')) {
        const bindings = node.importClause?.namedBindings;
        if (!node.importClause?.name && bindings && ts.isNamedImports(bindings) &&
          bindings.elements.every((entry) => entry.isTypeOnly)) return;
        visit(modulePath(path, node.moduleSpecifier.text));
      } else if (ts.isExportDeclaration(node) && !node.isTypeOnly && node.moduleSpecifier &&
        ts.isStringLiteral(node.moduleSpecifier) && node.moduleSpecifier.text.startsWith('.')) {
        if (node.exportClause && ts.isNamedExports(node.exportClause) &&
          node.exportClause.elements.every((entry) => entry.isTypeOnly)) return;
        visit(modulePath(path, node.moduleSpecifier.text));
      } else if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword &&
        node.arguments[0] && ts.isStringLiteral(node.arguments[0]) && node.arguments[0].text.startsWith('.')) {
        visit(modulePath(path, node.arguments[0].text));
      }
      ts.forEachChild(node, walk);
    };
    walk(source);
  };
  for (const path of ['LegacyXlsxEditor.tsx', 'XlsxAsyncInThreadSession.tsx', 'worker/LegacyXlsxWorkerEditor.tsx']) {
    seen.clear();
    visit(resolve(root, `packages/xlsx-react/src/${path}`));
    expect(seen.has(resolve(root, 'packages/xlsx-react/src/viewer/XlsxSessionViewer.tsx'))).toBe(path === 'LegacyXlsxEditor.tsx');
  }
});
