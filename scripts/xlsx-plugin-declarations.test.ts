import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');
const PACKAGES = ['xlsx-i18n', 'xlsx', 'xlsx-react'];
const LINKED = ['react', 'react-dom', '@types/react', '@types/react-dom', 'csstype'];

const CONSUMER = `import {
  ToolbarCommandButton,
  XlsxCommandAdmissionError,
  XlsxEditor,
  XlsxPluginToolbar,
  defineXlsxPlugin,
  useXlsxCommand,
  type XlsxEditorApi,
  type XlsxPluginContext,
  type XlsxPluginGrant,
  type XlsxPluginNavigationResult,
} from '@betteroffice/xlsx-react';

type State = { count: number };

function Panel({ context }: { context: XlsxPluginContext<State> }) {
  const mark = useXlsxCommand('plugin:acme.review/mark');
  return <button onClick={() => void mark.execute()}>{context.state.count} {mark.label}</button>;
}

export const review = defineXlsxPlugin<State>({
  id: 'acme.review',
  createState: () => ({ count: 0 }),
  async onEvent(context, event) {
    if (event.type === 'selection-change' && event.selection?.cells) {
      const result: XlsxPluginNavigationResult = await context.navigation.selectCells(
        { sheetId: event.selection.sheetId, selection: event.selection.cells },
        { expectVersion: event.version, focus: false }
      );
      void result;
    }
    if (event.type !== 'document-change') return;
    const read = await context.read.readCells({ ranges: [] });
    if (read.ok) context.setState({ count: read.sheets.length }, read.version);
  },
  panel: { title: 'Review', placement: 'right', render: Panel },
  overlay: ({ geometry }) => (
    <div style={{ left: geometry.getCellRect({ sheetId: geometry.layout.sheetId, row: 0, col: 0 })?.x }} />
  ),
  commands: [{ id: 'mark', label: 'Mark', mutatesDocument: false, execute: () => ({ ok: true, status: 'executed' }) }],
  toolbar: ['mark'],
});

const grant: XlsxPluginGrant = { document: 'write', editBatches: true, commands: ['zoom'] };
// @ts-expect-error edit batches need document: 'write'
const invalid: XlsxPluginGrant = { editBatches: true };
void invalid;

export async function version(api: XlsxEditorApi): Promise<string | null> {
  try {
    return await api.version();
  } catch (error) {
    if (error instanceof XlsxCommandAdmissionError && error.code === 'input-failed') return null;
    throw error;
  }
}

export const editor = (
  <XlsxEditor
    plugins={[review]}
    pluginGrants={{ 'acme.review': grant }}
    onPluginError={(error) => console.error(error.phase)}
    toolbar={
      <>
        <XlsxPluginToolbar />
        <ToolbarCommandButton id="plugin:acme.review/mark" />
        <ToolbarCommandButton id="bold" />
      </>
    }
  />
);
`;

function run(command: string[], cwd: string) {
  const result = spawnSync(process.execPath, command, {
    cwd,
    encoding: 'utf8',
    env: { ...process.env, NODE_OPTIONS: '--max-old-space-size=8192' },
  });
  if (result.status !== 0) throw new Error(result.stdout + result.stderr);
}

test('the XLSX plugin API and the demo plugin typecheck against built declarations', () => {
  const directory = mkdtempSync(join(tmpdir(), 'betteroffice-xlsx-plugins-'));
  try {
    for (const name of PACKAGES) {
      const source = join(root, 'packages', name);
      const manifest = JSON.parse(readFileSync(join(source, 'package.json'), 'utf8'));
      const destination = join(directory, 'node_modules', manifest.name);
      mkdirSync(destination, { recursive: true });
      writeFileSync(join(destination, 'package.json'), JSON.stringify(manifest));
      run(['x', 'tsup', '--out-dir', join(destination, 'dist')], source);
    }
    for (const name of LINKED) {
      const target = join(directory, 'node_modules', name);
      mkdirSync(dirname(target), { recursive: true });
      symlinkSync(join(root, 'node_modules', name), target, 'dir');
    }
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ type: 'module' }));
    writeFileSync(join(directory, 'consumer.tsx'), CONSUMER);
    copyFileSync(
      join(root, 'apps/demo/app/xlsx/ReviewPlugin.tsx'),
      join(directory, 'ReviewPlugin.tsx')
    );
    writeFileSync(
      join(directory, 'tsconfig.json'),
      JSON.stringify({
        compilerOptions: {
          target: 'ES2022',
          module: 'NodeNext',
          moduleResolution: 'NodeNext',
          lib: ['ES2022', 'DOM', 'DOM.Iterable'],
          jsx: 'react-jsx',
          strict: true,
          skipLibCheck: false,
          noEmit: true,
          types: [],
        },
        files: ['consumer.tsx', 'ReviewPlugin.tsx'],
      })
    );
    const tsc = createRequire(join(root, 'packages/xlsx-react/package.json')).resolve(
      'typescript/bin/tsc'
    );
    const check = spawnSync(process.execPath, [tsc, '-p', directory], {
      cwd: directory,
      encoding: 'utf8',
    });
    const declarations = readFileSync(
      join(directory, 'node_modules/@betteroffice/xlsx-react/dist/index.d.ts'),
      'utf8'
    );
    expect({
      status: check.status,
      output: check.stdout + check.stderr,
      privateImports: declarations.match(/from ['"](?:\.\.?\/|[^'"]*shared\/)[^'"]*['"]/g),
    }).toEqual({ status: 0, output: '', privateImports: null });
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}, 180_000);
