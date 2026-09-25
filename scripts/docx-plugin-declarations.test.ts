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
const PACKAGES = ['docx-i18n', 'docx', 'docx-react'];
const LINKED = ['react', 'react-dom', '@types/react', '@types/react-dom', 'csstype'];

const CONSUMER = `import {
  DocxEditor,
  DocxPluginToolbar,
  ToolbarCommandButton,
  defineDocxPlugin,
  useDocxCommand,
  type DocxPluginContext,
  type DocxPluginGrant,
  type DocxPluginSidebarItem,
} from '@betteroffice/docx-react';

type State = { count: number };

function Panel({ context }: { context: DocxPluginContext<State> }) {
  const mark = useDocxCommand('plugin:acme.review/mark');
  return <button onClick={() => void mark.execute()}>{context.state.count} {mark.label}</button>;
}

const card: DocxPluginSidebarItem<State> = {
  id: 'card',
  anchor: { version: 'v', story: 'body', paraId: '00000001' },
  render: ({ isExpanded }) => <p>{String(isExpanded)}</p>,
};

export const review = defineDocxPlugin<State>({
  id: 'acme.review',
  createState: () => ({ count: 0 }),
  async onEvent(context, event) {
    if (event.type !== 'document-change') return;
    const read = await context.read.readParagraphs({ view: 'accepted' });
    if (read.ok) context.setState({ count: read.paragraphs.length }, read.version);
  },
  panel: { title: 'Review', placement: 'right', render: Panel },
  overlay: ({ geometry }) => <div style={{ left: geometry.toOverlayRect({ x: 0, y: 0, width: 1, height: 1 }).x }} />,
  getSidebarItems: () => [card],
  commands: [{ id: 'mark', label: 'Mark', mutatesDocument: false, execute: () => ({ ok: true, status: 'executed' }) }],
  toolbar: ['mark'],
});

const grant: DocxPluginGrant = { document: 'write', editBatches: true, commands: ['bold'] };
// @ts-expect-error edit batches need document: 'write'
const invalid: DocxPluginGrant = { editBatches: true };
void invalid;

export const editor = (
  <DocxEditor
    plugins={[review]}
    pluginGrants={{ 'acme.review': grant }}
    onPluginError={(error) => console.error(error.phase)}
    toolbar={
      <>
        <DocxPluginToolbar />
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

test('the DOCX plugin API and the demo plugin typecheck against built declarations', () => {
  const directory = mkdtempSync(join(tmpdir(), 'betteroffice-docx-plugins-'));
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
      join(root, 'apps/demo/app/docx/ReviewPlugin.tsx'),
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
    const tsc = createRequire(join(root, 'packages/docx-react/package.json')).resolve(
      'typescript/bin/tsc'
    );
    const check = spawnSync(process.execPath, [tsc, '-p', directory], {
      cwd: directory,
      encoding: 'utf8',
    });
    const declarations = readFileSync(
      join(directory, 'node_modules/@betteroffice/docx-react/dist/index.d.mts'),
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
