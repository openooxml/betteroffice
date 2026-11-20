import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import {
  copyFileSync,
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  realpathSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');
const declarationMode = process.env.BETTEROFFICE_PACKAGE_DECLARATIONS;
const PACKAGES = ['docx-i18n', 'docx', 'docx-react', 'fonts'];
const LINKED = ['react', 'react-dom', '@types/react', '@types/react-dom', 'csstype'];
const DEMOS = ['ReviewPlugin.tsx', 'CompactToolbar.tsx'];

const CONSUMER = `import {
  DocxEditor,
  DocxPluginToolbar,
  ToolbarCommandButton,
  defineDocxPlugin,
  useDocxCommand,
  type DocxPluginContext,
  type DocxPluginSelectionSettleEvent,
  type DocxGeometryTarget,
  type DocxPluginRect,
  type DocxPluginGrant,
  type DocxPluginSidebarItem,
} from '@betteroffice/docx-react';
import {
  DOCX_MINI_TOOLBAR_ID,
  DOCX_MINI_TOOLBAR_ITEMS,
  docxMiniToolbar,
  type DocxMiniToolbarItem,
} from '@betteroffice/docx-react/mini-toolbar';

const miniItems: readonly DocxMiniToolbarItem[] = ['clearFormatting', 'pasteAsPlainText'];
const miniToolbar = docxMiniToolbar({ items: miniItems });
// @ts-expect-error mini toolbar items are built-in command ids it supports
docxMiniToolbar({ items: ['insertLink'] });
void DOCX_MINI_TOOLBAR_ITEMS;

type State = { count: number };

function Panel({ context }: { context: DocxPluginContext<State> }) {
  const mark = useDocxCommand('plugin:acme.review/mark');
  return <button onClick={() => void mark.execute()} onKeyDown={() => { const answer: { ok: true } | { ok: false } = context.focusDocument(); void answer; }}>{context.state.count} {mark.label}</button>;
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
    const focused: boolean = context.snapshot.focused;
    const locale: string = context.snapshot.locale;
    const settled: boolean = context.snapshot.selectionSettled;
    const inputFocused: boolean = context.snapshot.inputFocused;
    const target: DocxGeometryTarget = { kind: 'selection' };
    const anchor = context.geometry?.getAnchorGeometry(target);
    const anchors = await context.geometry?.readAnchorGeometries([target]);
    void [focused, locale, settled, inputFocused, anchor, anchors];
    const viewport: DocxPluginRect | undefined = context.geometry?.viewport;
    void viewport;
    switch (event.type) {
      case 'selection-settle': {
        const settled: DocxPluginSelectionSettleEvent = event;
        const source: 'pointer' | 'touch' | 'keyboard' = settled.source;
        const publicSources: Record<typeof settled.source, true> = { pointer: true, touch: true, keyboard: true };
        void [source, publicSources, settled.version, settled.selection];
        break;
      }
      case 'focus-change': {
        const focused: boolean = event.focused;
        void focused;
        break;
      }
      case 'locale-change': {
        const locale: string = event.locale;
        void locale;
        break;
      }
      case 'viewport-change': {
        const viewport: DocxPluginRect | null = event.viewport;
        void viewport;
        break;
      }
    }
    if (event.type !== 'document-change') return;
    const read = await context.read.readParagraphs({ view: 'accepted' });
    if (read.ok) context.setState({ count: read.paragraphs.length }, read.version);
  },
  panel: { title: 'Review', placement: 'right', render: Panel },
  overlay: ({ geometry }) => <div style={{ left: geometry.toOverlayRect({ x: 0, y: 0, width: 1, height: 1 })?.x }} />,
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
    plugins={[review, miniToolbar]}
    pluginGrants={{ 'acme.review': grant, [DOCX_MINI_TOOLBAR_ID]: { document: 'write', commands: miniItems } }}
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

function specifiers(code: string): string[] {
  return [...code.matchAll(/(?:\bfrom|\bimport)\s*\(?\s*["']([^"']+)["']/g)].map((match) => match[1]!);
}

function chunkFile(dist: string, specifier: string, declarations: boolean): string | null {
  if (!/^\.\/[^/]+$/.test(specifier)) return null;
  const name = specifier.slice(2);
  const candidates = declarations ? [name, name.replace(/\.m?js$/, '.d.mts'), `${name}.d.mts`] : [name];
  return candidates.find((file) => file.endsWith(declarations ? '.d.mts' : '.mjs') && existsSync(join(dist, file))) ?? null;
}

function moduleGraph(dist: string, entry: string): Set<string> {
  const files = new Set<string>();
  const visit = (file: string) => {
    if (files.has(file)) return;
    files.add(file);
    for (const specifier of specifiers(readFileSync(join(dist, file), 'utf8'))) {
      const chunk = chunkFile(dist, specifier, file.endsWith('.d.mts'));
      if (chunk) visit(chunk);
    }
  };
  visit(entry);
  return files;
}

function filesContaining(dist: string, extension: string, pattern: RegExp): string[] {
  return readdirSync(dist).filter((file) => file.endsWith(extension) && pattern.test(readFileSync(join(dist, file), 'utf8')));
}

test.skipIf(declarationMode === 'skip')('the DOCX plugin API and the demo consumers typecheck against built declarations', () => {
  const directory = mkdtempSync(join(tmpdir(), 'betteroffice-docx-plugins-'));
  try {
    for (const name of PACKAGES) {
      const source = join(root, 'packages', name);
      const manifest = JSON.parse(readFileSync(join(source, 'package.json'), 'utf8'));
      const destination = join(directory, 'node_modules', manifest.name);
      mkdirSync(destination, { recursive: true });
      writeFileSync(join(destination, 'package.json'), JSON.stringify(manifest));
      if (declarationMode === 'prebuilt') cpSync(join(source, 'dist'), join(destination, 'dist'), { recursive: true });
      else run(['x', 'tsup', '--out-dir', join(destination, 'dist')], source);
    }
    for (const name of LINKED) {
      const target = join(directory, 'node_modules', name);
      mkdirSync(dirname(target), { recursive: true });
      symlinkSync(join(root, 'node_modules', name), target, 'dir');
    }
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ type: 'module' }));
    writeFileSync(join(directory, 'consumer.tsx'), CONSUMER);
    for (const demo of DEMOS) {
      copyFileSync(join(root, 'apps/demo/app/docx', demo), join(directory, demo));
    }
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
        files: ['consumer.tsx', ...DEMOS],
      })
    );
    const tsc = createRequire(join(root, 'packages/docx-react/package.json')).resolve(
      'typescript/bin/tsc'
    );
    const check = spawnSync(process.execPath, [tsc, '-p', directory], {
      cwd: directory,
      encoding: 'utf8',
    });
    const dist = join(directory, 'node_modules/@betteroffice/docx-react/dist');
    expect({
      status: check.status,
      output: check.stdout + check.stderr,
      privateImports: [...new Set([...moduleGraph(dist, 'index.d.mts'), ...moduleGraph(dist, 'mini-toolbar.d.mts')])]
        .flatMap((file) => specifiers(readFileSync(join(dist, file), 'utf8')))
        .filter((specifier) => specifier.includes('shared/') || (specifier.startsWith('.') && !chunkFile(dist, specifier, true))),
    }).toEqual({ status: 0, output: '', privateImports: [] });
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}, 180_000);

test.skipIf(declarationMode === 'skip')('the built mini toolbar subpath shares the editor chunks instead of a copy or the package root', () => {
  const directory = mkdtempSync(join(tmpdir(), 'betteroffice-docx-mini-toolbar-'));
  try {
    const source = join(root, 'packages/docx-react');
    const manifestText = readFileSync(join(source, 'package.json'), 'utf8');
    const manifest = JSON.parse(manifestText) as { exports: Record<string, unknown>; scripts: Record<string, string> };
    const built = join(directory, 'built');
    if (declarationMode === 'prebuilt') cpSync(join(source, 'dist'), built, { recursive: true });
    else {
      const steps = manifest.scripts.build!.split('&&').map((step) => step.trim().split(/\s+/));
      for (const step of steps.filter(([tool]) => tool === 'tsup')) run(['x', ...step, '--out-dir', built], source);
    }

    expect(JSON.stringify(manifest.exports)).not.toContain('"require"');
    expect(readdirSync(built).filter((file) => /\.c[jt]s$/.test(file))).toEqual([]);

    const bundle = readFileSync(join(built, 'mini-toolbar.mjs'), 'utf8');
    console.log(`dist/mini-toolbar.mjs: ${bundle.length} bytes`);
    const imports = specifiers(bundle);
    expect(imports.filter((specifier) => specifier.startsWith('@betteroffice/docx-react'))).toEqual([]);
    expect(
      imports.filter((specifier) => !chunkFile(built, specifier, false) && !['react', 'react/jsx-runtime'].includes(specifier))
    ).toEqual([]);
    expect(bundle.length).toBeLessThan(60_000);

    const editorModules = moduleGraph(built, 'index.mjs');
    const toolbarModules = moduleGraph(built, 'mini-toolbar.mjs');
    toolbarModules.delete('mini-toolbar.mjs');
    expect(toolbarModules.size).toBeGreaterThan(0);
    expect([...toolbarModules].filter((file) => !editorModules.has(file))).toEqual([]);
    for (const marker of ['createState must be a function', 'overlay must be a component']) {
      const holders = filesContaining(built, '.mjs', new RegExp(marker));
      expect(holders).toHaveLength(1);
      expect([toolbarModules.has(holders[0]!), editorModules.has(holders[0]!)]).toEqual([true, true]);
    }

    const types = readFileSync(join(built, 'mini-toolbar.d.mts'), 'utf8');
    expect(specifiers(types).filter((specifier) => specifier.startsWith('@betteroffice/docx-react'))).toEqual([]);
    expect(types).not.toMatch(/\bdocxPluginBrand\b|interface DocxPlugin\b/);
    const brands = filesContaining(built, '.d.mts', /declare const docxPluginBrand\b/);
    expect(brands).toHaveLength(1);
    expect(filesContaining(built, '.d.mts', /interface DocxPlugin\b/)).toEqual(brands);
    expect(moduleGraph(built, 'mini-toolbar.d.mts').has(brands[0]!)).toBe(true);
    expect(moduleGraph(built, 'index.d.mts').has(brands[0]!)).toBe(true);

    const consumerRoot = join(directory, 'consumer');
    const installed = join(consumerRoot, 'node_modules/@betteroffice/docx-react');
    mkdirSync(installed, { recursive: true });
    writeFileSync(join(installed, 'package.json'), manifestText);
    cpSync(built, join(installed, 'dist'), { recursive: true });
    writeFileSync(join(consumerRoot, 'package.json'), JSON.stringify({ type: 'module' }));
    const consumer = createRequire(join(consumerRoot, 'consumer.js'));
    expect(realpathSync(consumer.resolve('@betteroffice/docx-react/mini-toolbar'))).toBe(
      realpathSync(join(installed, 'dist/mini-toolbar.mjs'))
    );
    const subpath = manifest.exports['./mini-toolbar'] as { types: string };
    expect(existsSync(join(installed, subpath.types))).toBe(true);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}, 180_000);
