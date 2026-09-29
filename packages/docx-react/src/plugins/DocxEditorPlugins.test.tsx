import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { StrictMode, createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type DocxEditRequest } from '@betteroffice/docx/yrs';
import * as publicApi from '../index';
import {
  DocxEditor,
  DocxPluginToolbar,
  EditorToolbar,
  ToolbarCommandButton,
  defineDocxPlugin,
  useDocxCommand,
  useDocxCommands,
  type DocxCommandResult,
  type DocxEditorProps,
  type DocxEditorRef,
  type DocxPlugin,
  type DocxPluginCommandResult,
  type DocxPluginContext,
  type DocxPluginDefinition,
  type DocxPluginError,
  type DocxPluginEvent,
  type DocxPluginGeometry,
  type DocxAnchorGeometryResult,
  type DocxAnchorRect,
  type DocxGeometryTarget,
} from '../index';
import { isMacPlatform } from '../commands/descriptors';
import * as canvasReplay from '../components/DocxEditor/canvasReplay';

const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

const FIXTURE = resolve(
  import.meta.dir,
  '../components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx'
);
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

function documentBytes(): ArrayBuffer {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

const DRAWING =
  '<w:r><w:drawing><wp:inline><wp:extent cx="457200" cy="228600"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

async function inlineImageDocument(
  body = `<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Before </w:t></w:r>${DRAWING}<w:r><w:t xml:space="preserve"> and the text after it</w:t></w:r></w:p>`
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const office = 'application/vnd.openxmlformats-officedocument';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${office}.wordprocessingml.document.main+xml"/></Types>`
  );
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${rel}/image" Target="media/image1.png"/></Relationships>`
  );
  zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${rel}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function settle(ms = 20) {
  await act(async () => {
    await new Promise((done) => setTimeout(done, ms));
  });
}

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) await settle(10);
  expect(done()).toBe(true);
}

type State = { count: number };

function describeEvent(event: DocxPluginEvent): string | null {
  switch (event.type) {
    case 'load':
      return `load:${event.reason}`;
    case 'document-change':
      return `document-change:${event.version}`;
    case 'mode-change':
      return `mode-change:${event.mode}:${event.readOnly}`;
    case 'grants-change':
      return 'grants-change';
    default:
      return null;
  }
}

function recorder(id = 'acme.review', extra: Partial<DocxPluginDefinition<State>> = {}) {
  const log: string[] = [];
  const contexts: DocxPluginContext<State>[] = [];
  const plugin = defineDocxPlugin<State>({
    id,
    createState: () => ({ count: 0 }),
    initialize(context) {
      contexts.push(context);
      log.push('initialize');
      context.onCleanup((reason) => {
        log.push(`cleanup:${reason}`);
      });
    },
    onEvent(context, event) {
      contexts.push(context);
      const entry = describeEvent(event);
      if (entry) log.push(entry);
    },
    ...extra,
  });
  return { plugin, log, contexts };
}

async function mount(
  props: Partial<DocxEditorProps> = {},
  strict = false,
  buffer: ArrayBuffer = documentBytes()
) {
  const ref = createRef<DocxEditorRef>();
  const element = (next: Partial<DocxEditorProps>) => {
    const editor = <DocxEditor ref={ref} documentBuffer={buffer} {...next} />;
    return strict ? <StrictMode>{editor}</StrictMode> : editor;
  };
  const view = render(element(props));
  await until(() => ref.current?.commands.getState('save').enabled === true);
  return { ref, view, rerender: (next: Partial<DocxEditorProps>) => view.rerender(element(next)) };
}

async function firstParagraph(ref: React.RefObject<DocxEditorRef | null>) {
  const read = await ref.current!.readParagraphs({ view: 'accepted' });
  if (!read.ok) throw new Error(read.failure.message);
  const paragraph = read.paragraphs.find((candidate) => candidate.text.length >= 3)!;
  return { version: read.version, paragraph };
}

function appendRequest(version: string, paraId: string, text = '!'): DocxEditRequest {
  return {
    expectVersion: version,
    steps: [
      { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId }, at: 'end', text },
    ],
  };
}

async function selectFirstWord(ref: React.RefObject<DocxEditorRef | null>) {
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  const paragraph = session.paragraphs('body').find((candidate) => candidate.text.length >= 3)!;
  await act(async () => {
    session.setSelection(
      { story: 'body', paraId: paragraph.paraId, offset: 0 },
      { story: 'body', paraId: paragraph.paraId, offset: 3 }
    );
    editor.syncYrsInputState(false);
  });
}

const WRITE = { 'acme.review': { document: 'write', editBatches: true } } as const;

describe('DocxEditor plugins', () => {
  test('initialize, load, one change per batch, and cleanup on removal and readdition', async () => {
    const { plugin, log, contexts } = recorder();
    const { ref, rerender } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    expect(log.slice(0, 2)).toEqual(['initialize', 'load:loaded']);
    const loadVersion = contexts[1].snapshot.version;

    const { version, paragraph } = await firstParagraph(ref);
    expect(version).toBe(loadVersion);
    let applied!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
    await act(async () => {
      applied = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await settle(150);
    const changes = log.filter((entry) => entry.startsWith('document-change'));
    expect(changes).toEqual([`document-change:${applied.ok ? applied.version : ''}`]);

    await act(async () => {
      const stale = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
      expect(stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    });
    await settle(150);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(1);

    const session = ref.current!.getEditorRef()!.getYrsSession()!;
    for (const step of ['undo', 'redo'] as const) {
      await act(async () => {
        ref.current!.getEditorRef()![step]();
      });
      await settle(150);
      expect(log.filter((entry) => entry.startsWith('document-change')).at(-1)).toBe(
        `document-change:${session.version()}`
      );
    }
    const current = await firstParagraph(ref);
    await act(async () => {
      const noop = await ref.current!.applyEdits({
        expectVersion: current.version,
        steps: [
          {
            op: 'replaceText',
            target: { kind: 'paragraph', story: 'body', paraId: current.paragraph.paraId },
            text: current.paragraph.text,
          },
        ],
      });
      expect(noop).toMatchObject({ ok: true, applied: false });
    });
    await settle(150);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(3);
    const replica = await createYrsSession({ clientId: 5151 });
    try {
      replica.applyUpdate(session.encodeStateAsUpdate());
      const last = replica.paragraphs('body').at(-1)!;
      replica.insertText(
        { story: 'body', paraId: last.paraId, offset: last.text.length },
        ' remote'
      );
      await act(async () => {
        session.applyUpdate(replica.encodeStateAsUpdate(session.encodeStateVector()));
      });
      await settle(150);
    } finally {
      replica.destroy();
    }
    const versions = log.filter((entry) => entry.startsWith('document-change'));
    expect(versions).toHaveLength(4);
    expect(versions.at(-1)).toBe(`document-change:${session.version()}`);

    rerender({ plugins: [] });
    await settle();
    expect(log.filter((entry) => entry.startsWith('cleanup'))).toEqual(['cleanup:removed']);
    const removed = contexts[0];
    expect(removed.lifetimeSignal.aborted).toBe(true);
    expect(removed.setState({ count: 1 })).toBe(false);
    expect(await removed.read.version()).toMatchObject({
      ok: false,
      failure: { code: 'plugin-unavailable' },
    });

    rerender({ plugins: [plugin] });
    await until(() => log.includes('load:attached'));
  });

  test('a default plugin reads and navigates but cannot mutate through any path', async () => {
    const calls: string[] = [];
    function Probe() {
      const commands = useDocxCommands();
      const bold = commands.getState('bold');
      return (
        <>
          <button
            type="button"
            data-testid="probe"
            data-bold={bold.enabled ? 'enabled' : bold.disabledReason.code}
            onClick={() =>
              void commands.execute('bold', null).then((result) => {
                calls.push(result.ok ? 'ok' : result.failure.code);
              })
            }
          />
          <ToolbarCommandButton id="italic" />
        </>
      );
    }
    const { plugin, contexts, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'right', render: Probe },
    });
    const { ref, view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    await selectFirstWord(ref);
    await settle();
    const probe = within(view.container).getByTestId('probe');
    expect(probe.getAttribute('data-bold')).toBe('permission-denied');
    fireEvent.click(probe);
    await settle();
    expect(calls).toEqual(['permission-denied']);
    const italic = within(view.getByTestId('plugin-dock-right')).getByRole('button', {
      name: 'Italic',
    });
    expect(italic.getAttribute('aria-disabled')).toBe('true');
    expect(document.getElementById(italic.getAttribute('aria-describedby')!)?.textContent).toBe(
      'The host application has not allowed this plugin to use this command.'
    );
    fireEvent.click(italic);
    await settle();
    expect(ref.current!.commands.getState('italic').active).toBe(false);

    const context = contexts.at(-1)!;
    expect(context.edits).toBeNull();
    expect(await context.commands.execute('bold', null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    const read = await context.read.readParagraphs({ view: 'accepted' });
    expect(read.ok).toBe(true);
    if (!read.ok) return;
    const target = read.paragraphs.find((paragraph) => paragraph.text.length > 0)!;
    const validation = await context.read.validateEdits(appendRequest(read.version, target.paraId));
    expect(validation).toMatchObject({ ok: true, wouldApply: true });
    expect(
      await context.navigation.scrollToParagraph(
        { story: 'body', paraId: 'missing' },
        { expectVersion: read.version }
      )
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    expect(
      await context.navigation.scrollToParagraph(
        { story: 'body', paraId: target.paraId },
        { expectVersion: 'stale' }
      )
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  });

  test('granted built-in mutations still refuse without a policy path', async () => {
    function Probe() {
      const bold = useDocxCommands().getState('bold');
      return (
        <output data-testid="probe">{bold.enabled ? 'enabled' : bold.disabledReason.code}</output>
      );
    }
    const { plugin, contexts, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'left', render: Probe },
    });
    const { ref, view } = await mount({
      plugins: [plugin],
      pluginGrants: { 'acme.review': { document: 'write', commands: ['bold', 'reviewNext'] } },
    });
    await until(() => log.includes('load:loaded'));
    await selectFirstWord(ref);
    await settle();
    expect(within(view.container).getByTestId('probe').textContent).toBe('unsupported-policy');
    const context = contexts.at(-1)!;
    expect(await context.commands.execute('bold', null)).toMatchObject({
      ok: false,
      failure: { code: 'unsupported-policy' },
    });
    expect(await context.commands.execute('reviewNext', null)).toMatchObject({
      ok: false,
      failure: { code: 'no-revisions' },
    });
    expect(ref.current!.commands.getState('bold').enabled).toBe(true);
  });

  test('edit batches need their grant, respect viewing mode and revocation', async () => {
    const { plugin, contexts, log } = recorder();
    const { ref, rerender } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    await until(() => log.includes('load:loaded'));
    const context = contexts.at(-1)!;
    const edits = context.edits!;
    expect(edits).not.toBeNull();

    let { version, paragraph } = await firstParagraph(ref);
    const original = paragraph.text;
    const editor = ref.current!.getEditorRef()!;
    await act(async () => {
      const session = editor.getYrsSession()!;
      session.setSelection({ story: 'body', paraId: paragraph.paraId, offset: original.length });
      editor.syncYrsInputState(false);
      editor.insertText(' typed');
    });
    const typed = `${original} typed`;
    const read = await context.read.readParagraphs({
      view: 'accepted',
      paraIds: [paragraph.paraId],
    });
    expect(read.ok && read.paragraphs[0].text).toBe(typed);
    const target = { kind: 'paragraph', story: 'body', paraId: paragraph.paraId } as const;
    const applied = await act(() =>
      edits.applyEdits({
        expectVersion: read.ok ? read.version : '',
        steps: [
          { op: 'insertText', target, at: 'end', text: '!' },
          { op: 'insertText', target, at: 'start', text: '¡' },
        ],
      })
    );
    expect(applied).toMatchObject({ ok: true, applied: true });
    expect((await firstParagraph(ref)).paragraph.text).toBe(`¡${typed}!`);
    await act(async () => {
      expect(editor.undo()).toBe(true);
    });
    expect((await firstParagraph(ref)).paragraph.text).toBe(typed);
    await act(async () => {
      expect(editor.undo()).toBe(true);
    });
    expect((await firstParagraph(ref)).paragraph.text).toBe(original);
    await act(async () => {
      editor.redo();
      editor.redo();
    });
    ({ version, paragraph } = await firstParagraph(ref));
    expect(paragraph.text).toBe(`¡${typed}!`);
    expect(
      await edits.applyEdits({ ...appendRequest(version, paragraph.paraId), history: 'none' })
    ).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });

    rerender({ plugins: [plugin], pluginGrants: WRITE, mode: 'viewing' });
    await until(() => log.includes('mode-change:viewing:true'));
    expect(await edits.applyEdits(appendRequest(version, paragraph.paraId))).toMatchObject({
      ok: false,
      failure: { code: 'read-only' },
    });

    rerender({ plugins: [plugin], pluginGrants: {}, mode: 'editing' });
    await until(() => log.includes('grants-change'));
    expect(await edits.applyEdits(appendRequest(version, paragraph.paraId))).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    expect((await firstParagraph(ref)).version).toBe(version);

    const grant: { document: 'write'; editBatches?: true } = {
      document: 'write',
      editBatches: true,
    };
    const grants = { 'acme.review': grant };
    const changes = () => log.filter((entry) => entry === 'grants-change').length;
    rerender({ plugins: [plugin], pluginGrants: grants });
    await until(() => changes() === 2);
    delete grant.editBatches;
    rerender({ plugins: [plugin], pluginGrants: grants });
    await until(() => changes() === 3);
    expect(await edits.applyEdits(appendRequest(version, paragraph.paraId))).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
  });

  test('a load hook that applies a batch every run loads once and keeps its state', async () => {
    const loads: string[] = [];
    const plugin = defineDocxPlugin<State>({
      id: 'acme.review',
      createState: () => ({ count: 0 }),
      async onEvent(context, event) {
        if (event.type !== 'load') return;
        loads.push(event.version);
        const read = await context.read.readParagraphs({ view: 'accepted' });
        const paragraph = read.ok ? read.paragraphs.find((c) => c.text.length >= 3) : undefined;
        if (!read.ok || !paragraph || !context.edits) return;
        const applied = await context.edits.applyEdits(
          appendRequest(read.version, paragraph.paraId)
        );
        if (applied.ok) context.setState({ count: loads.length }, applied.version);
      },
      panel: {
        title: 'Loaded',
        placement: 'left',
        render: ({ context }) => <output data-testid="loaded">{context.state.count}</output>,
      },
    });
    const { ref, view } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    const body = within(view.container);
    await until(() => body.queryByTestId('loaded')?.textContent === '1');
    await settle(150);
    expect(loads).toHaveLength(1);
    expect((await firstParagraph(ref)).paragraph.text).toMatch(/[^!]!$/);
  });

  test('contributed commands run with their plugin, from the toolbar, ref and shortcuts', async () => {
    const calls: string[] = [];
    const errors: DocxPluginError[] = [];
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+M', 'Mod+B', 'Mod+V'],
          getState: (context) => ({ enabled: true, active: context.state.count > 0 }),
          execute(context) {
            calls.push(context.pluginId);
            context.setState((previous) => ({ count: previous.count + 1 }));
            return { ok: true, status: 'executed' };
          },
        },
      ],
      toolbar: ['mark'],
    });
    const other = recorder('other');
    const { ref, view } = await mount({
      plugins: [owner.plugin, other.plugin],
      onPluginError: (error) => errors.push(error),
    });
    await until(() => owner.log.includes('load:loaded') && other.log.includes('load:loaded'));
    await settle();
    const id = 'plugin:acme.review/mark' as const;
    expect(ref.current!.commands.getDescriptor(id)).toMatchObject({ label: 'Mark' });
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([
      ['acme.review', 'definition'],
      ['acme.review', 'definition'],
    ]);
    expect(String(errors[1].error)).toContain('already used by text editing');

    const button = within(view.container).getByRole('button', { name: 'Mark' });
    fireEvent.click(button);
    await settle();
    expect(calls).toEqual(['acme.review']);
    expect(ref.current!.commands.getState(id).active).toBe(true);

    expect(await act(() => ref.current!.commands.execute(id, null))).toMatchObject({ ok: true });
    const editor = view.container.querySelector('[data-testid="docx-editor"]')!;
    fireEvent.keyDown(editor, { key: 'M', shiftKey: true, ...MOD });
    await settle();
    expect(calls).toEqual(['acme.review', 'acme.review', 'acme.review']);

    expect(await other.contexts.at(-1)!.commands.execute(id, null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });

    ref.current!.setZoom(1);
    const custom = await mount({
      plugins: [owner.plugin],
      toolbar: (
        <EditorToolbar>
          <EditorToolbar.Toolbar>
            <DocxPluginToolbar />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      ),
    });
    await until(
      () => within(custom.view.container).queryByRole('button', { name: 'Mark' }) !== null
    );
  });

  test('contributed commands return their own failures and batch refusals unchanged', async () => {
    const returned: DocxPluginCommandResult[] = [];
    const paused = { code: 'paused', message: 'Paused' };
    let paraId = '';
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'quota',
          label: 'Quota',
          mutatesDocument: false,
          getState: (context) =>
            context.state.count > 0
              ? { enabled: false, disabledReason: paused }
              : { enabled: true },
          execute(context) {
            context.setState({ count: 1 });
            returned.push({ ok: false, failure: { code: 'quota-exceeded', message: 'Try later' } });
            return returned.at(-1)!;
          },
        },
        {
          id: 'stale',
          label: 'Stale',
          mutatesDocument: true,
          async execute(context) {
            const result = await context.edits!.applyEdits(appendRequest('stale', paraId));
            returned.push(result.ok ? { ok: true, status: 'executed' } : result);
            return returned.at(-1)!;
          },
        },
      ],
    });
    const { ref } = await mount({ plugins: [owner.plugin], pluginGrants: WRITE });
    await until(() => owner.log.includes('load:loaded'));
    const { version, paragraph } = await firstParagraph(ref);
    paraId = paragraph.paraId;
    const commands = ref.current!.commands;

    expect(await act(() => commands.execute('plugin:acme.review/quota', null))).toBe(returned[0]);
    expect(await act(() => commands.execute('plugin:acme.review/quota', null))).toEqual({
      ok: false,
      failure: paused,
    });
    const stale = await act(() => commands.execute('plugin:acme.review/stale', null));
    expect(stale).toBe(returned[1]);
    if (stale.ok || !('version' in stale)) throw new Error('expected a batch refusal');
    expect([stale.version, stale.failure.code]).toEqual([version, 'stale-version']);

    const builtIn: DocxCommandResult = await act(() => commands.execute('zoom', { scale: 1 }));
    expect(builtIn.ok).toBe(true);
  });

  test('a failing contribution is isolated and the others keep working', async () => {
    const errors: DocxPluginError[] = [];
    const broken = defineDocxPlugin({
      id: 'broken',
      createState: () => null,
      panel: {
        title: 'Broken',
        placement: 'left',
        render: () => {
          throw new Error('render failed');
        },
      },
      commands: [
        {
          id: 'go',
          label: 'Go',
          mutatesDocument: false,
          execute: () => ({ ok: true, status: 'executed' }),
        },
      ],
      toolbar: ['go'],
    });
    const cleanups: string[] = [];
    const brokenHook = defineDocxPlugin({
      id: 'throws',
      createState: () => null,
      initialize(context) {
        context.onCleanup((reason) => {
          cleanups.push(reason);
          throw new Error('cleanup failed');
        });
      },
      onEvent(_context, event) {
        if (event.type === 'load') throw new Error('load failed');
      },
    });
    const rejectsInit = defineDocxPlugin({
      id: 'rejects-init',
      createState: () => null,
      initialize: async () => {
        throw new Error('initialize rejected');
      },
      panel: { title: 'Never', placement: 'bottom', render: () => <p data-testid="never">no</p> },
    });
    const rejectsAction = defineDocxPlugin({
      id: 'rejects-action',
      createState: () => null,
      panel: {
        title: 'Action',
        placement: 'bottom',
        render: ({ context }) => (
          <button
            type="button"
            data-testid="failing-action"
            onClick={() =>
              void context.run(async () => {
                throw new Error('action rejected');
              })
            }
          />
        ),
      },
    });
    const healthy = recorder('healthy', {
      panel: { title: 'Healthy', placement: 'left', render: () => <p data-testid="healthy">ok</p> },
    });
    const { ref, view } = await mount({
      plugins: [broken, brokenHook, rejectsInit, rejectsAction, healthy.plugin],
      onPluginError: (error) => {
        errors.push(error);
        throw new Error('reporter failed');
      },
    });
    await until(() =>
      errors.some((error) => error.pluginId === 'broken' && error.phase === 'render')
    );
    await until(() =>
      errors.some((error) => error.pluginId === 'throws' && error.phase === 'cleanup')
    );
    expect(errors.find((error) => error.pluginId === 'throws')?.phase).toBe('event');
    expect(cleanups).toEqual(['failed']);
    await until(() =>
      errors.some((error) => error.pluginId === 'rejects-init' && error.phase === 'initialize')
    );
    expect(within(view.container).queryByTestId('never')).toBeNull();
    await until(() => within(view.container).queryByTestId('failing-action') !== null);
    fireEvent.click(within(view.container).getByTestId('failing-action'));
    await until(() =>
      errors.some((error) => error.pluginId === 'rejects-action' && error.phase === 'action')
    );
    await settle();
    expect(within(view.container).queryByTestId('failing-action')).toBeNull();
    expect(within(view.container).getByTestId('healthy').textContent).toBe('ok');
    expect(ref.current!.commands.getDescriptor('plugin:broken/go')).toBeNull();
    expect(ref.current!.commands.getState('save').enabled).toBe(true);
  });

  test('replacing the document ends the activation and loads a fresh one', async () => {
    const { plugin, log, contexts } = recorder();
    const { ref } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const first = contexts[0];
    await act(async () => {
      await ref.current!.loadDocumentBuffer(documentBytes());
    });
    await until(() => log.includes('load:replaced'));
    expect(log.indexOf('cleanup:document-replaced')).toBeLessThan(log.indexOf('load:replaced'));
    expect(first.lifetimeSignal.aborted).toBe(true);
    expect(await first.read.version()).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    const fresh = contexts.at(-1)!;
    expect(fresh.snapshot.generation).not.toBe(first.snapshot.generation);
  });

  test('loading a parsed document activates plugins once, over the new session', async () => {
    const { plugin, log } = recorder();
    const { ref } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    await act(async () => {
      ref.current!.loadDocument(ref.current!.getDocument()!);
    });
    await until(() => log.includes('load:replaced'));
    await settle(150);
    expect(log.slice(log.indexOf('load:loaded') + 1)).toEqual([
      'cleanup:document-replaced',
      'initialize',
      'load:replaced',
    ]);
  });

  test('StrictMode setup, cleanup and setup leaves one live activation', async () => {
    const { plugin, log, contexts } = recorder();
    await mount({ plugins: [plugin] }, true);
    await until(() => log.includes('load:loaded'));
    await settle();
    const initialized = log.filter((entry) => entry === 'initialize').length;
    const cleaned = log.filter((entry) => entry.startsWith('cleanup')).length;
    expect(initialized - cleaned).toBe(1);
    const live = new Set(
      contexts
        .filter((context) => !context.lifetimeSignal.aborted)
        .map((context) => context.lifetimeSignal)
    );
    expect(live.size).toBe(1);
  });

  test('unmounting the editor cleans up with reason unmounted', async () => {
    const { plugin, log, contexts } = recorder();
    const { view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    view.unmount();
    await settle();
    expect(log.filter((entry) => entry.startsWith('cleanup'))).toEqual(['cleanup:unmounted']);
    expect(contexts[0].lifetimeSignal.aborted).toBe(true);
  });

  test('overlays follow the rendered layout version; navigation reports missing geometry', async () => {
    const layouts: (string | null)[] = [];
    const contexts: DocxPluginContext<null>[] = [];
    const geometries: DocxPluginGeometry[] = [];
    const plugin = defineDocxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      onEvent(context, event) {
        contexts.push(context);
        if (event.type === 'layout-change') layouts.push(event.layout?.version ?? null);
      },
      overlay: ({ context, geometry }) => {
        geometries.push(geometry);
        return (
          <div
            data-testid="layout-marker"
            data-version={geometry.layout.version}
            data-snapshot={context.snapshot.version}
          />
        );
      },
    });
    const { ref, view } = await mount({ plugins: [plugin] });
    const marker = () => view.container.querySelector<HTMLElement>('[data-testid="layout-marker"]');
    await until(() => marker() !== null);
    const { version, paragraph } = await firstParagraph(ref);
    expect(marker()!.dataset).toMatchObject({ version, snapshot: version });
    const unit = { x: 0, y: 0, width: 1, height: 1 };
    const retained = geometries.at(-1)!;
    expect(retained.toOverlayRect(unit)).not.toBeNull();

    let next = '';
    await act(async () => {
      const applied = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
      if (applied.ok) next = applied.version;
    });
    await until(() => marker()?.dataset.version === next);
    const afterEdit = layouts.slice(layouts.lastIndexOf(version) + 1);
    expect(afterEdit[0]).toBeNull();
    expect(afterEdit.at(-1)).toBe(next);
    expect(retained.toOverlayRect(unit)).toBeNull();
    expect(geometries.at(-1)!.toOverlayRect(unit)).not.toBeNull();

    // happy-dom lays out no pixels, so the pages have no client geometry to scroll to.
    expect(
      await contexts
        .at(-1)!
        .navigation.scrollToParagraph(
          { story: 'body', paraId: paragraph.paraId },
          { expectVersion: next }
        )
    ).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
  });

  test('an overlay resolves paragraph and search anchors at the rendered version', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.anchors',
      createState: () => null,
      overlay: (props) => {
        geometry = props.geometry;
        return <div data-testid="anchor-overlay" />;
      },
    });
    const { ref } = await mount({ plugins: [plugin] });
    await until(() => geometry !== null);
    const session = ref.current!.getEditorRef()!.getYrsSession()!;
    await act(async () => {
      expect(session.persistParagraphIds().status).toBe('applied');
      ref.current!.getEditorRef()!.syncYrsInputState(true);
    });
    const { version, paragraph } = await firstParagraph(ref);
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === version);
    const entry = session
      .paragraphIdentities()
      .paragraphs.find(
        (candidate) =>
          candidate.session?.story === paragraph.story &&
          candidate.session.paraId === paragraph.paraId
      )!;
    const identity = entry.session!;
    expect(entry.persisted).not.toBeNull();
    const current = geometry! as DocxPluginGeometry;
    current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
    for (const canvas of current.dom.pagesContainer.querySelectorAll('canvas[data-page-index]')) {
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
    }
    const targets = [
      { kind: 'paragraph' as const, paragraph: identity },
      { kind: 'paragraph' as const, paragraph: entry.persisted! },
      { kind: 'search' as const, paragraph: entry.persisted!, text: paragraph.text.slice(0, 3) },
    ];
    for (const target of targets) {
      const result = current.getAnchorGeometry(target);
      expect(result).toMatchObject({
        ok: true,
        version,
        previewVersion: 0,
        layoutId: current.layout.id,
      });
      if (!result.ok) throw new Error(result.failure.message);
      expect(result.rects.length).toBeGreaterThan(0);
      expect(result.anchor.width).toBe(0);
    }
    expect(
      current.getAnchorGeometry({
        kind: 'range',
        version: `${version}-stale`,
        range: {
          story: paragraph.story,
          start: { paraId: paragraph.paraId, offset: 0 },
          end: { paraId: paragraph.paraId, offset: 3 },
          view: 'accepted',
        },
      })
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  });

  test('overlays and layout events re-anchor once the pages show a new zoom', async () => {
    const results: DocxAnchorGeometryResult[] = [];
    const events: (DocxAnchorGeometryResult | null)[] = [];
    let target: DocxGeometryTarget | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.zoom-anchor',
      createState: () => null,
      onEvent(context, event) {
        if (event.type === 'layout-change' && event.layout?.zoom === 1.5 && target) {
          events.push(context.geometry?.getAnchorGeometry(target) ?? null);
        }
      },
      overlay: ({ geometry }) => {
        if (target) results.push(geometry.getAnchorGeometry(target));
        return null;
      },
    });
    const { ref } = await mount({ plugins: [plugin] });
    const { paragraph } = await firstParagraph(ref);
    const anchor = ref
      .current!.getEditorRef()!
      .getYrsSession()!
      .paragraphIdentities()
      .paragraphs.find((entry) => entry.session?.paraId === paragraph.paraId)!.session!;
    target = { kind: 'paragraph', paragraph: anchor };
    const rect = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue(
      new DOMRect(0, 0, 800, 1000)
    );
    const present = canvasReplay.presentCanvasReplay;
    const paints: (() => void)[] = [];
    const held = spyOn(canvasReplay, 'presentCanvasReplay').mockImplementation(
      async (preparations, isCurrent) => {
        await new Promise<void>((resolve) => paints.push(resolve));
        return present(preparations, isCurrent);
      }
    );
    try {
      await act(async () => ref.current!.setZoom(1.5));
      await until(() => events.length > 0 && paints.length > 0);
      expect(events).toMatchObject([{ ok: false, failure: { code: 'layout-unavailable' } }]);
      expect(results.at(-1)).toMatchObject({ ok: false });
      await act(async () => {
        for (const paint of paints.splice(0)) paint();
      });
      held.mockRestore();
      await until(() => {
        const last = results.at(-1);
        return !!last?.ok && last.rects.length > 0;
      });
      await until(() => events.at(-1)?.ok === true);
    } finally {
      held.mockRestore();
      rect.mockRestore();
    }
  });

  test('an anchor ends at the last unit of its range, not at an inline image drawn after it', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.image-anchor',
      createState: () => null,
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const { ref } = await mount({ plugins: [plugin] }, false, await inlineImageDocument());
    await until(() => geometry !== null);
    const read = await ref.current!.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    const paragraph = read.paragraphs.find((candidate) => candidate.text.includes('\uFFFC'))!;
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === read.version);
    const current = geometry! as DocxPluginGeometry;
    current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
    for (const canvas of current.dom.pagesContainer.querySelectorAll('canvas[data-page-index]')) {
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
    }
    const result = current.getAnchorGeometry({
      kind: 'range',
      version: read.version,
      range: {
        story: paragraph.story,
        start: { paraId: paragraph.paraId, offset: paragraph.text.indexOf('\uFFFC') },
        end: { paraId: paragraph.paraId, offset: paragraph.text.length },
        view: 'accepted',
      },
    });
    if (!result.ok) throw new Error(result.failure.message);
    const ends = result.rects.map((rect) => rect.x + rect.width);
    expect(result.rects.length).toBeGreaterThan(1);
    expect(ends.at(-1)).toBeLessThan(Math.max(...ends));
    expect(result.anchor.x).toBeCloseTo(Math.max(...ends));
    const atom = paragraph.text.indexOf('\uFFFC');
    const image = current.getAnchorGeometry({
      kind: 'range',
      version: read.version,
      range: {
        story: paragraph.story,
        start: { paraId: paragraph.paraId, offset: atom },
        end: { paraId: paragraph.paraId, offset: atom + 1 },
        view: 'accepted',
      },
    });
    if (!image.ok) throw new Error(image.failure.message);
    expect(image.rects).toHaveLength(1);
    expect(image.anchor.x).toBeCloseTo(image.rects[0]!.x + image.rects[0]!.width);
  });

  test('an anchor stays inside its range across a line break and after an image', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.break-anchor',
      createState: () => null,
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const body = [
      '<w:p w14:paraId="00000001"><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>',
      `<w:p w14:paraId="00000002">${DRAWING}</w:p>`,
    ].join('');
    const { ref } = await mount({ plugins: [plugin] }, false, await inlineImageDocument(body));
    await until(() => geometry !== null);
    const read = await ref.current!.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === read.version);
    const current = geometry! as DocxPluginGeometry;
    current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
    for (const canvas of current.dom.pagesContainer.querySelectorAll('canvas[data-page-index]')) {
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
    }
    const anchor = (paraId: string, start: number, end: number) => {
      const result = current.getAnchorGeometry({
        kind: 'range',
        version: read.version,
        range: {
          story: 'body',
          start: { paraId, offset: start },
          end: { paraId, offset: end },
          view: 'accepted',
        },
      });
      if (!result.ok) throw new Error(result.failure.message);
      return result;
    };
    const [broken, picture] = read.paragraphs;
    const alpha = anchor(broken!.paraId, 0, 6);
    const line = alpha.rects.at(-1)!;
    expect(alpha.anchor.y).toBeCloseTo(line.y);
    expect(alpha.anchor.x).toBeCloseTo(line.x + line.width);
    const image = anchor(picture!.paraId, 0, 1).rects.reduce((wide, rect) =>
      rect.width > wide.width ? rect : wide
    );
    expect(anchor(picture!.paraId, 0, 1).anchor.x).toBeCloseTo(image.x + image.width);
    expect(anchor(picture!.paraId, 1, 1).anchor.x).toBeCloseTo(image.x + image.width);
    expect(anchor(picture!.paraId, 0, 0).anchor.x).toBeCloseTo(image.x);
  });

  test('anchors a table cell, an empty paragraph and a paragraph across two pages', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.page-anchor',
      createState: () => null,
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const body = [
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>',
      '<w:p w14:paraId="00000011"><w:r><w:t>In the cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      '<w:p w14:paraId="00000012"/>',
      `<w:p w14:paraId="00000013"><w:r>${'<w:t>line</w:t><w:br/>'.repeat(80)}<w:t>end</w:t></w:r></w:p>`,
    ].join('');
    const { ref } = await mount({ plugins: [plugin] }, false, await inlineImageDocument(body));
    await until(() => geometry !== null);
    const session = ref.current!.getEditorRef()!.getYrsSession()!;
    const version = session.version();
    await until(() => {
      const layout = (geometry as DocxPluginGeometry | null)?.layout;
      return layout?.version === version && layout.pageCount > 1;
    });
    const current = geometry! as DocxPluginGeometry;
    const PAGE_GAP = 1200;
    current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 900, 4000);
    for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>(
      'canvas[data-page-index]'
    )) {
      const page = Number(canvas.dataset.pageIndex);
      canvas.getBoundingClientRect = () => new DOMRect(0, page * PAGE_GAP, 816, 1056);
    }
    const byWordId = (ooxmlParaId: string) =>
      session
        .paragraphIdentities()
        .paragraphs.find((entry) => entry.ooxmlParaId === ooxmlParaId)!.session!;
    const at = (ooxmlParaId: string) => {
      const result = current.getAnchorGeometry({
        kind: 'paragraph',
        paragraph: byWordId(ooxmlParaId),
      });
      if (!result.ok) throw new Error(result.failure.message);
      return result;
    };
    const inside = (anchor: DocxAnchorRect, page: { x: number; y: number; height: number }) =>
      anchor.y >= page.y && anchor.y + anchor.height <= page.y + page.height;

    const cell = at('00000011');
    expect(byWordId('00000011').story).toStartWith('body:');
    expect(cell.rects.length).toBeGreaterThan(0);
    expect(cell.rects.every((rect) => rect.pageIndex === 0)).toBe(true);
    const cellEnd = cell.rects.at(-1)!;
    expect(cell.anchor).toMatchObject({ pageIndex: 0, width: 0 });
    expect(cell.anchor.x).toBeCloseTo(cellEnd.x + cellEnd.width);
    expect(inside(cell.anchor, cell.pageRect)).toBe(true);

    const empty = at('00000012');
    expect(empty.rects).toEqual([]);
    expect(empty.anchor.width).toBe(0);
    expect(empty.anchor.y).toBeGreaterThan(cell.anchor.y);

    const spanning = at('00000013');
    const pages = new Set(spanning.rects.map((rect) => rect.pageIndex));
    expect(pages.size).toBeGreaterThan(1);
    const last = spanning.rects.at(-1)!;
    expect(spanning.anchor.pageIndex).toBe(Math.max(...pages));
    expect(spanning.anchor.pageIndex).toBe(last.pageIndex);
    expect(spanning.pageRect.y).toBeCloseTo(spanning.anchor.pageIndex * PAGE_GAP);
    expect(inside(spanning.anchor, spanning.pageRect)).toBe(true);
    for (const rect of spanning.rects) {
      expect(rect.y).toBeGreaterThanOrEqual(rect.pageIndex * PAGE_GAP);
      expect(rect.y).toBeLessThan(rect.pageIndex * PAGE_GAP + 1056);
    }
  });

  test('an anchor ends where the text of its last unit ends, left in right-to-left runs', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin({
      id: 'acme.bidi-anchor',
      createState: () => null,
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const rtl = (text: string) => `<w:r><w:rPr><w:rtl/></w:rPr><w:t>${text}</w:t></w:r>`;
    const ltr = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body = [
      `<w:p w14:paraId="00000001"><w:pPr><w:bidi/></w:pPr>${rtl('\u05D0\u05D1\u05D2')}</w:p>`,
      `<w:p w14:paraId="00000002">${ltr('abc ')}${rtl('\u05D0\u05D1\u05D2')}${ltr(' def')}</w:p>`,
      `<w:p w14:paraId="00000003"><w:pPr><w:bidi/></w:pPr>${rtl('\u05D0\u05D1\u05D2')}<w:r><w:rPr><w:vanish/></w:rPr><w:t>${'x'.repeat(70)}</w:t></w:r></w:p>`,
    ].join('');
    const { ref } = await mount({ plugins: [plugin] }, false, await inlineImageDocument(body));
    await until(() => geometry !== null);
    const read = await ref.current!.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === read.version);
    const current = geometry! as DocxPluginGeometry;
    current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
    for (const canvas of current.dom.pagesContainer.querySelectorAll('canvas[data-page-index]')) {
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
    }
    const anchor = (paraId: string, start: number, end: number) => {
      const result = current.getAnchorGeometry({
        kind: 'range',
        version: read.version,
        range: {
          story: 'body',
          start: { paraId, offset: start },
          end: { paraId, offset: end },
          view: 'accepted',
        },
      });
      if (!result.ok) throw new Error(result.failure.message);
      return result;
    };
    const [hebrew, mixed, suffixed] = read.paragraphs;
    for (const [paragraph, start] of [
      [hebrew!, 0],
      [mixed!, 4],
    ] as const) {
      const last = anchor(paragraph.paraId, start + 2, start + 3);
      expect(last.rects).toHaveLength(1);
      expect(last.anchor.x).toBeCloseTo(last.rects[0]!.x);
      const word = anchor(paragraph.paraId, start, start + 3);
      expect(word.anchor.x).toBeCloseTo(Math.min(...word.rects.map((rect) => rect.x)));
    }
    const english = anchor(mixed!.paraId, 0, 3);
    expect(english.anchor.x).toBeCloseTo(english.rects[0]!.x + english.rects[0]!.width);
    const word = anchor(hebrew!.paraId, 0, 3).rects[0]!;
    expect(anchor(hebrew!.paraId, 0, 0).anchor.x).toBeCloseTo(word.x + word.width);
    expect(anchor(hebrew!.paraId, 3, 3).anchor.x).toBeCloseTo(word.x);
    const hidden = anchor(suffixed!.paraId, 0, suffixed!.text.length);
    expect(hidden.anchor.x).toBeCloseTo(Math.min(...hidden.rects.map((rect) => rect.x)));
  });

  test('public presenters and hooks bind contributed commands', async () => {
    const calls: string[] = [];
    function Bound() {
      const mark = useDocxCommand('plugin:acme.review/mark');
      return (
        <output data-testid="bound">{`${mark.label}:${mark.state.enabled}:${mark.shortcut}`}</output>
      );
    }
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+M'],
          execute(context) {
            calls.push(context.pluginId);
            return { ok: true, status: 'executed' };
          },
        },
      ],
    });
    const { view } = await mount({
      plugins: [owner.plugin],
      toolbar: (
        <EditorToolbar>
          <EditorToolbar.Toolbar>
            <ToolbarCommandButton id="plugin:acme.review/mark" />
            <ToolbarCommandButton id="plugin:acme.review/absent" label="Absent" />
            <Bound />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      ),
    });
    const body = within(view.container);
    await until(() => body.queryByRole('button', { name: 'Mark' }) !== null);
    expect(body.queryByRole('button', { name: 'Absent' })).toBeNull();
    expect(body.getByTestId('bound').textContent).toMatch(/^Mark:true:.+M$/);
    fireEvent.click(body.getByRole('button', { name: 'Mark' }));
    await settle();
    expect(calls).toEqual(['acme.review']);
  });

  test('installing plugins materializes no compatibility document', async () => {
    const yrs = await import('@betteroffice/docx/yrs');
    const materialized = spyOn(yrs, 'yrsToDocument');
    const edit = async (props: Partial<DocxEditorProps>) => {
      const { ref } = await mount(props);
      await settle(300);
      materialized.mockClear();
      const { version, paragraph } = await firstParagraph(ref);
      await act(async () => {
        await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
      });
      await settle(300);
      const count = materialized.mock.calls.length;
      cleanup();
      return count;
    };
    try {
      const plain = await edit({});
      const reads: string[] = [];
      const reader = defineDocxPlugin({
        id: 'reader',
        createState: () => null,
        async onEvent(context, event) {
          if (event.type !== 'document-change') return;
          const read = await context.read.readParagraphs({ view: 'accepted' });
          if (read.ok) reads.push(read.version);
        },
      });
      expect(await edit({ plugins: [reader] })).toBe(plain);
      expect(reads).toHaveLength(1);
      expect(await edit({ onChange: () => {} })).toBeGreaterThan(plain);
    } finally {
      materialized.mockRestore();
    }
  });

  test('the demo review plugin works through the documented exports', async () => {
    mock.module('@betteroffice/docx-react', () => publicApi);
    const demo = '../../../../apps/demo/app/docx/ReviewPlugin';
    const { reviewPlugin } = (await import(demo)) as { reviewPlugin: DocxPlugin };
    const { ref, view, rerender } = await mount({ plugins: [reviewPlugin] });
    const body = within(view.container);
    const action = () => body.getByRole('button', { name: 'Mark first paragraph reviewed' });
    await until(() => body.queryByText(/paragraphs, version/) !== null);
    expect(action().hasAttribute('disabled')).toBe(true);
    const toolbarMark = body.getByRole('button', { name: 'Mark reviewed' });
    expect(toolbarMark.getAttribute('aria-disabled')).toBe('true');

    const grants = { 'demo.review': { document: 'write', editBatches: true } } as const;
    rerender({ plugins: [reviewPlugin], pluginGrants: grants });
    await until(() => !action().hasAttribute('disabled'));
    const before = (await firstParagraph(ref)).paragraph;
    fireEvent.click(action());
    await until(() => body.queryByText('Marked the first paragraph as reviewed.') !== null);
    const read = await ref.current!.readParagraphs({ view: 'accepted', paraIds: [before.paraId] });
    expect(read.ok && read.paragraphs[0].text).toBe(`${before.text} ✓`);

    rerender({ plugins: [reviewPlugin], pluginGrants: grants, mode: 'viewing' });
    await until(() => action().hasAttribute('disabled'));
    await until(
      () =>
        body.getByRole('button', { name: 'Mark reviewed' }).getAttribute('aria-disabled') === 'true'
    );

    rerender({ plugins: [], pluginGrants: grants });
    await until(
      () => body.queryByRole('button', { name: 'Mark first paragraph reviewed' }) === null
    );
  });

  test('plugins that are not defined through defineDocxPlugin are reported, not run', async () => {
    const errors: DocxPluginError[] = [];
    await mount({
      plugins: [{ id: 'raw' } as unknown as DocxPlugin],
      onPluginError: (error) => errors.push(error),
    });
    await settle();
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([['raw', 'definition']]);
  });
});
